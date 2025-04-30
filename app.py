import os
import uuid
import io
import requests
from concurrent.futures import ThreadPoolExecutor
from flask import Flask, request, jsonify
from flask_restful import Api, Resource
from werkzeug.utils import secure_filename
from werkzeug.exceptions import BadRequest, InternalServerError, RequestURITooLarge
from markitdown import MarkItDown
from pdfminer.pdfparser import PDFSyntaxError # 用于捕获PDF页数解析错误
from pdfminer.pdfdocument import PDFTextExtractionNotAllowed # 用于捕获PDF页数解析错误
from pathlib import Path
import warnings
import time
from time import sleep

# 过滤特定消息模式的警告
warnings.filterwarnings('ignore', message="Cannot set gray non-stroke color.*", module='pdfminer')
# 尝试导入 openai，如果需要 LLM 功能
try:
    from openai import AzureOpenAI
except ImportError:
    AzureOpenAI = None
from contextlib import contextmanager
from enum import Enum

# --- 配置 ---
MAX_FILE_SIZE = 500 * 1024 * 1024  # 100 MB
MAX_PDF_PAGES = 500
SUPPORTED_MIMETYPES = [
    # 文档格式
    'application/pdf',  # PDF
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',  # Word
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',  # Excel
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',  # PowerPoint
    'application/msword',  # Word doc
    
    # 图片格式
    'image/jpeg',
    'image/png',
    'image/gif',
    'image/webp',
    'image/tiff',
    
    # 音频格式
    'audio/mpeg',
    'audio/wav',
    'audio/ogg',
    'audio/webm',
    
    # 文本格式
    'text/plain',
    'text/html',
    'text/csv',
    'application/json',
    'application/xml',
    'text/xml',
    'text/markdown',
    
    # 电子书
    'application/epub+zip',
    
    # 压缩文件
    'application/zip',
    'application/x-zip-compressed',
    
    #eml
    'message/rfc822',
]
MAX_WORKERS = 8 # 最大并发线程数


# 删除这行，因为已经重复定义了
# UPLOAD_FOLDER = '/tmp/markitdown_uploads' # 临时文件存储目录

# 使用相对路径，创建 uploads 目录在当前文件夹下
UPLOAD_FOLDER = Path(__file__).parent / 'uploads'
if not UPLOAD_FOLDER.exists():
    UPLOAD_FOLDER.mkdir(parents=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE
api = Api(app)

# --- 异步任务处理 ---
executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)
from dataclasses import dataclass
# import threading

@dataclass
class TaskStatus:
    status: str
    result: str = None
    error: str = None
    metadata: dict = None
    timestamp: float = None  # 新增字段，记录任务创建时间戳

tasks = {}  # {task_id: TaskStatus}
TASK_EXPIRE_SECONDS = 7200  # 2小时

def set_task(task_id, task_status):
    tasks[task_id] = task_status

def get_task(task_id):
    now = time.time()
    task_status=tasks.get(task_id)
    # 清理过期任务
    expired = [k for k, v in tasks.items() if v.timestamp and now - v.timestamp > TASK_EXPIRE_SECONDS]
    for k in expired:
        del tasks[k]
    return task_status

# def cleanup_tasks():
#     while True:
#         now = time.time()
#         expired = [k for k, v in tasks.items() if v.timestamp and now - v.timestamp > TASK_EXPIRE_SECONDS]
#         for k in expired:
#             del tasks[k]
#         time.sleep(600)  # 每10分钟清理一次

# threading.Thread(target=cleanup_tasks, daemon=True).start()

# --- MarkItDown 实例 ---
# 根据请求参数动态创建 MarkItDown 实例
from typing import Dict, Optional, Union

def get_markitdown_instance(args: Dict[str, str]) -> MarkItDown:
    enable_plugins = args.get('enable_plugins', 'false').lower() == 'true'
    use_docintel = args.get('use_docintel', 'false').lower() == 'true'
    docintel_endpoint = args.get('docintel_endpoint')
    use_llm = args.get('use_llm', 'false').lower() == 'true'
    llm_model = args.get('llm_model', 'gpt-4o') # 默认模型

    md_kwargs = {'enable_plugins': enable_plugins}

    if use_docintel and docintel_endpoint:
        md_kwargs['docintel_endpoint'] = docintel_endpoint
        # 注意：可能需要配置 Azure 凭证，这里假设使用默认凭证链
        # from azure.identity import DefaultAzureCredential
        # md_kwargs['docintel_credential'] = DefaultAzureCredential()

    if use_llm and AzureOpenAI:
        # 注意：需要配置 OpenAI API Key，通常通过环境变量 OPENAI_API_KEY
        try:
            # Azure OpenAI 配置从环境变量读取
            openai_api_key = os.environ.get("AZURE_OPENAI_API_KEY")
            openai_api_base = os.environ.get("AZURE_OPENAI_ENDPOINT")
            openai_api_version = os.environ.get("AZURE_OPENAI_API_VERSION", "2025-02-01-preview")
            openai_deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")

            if not openai_api_key or not openai_api_base:
                raise BadRequest("缺少 Azure OpenAI 配置，请设置环境变量。")

            client = AzureOpenAI(
                api_key=openai_api_key,
                azure_endpoint=openai_api_base,
                api_version=openai_api_version
            )
            md_kwargs['llm_client'] = client
            md_kwargs['llm_model'] = llm_model
        except Exception as e:
            app.logger.warning(f"无法初始化 OpenAI 客户端: {e}. LLM 功能将不可用。")
            # 可以选择在这里报错或仅禁用 LLM
            pass # 或者 raise BadRequest("无法初始化 LLM 客户端，请检查 API Key")

    return MarkItDown(**md_kwargs)

# --- 文件处理和解析任务 ---
def get_pdf_page_count(file_stream):
    """尝试获取 PDF 页数，失败则返回 -1"""
    try:
        from pdfminer.pdfparser import PDFParser
        from pdfminer.pdfdocument import PDFDocument
        parser = PDFParser(file_stream)
        document = PDFDocument(parser)
        return document.catalog.get('Pages').resolve().get('Count', -1)
    except (PDFSyntaxError, PDFTextExtractionNotAllowed, Exception) as e:
        # 捕获 pdfminer 可能的错误以及其他潜在问题
        app.logger.error(f"获取 PDF 页数时出错: {e}")
        return -1 # 表示无法确定页数或文件无效

def process_file(task_id, file_path_or_url, is_url, original_filename, content_type, args):
    """后台处理文件解析的任务"""
    tasks[task_id]['status'] = 'processing'
    file_stream = None
    temp_file_path = None
    downloaded = False

    try:
        md = get_markitdown_instance(args)
        metadata = tasks[task_id]['metadata']

        if is_url:
            # --- URL 处理 ---
            response = requests.get(file_path_or_url, stream=True, timeout=30) # 增加超时
            response.raise_for_status() # 检查请求是否成功

            # 校验 Content-Type (如果服务器提供了)
            url_content_type = response.headers.get('content-type', '').split(';')[0]
            if url_content_type and url_content_type not in SUPPORTED_MIMETYPES:
                 raise BadRequest(f"不支持的文件类型: {url_content_type}")

            # 校验大小
            content_length = response.headers.get('content-length')
            if content_length and int(content_length) > MAX_FILE_SIZE:
                raise RequestURITooLarge("通过 URL 下载的文件超过大小限制")

            # 下载到临时文件以传递给 markitdown
            temp_file_path = os.path.join(UPLOAD_FOLDER, f"{task_id}_{secure_filename(original_filename or 'downloaded_file')}")
            file_size = 0
            with open(temp_file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    file_size += len(chunk)
                    if file_size > MAX_FILE_SIZE:
                        raise RequestURITooLarge("通过 URL 下载的文件超过大小限制")
                    f.write(chunk)
            downloaded = True
            metadata['size'] = file_size
            metadata['mime_type'] = url_content_type or 'unknown' # 更新 MIME 类型
            file_to_process = temp_file_path
            processing_input = temp_file_path # Markitdown 可能需要路径

        else:
            # --- 文件流处理 ---
            file_to_process = file_path_or_url # 这是初始保存的路径
            metadata['size'] = os.path.getsize(file_to_process)
            metadata['mime_type'] = content_type
            processing_input = file_to_process # Markitdown 可能需要路径

        # --- 通用校验和处理 ---
        # PDF 页数校验 (仅对 PDF)
        if metadata.get('mime_type') == 'application/pdf':
            page_count = -1
            try:
                with open(file_to_process, 'rb') as f:
                    page_count = get_pdf_page_count(f)
            except Exception as e:
                 app.logger.error(f"读取文件进行页数检查时出错 ({file_to_process}): {e}")
                 raise InternalServerError("读取文件时出错")

            if page_count == -1:
                 app.logger.warning(f"无法确定 PDF 页数: {metadata.get('name')}")
                 # 根据需求决定是否严格失败，这里选择警告并继续
                 metadata['pages'] = 'Unknown'
            elif page_count > MAX_PDF_PAGES:
                raise BadRequest(f"文件页数超过限制 ({page_count}/{MAX_PDF_PAGES})")
            else:
                metadata['pages'] = page_count

        # --- 调用 MarkItDown ---
        # MarkItDown.convert 现在接受文件路径或流。如果需要流，需要打开文件。
        # 查阅 markitdown 文档确认 convert 的最新用法。假设它能处理路径。
        # 如果 convert 需要流:
        # with open(processing_input, 'rb') as f:
        #    result = md.convert(f)
        # 假设 convert 可以接受路径:
        result = md.convert(processing_input)

        tasks[task_id]['status'] = 'success'
        tasks[task_id]['result'] = result.text_content
        # 可以考虑也存储 result.metadata

    except requests.exceptions.RequestException as e:
        app.logger.error(f"下载 URL 时出错 ({file_path_or_url}): {e}")
        tasks[task_id]['status'] = 'error'
        tasks[task_id]['error'] = f"下载文件失败: {e}"
    except BadRequest as e:
        tasks[task_id]['status'] = 'error'
        tasks[task_id]['error'] = str(e.description)
    except RequestURITooLarge as e:
        tasks[task_id]['status'] = 'error'
        tasks[task_id]['error'] = str(e.description)
    # except ConversionError as e:
    #     app.logger.error(f"MarkItDown 解析失败 ({original_filename}): {e}")
    #     tasks[task_id]['status'] = 'error'
    #     tasks[task_id]['error'] = f"文件解析失败: {e}"
    except Exception as e:
        app.logger.exception(f"处理任务 {task_id} 时发生未知错误") # 使用 exception 记录堆栈
        tasks[task_id]['status'] = 'error'
        tasks[task_id]['error'] = f"内部服务器错误: {e}"
    finally:
        # 清理临时文件
        if downloaded and temp_file_path and os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
            except OSError as e:
                app.logger.error(f"无法删除临时下载文件 {temp_file_path}: {e}")
        # 如果上传的文件也需要删除（取决于策略）
        if not is_url and file_path_or_url and os.path.exists(file_path_or_url):
             try:
                 os.remove(file_path_or_url)
                 app.logger.info(f"删除临时上传文件: {file_path_or_url}")
             except OSError as e:
                 app.logger.error(f"无法删除临时上传文件 {file_path_or_url}: {e}")


# --- API 资源 ---
class UploadResource(Resource):
    def post(self):
        task_id = str(uuid.uuid4())
        file_path = None
        is_url = False
        original_filename = None
        content_type = None
        args = request.args.to_dict() # 获取查询参数 ?enable_plugins=true&...

        try:
            if 'file' in request.files:
                # --- 文件流上传 ---
                file = request.files['file']
                if file.filename == '':
                    raise BadRequest("未选择文件")

                original_filename = secure_filename(file.filename)
                content_type = file.content_type

                if content_type not in SUPPORTED_MIMETYPES:
                    raise BadRequest(f"不支持的文件类型: {content_type}")

                # 保存文件到临时位置
                file_path = os.path.join(UPLOAD_FOLDER, f"{task_id}_{original_filename}")
                file.save(file_path)

                # 立即检查文件大小 (虽然 Flask 的 MAX_CONTENT_LENGTH 应该已经处理了)
                file_size = os.path.getsize(file_path)
                if file_size > MAX_FILE_SIZE:
                    # 理论上不会到这里，除非 MAX_CONTENT_LENGTH 配置失效或文件在保存后变大
                    raise RequestURITooLarge("文件超过大小限制")

                app.logger.info(f"接收到文件: {original_filename}, 类型: {content_type}, 大小: {file_size}, 任务ID: {task_id}")


            elif request.is_json and 'url' in request.json:
                # --- URL 上传 ---
                is_url = True
                file_path_or_url = request.json['url']
                if not file_path_or_url.startswith(('http://', 'https://')):
                     raise BadRequest("无效的 URL 格式")
                original_filename = file_path_or_url.split('/')[-1] # 尝试从 URL 获取文件名
                # content_type 和 size 将在下载时确定
                content_type = 'unknown'
                file_size = 'unknown'
                file_path = file_path_or_url # 传递 URL 给处理函数
                app.logger.info(f"接收到 URL: {file_path_or_url}, 任务ID: {task_id}")

            else:
                raise BadRequest("请求必须包含 'file' (multipart/form-data) 或 'url' (json)")

            # 初始化任务状态
            tasks[task_id] = {
                "status": "pending",
                "result": None,
                "error": None,
                "metadata": {
                    "name": original_filename,
                    "size": file_size if not is_url else 'pending download',
                    "mime_type": content_type,
                    "pages": "" # 稍后更新
                }
            }

            # 提交后台处理
            executor.submit(process_file, task_id, file_path, is_url, original_filename, content_type, args)

            return jsonify({
                "code": 200,
                "message": "文件上传成功，正在处理中",
                "data": {
                    "task_id": task_id,
                    "status": "pending",
                    "metadata": tasks[task_id]['metadata']
                }
            })

        except BadRequest as e:
             # 如果在提交任务前发生错误，需要清理已保存的文件
             if file_path and os.path.exists(file_path) and not is_url:
                 try:
                     os.remove(file_path)
                 except OSError as rm_err:
                     app.logger.error(f"无法删除部分上传的文件 {file_path}: {rm_err}")
             return {"code": 400, "message": str(e.description), "data": None}, 400
        except RequestURITooLarge as e:
             if file_path and os.path.exists(file_path) and not is_url:
                 try:
                     os.remove(file_path)
                 except OSError as rm_err:
                     app.logger.error(f"无法删除过大的文件 {file_path}: {rm_err}")
             return {"code": 413, "message": str(e.description), "data": None}, 413
        except Exception as e:
            app.logger.exception(f"上传处理中发生未知错误")
            if file_path and os.path.exists(file_path) and not is_url:
                 try:
                     os.remove(file_path)
                 except OSError as rm_err:
                     app.logger.error(f"错误处理中无法删除文件 {file_path}: {rm_err}")
            return {"code": 500, "message": "内部服务器错误", "data": None}, 500


class ParseStatusResource(Resource):
    # print(tasks)
    def get(self, task_id):
        task = get_task(task_id) #tasks.get(task_id)
        if not task:
            return {"code": 404, "message": "未找到任务", "data": None}, 404

        response_data = {
            "task_id": task_id,
            "status": task['status'],
            "metadata": task.get('metadata', {}),
            "content": task.get('result'),
            "error": task.get('error')
        }

        return jsonify({
            "code": 200,
            "message": "查询成功",
            "data": response_data
        })

class UploadSyncResource(Resource):
    def post(self):
        # 获取初始响应
        # print(Resource.__dict__)
        upload_response = UploadResource().post()
        # print(upload_response)
        if isinstance(upload_response, tuple):
            return upload_response  # 如果是错误响应，直接返回
        
        # 从响应中提取任务ID
        task_id = upload_response.json['data']['task_id']
        start_time = time.time()
        timeout = 180  # 3分钟超时
        
        while True:
            # 检查是否超时
            if time.time() - start_time > timeout:
                return jsonify({
                    "code": 408,
                    "message": "处理超时",
                    "data": {
                        "task_id": task_id,
                        "status": "timeout"
                    }
                }), 408
            
            # 获取任务状态
            status_response = ParseStatusResource().get(task_id)
            if isinstance(status_response, tuple):
                return status_response
            
            status_data = status_response.json['data']
            
            if status_data['status'] in ['success', 'error']:
                return status_response
            
            # 使用非阻塞的 sleep
            sleep(1)
# --- 路由 ---
api.add_resource(UploadResource, '/api/v1/upload')
api.add_resource(ParseStatusResource, '/api/v1/parse/<string:task_id>')
api.add_resource(UploadSyncResource, '/api/v1/upload/parse')

# --- 错误处理 ---
@app.errorhandler(400)
def handle_bad_request(e):
    return jsonify({"code": 400, "message": f"错误的请求: {e.description}", "data": None}), 400

@app.errorhandler(404)
def handle_not_found(e):
    # Flask-RESTful 通常会处理 404，但可以保留以防万一
    return jsonify({"code": 404, "message": "资源未找到", "data": None}), 404

@app.errorhandler(413)
def handle_payload_too_large(e):
    return jsonify({"code": 413, "message": f"请求体过大: {e.description}", "data": None}), 413

@app.errorhandler(500)
@app.errorhandler(InternalServerError) # 捕获 werkzeug 的 InternalServerError
def handle_internal_error(e):
    # 对于 werkzeug 异常，原始异常在 e.original_exception
    error_message = str(e.description) if hasattr(e, 'description') else str(e)
    app.logger.error(f"内部服务器错误: {error_message}") # 记录错误
    return jsonify({"code": 500, "message": f"内部服务器错误: {error_message}", "data": None}), 500




@contextmanager
def temp_file(file_path):
    try:
        yield file_path
    finally:
        if file_path.exists():
            file_path.unlink()




if __name__ == '__main__':
    # 仅用于本地开发调试
    # 生产环境应使用 Gunicorn
    app.run(debug=false, host='0.0.0.0', port=5050)


