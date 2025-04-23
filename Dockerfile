# 使用与 markitdown 兼容的 Python 版本 (根据需求文档是 3.13)
FROM python:3.13-slim-bullseye 

# 设置工作目录
WORKDIR /app

# 安装系统依赖 (如果 markitdown[all] 需要)
# 例如，pdfminer.six 可能需要字体或库，音频处理需要 ffmpeg
# 查阅 markitdown 及其依赖项的文档确定所需系统包
# RUN apt-get update && apt-get install -y --no-install-recommends \
#     ffmpeg \
#     # 其他可能的依赖，如 build-essential, libpoppler-cpp-dev 等
#     && rm -rf /var/lib/apt/lists/*
# 暂时注释掉，如果遇到问题再取消注释并添加具体包

# 复制依赖文件
COPY requirements.txt .

# 安装 Python 依赖
# 注意：如果 markitdown 是本地包，需要先复制整个项目或 markitdown 包到镜像中
# 并修改 requirements.txt 或使用 pip install -e ./packages/markitdown[all]
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用代码
COPY . /app

# 创建临时上传目录并设置权限 (Gunicorn 默认可能以非 root 用户运行)
RUN mkdir -p /tmp/markitdown_uploads && chown -R nobody:nogroup /tmp/markitdown_uploads
# 如果 Gunicorn 以其他用户运行，需要相应调整

# 暴露端口
EXPOSE 5000

# 设置环境变量 (可选)
# ENV OPENAI_API_KEY="your_openai_key" # 不推荐硬编码，最好在运行时注入

# 运行 Gunicorn
# 使用 4 workers, 每个 worker 2 threads (总共 8 并发处理)
# 增加超时时间以处理可能较长的解析任务
CMD ["gunicorn", "--workers", "4", "--threads", "2", "--bind", "0.0.0.0:5000", "--timeout", "120", "app:app"]