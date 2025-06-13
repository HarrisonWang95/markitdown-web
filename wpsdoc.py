import re
from typing import IO

class Run:
    def __init__(self, text,html, font_name, bold, font_size):
        self.text = text.replace("&nbsp;", " ")
        self.html = html
        self.font = Font(font_name, bold, font_size)


class Font:
    def __init__(self, name, bold, size):
        self.name = name
        self.bold = bold
        self.size = size


class ParagraphFormat:
    def __init__(self):
        self.first_line_indent = None
        self.left_indent = None


class Paragraph:
    def __init__(self, runs):
        self.runs = runs
        self.type = ""
        self.text = ''.join([run.text for run in runs])
        self.html = ''.join([run.html for run in runs])
        self.paragraph_format = ParagraphFormat()
        self._set_font_properties()
        self.info=f"{self.font},{self.bold},{self.size},{self.paragraph_format.first_line_indent},{self.paragraph_format.left_indent},{self.text}"
        

    def _set_font_properties(self):
        font_names = set([run.font.name for run in self.runs])
        font_sizes = set([run.font.size for run in self.runs])
        bold_values = set([run.font.bold for run in self.runs])

        self.font = font_names.pop() if len(font_names) == 1 else None
        self.bold = bold_values.pop() if len(bold_values) == 1 else None
        self.size = font_sizes.pop() if len(font_sizes) == 1 else None
    def set_type(self, type: str, level: int = None):
        """设置段落类型
        :param type: 段落类型（body/heading）
        :param level: 标题等级（1-6），仅当type=heading时有效
        """
        if type == 'body':
            self.type = 'body'
        elif type == 'heading':
            if not 1 <= level <= 10:
                raise ValueError("标题等级必须在1-10之间")
            self.type = f'heading{level}'
        else:
            raise ValueError("无效的段落类型")
        
        # 保留原有样式属性
        self._keep_original_style()


class DocumentObject:
    def __init__(self, html_string):
        # 提取<body></body>里的内容
        pattern = r'<body.*?>(.*?)</body>'
        match = re.search(pattern, html_string, re.DOTALL)
        if match:
            body_content = match.group(1)
        else:
            body_content = ""
        # 提取段落信息，假设段落以<p class=MsoNormal>标签分隔
        paragraph_pattern = r'<p class=.*?>.*?</p>'
        paragraph_matches = re.findall(paragraph_pattern, body_content, re.DOTALL)
        self.paragraphs = []
        for para_match in paragraph_matches:
            run_pattern = r'<span.*?>.*?</span>'
            run_matches = re.findall(run_pattern, para_match, re.DOTALL)
            # 去除 HTML 标签，只保留文本内容
            runs=[]
            for run_match in run_matches:
                clean_text = re.sub(r'<[^>]+>', '', run_match)
                # 简单假设字体信息从<p>标签的 style 属性中提取
                font_name_pattern = r'font-family:([^;]+);'
                font_name_match = re.search(font_name_pattern, run_match)
                font_name = font_name_match.group(1) if font_name_match else None
                font_size_pattern = r'font-size:([^;]+)pt;'
                font_size_match = re.search(font_size_pattern, run_match)
                font_size = int(float(font_size_match.group(1))) if font_size_match else None
                bold = None  # 这里简单假设无法直接从示例中提取 bold 信息
                run = Run(clean_text,run_match, font_name, bold, font_size)
                runs.append(run)
            paragraph = Paragraph(runs)
            # 提取首行缩进信息
            first_line_indent_pattern = r'mso-char-indent-count:([^;]+);'
            first_line_indent_match = re.search(first_line_indent_pattern , para_match)
            if first_line_indent_match:
                paragraph.paragraph_format.first_line_indent = int(float(first_line_indent_match.group(1)))
            
            #回行顶格不需要校验
            # left_indent_pattern = r'mso-list:l1 level1 lfo1; margin-left:([^;]+);'
            # left_indent_match = re.search(left_indent_pattern, para_match)
            # if left_indent_match:
            #     paragraph.paragraph_format.left_indent = int(float(left_indent_match.group(1)))

            if(paragraph.text!="&nbsp;" and paragraph.text != " "):
                self.paragraphs.append(paragraph)

def Document(html: str | IO[bytes] | None = None) -> DocumentObject:
    if isinstance(html, str):
        # 处理文件路径
        with open(html, 'r') as f:
            html_content = f.read()
    elif hasattr(html, 'read'):
        # 处理文件对象
        html_content = html.read()
        # 如果读取到的是字节，则需要解码
        if isinstance(html_content, bytes):
            html_content = html_content.decode('utf-8')  # 使用适当的编码进行解码
    else:
        raise ValueError("Invalid input type")
    
    return DocumentObject(html_content)

if  __name__ == '__main__':
    # 读取 p1_caped_html.txt 中的 HTML 字符串
    # with open('p1_caped_html.txt', 'r', encoding='utf-8') as file:
    #     html_string = file.read()

    # 创建 Document 对象
    doc = Document('p1_caped_html.txt')
    # print(len(html_string))

    # 示例：打印文档中的段落文本
    for para in doc.paragraphs:
        print(para.text)

