from typing import Dict, List, Optional
from dataclasses import dataclass
from wpsdoc import Document
import json
import re
import io
WORD_LENGTH_THRESHOLD = 20
@dataclass
class Suggestion:
    operation: str  # "替换" or "提醒"
    after: str     # 替换后的内容，如果operation为"提醒"则为空

@dataclass
class Issue:
    issueType: str
    specificWord: str
    sentence: str
    suggestion: Suggestion
    rule_id: str
    additionalNotes: str

@dataclass
class DocumentReviewResult:
    issues: List[Issue]

    def to_dict(self) -> Dict:
        return {
            "issues": [
                {
                    "issueType": issue.issueType,
                    "specificWord": issue.specificWord,
                    "sentence": issue.sentence,
                    "suggestion": {
                        "operation": issue.suggestion.operation,
                        "after": issue.suggestion.after
                    },
                    "rule_id": issue.rule_id,
                    "additionalNotes": issue.additionalNotes
                }
                for issue in self.issues
            ]
        }

def load_rules(rules_file: str) -> Dict[str, Dict]:
    """加载规则配置文件"""
    rules_data = {}
    with open(rules_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        header_skipped = False
        for line in lines:
            if not header_skipped:
                if '---' in line: # Detect separator line after header
                    header_skipped = True
                continue
            if line.strip().startswith('|') and line.strip().endswith('|'):
                parts = [p.strip() for p in line.split('|')[1:-1]] # Remove leading/trailing empty strings from split
                if len(parts) == 5:
                    rule_id = parts[0]
                    rules_data[rule_id] = {
                        'id': rule_id,
                        'type_scene': parts[1],
                        'description': parts[2],
                        'example': parts[3],
                        'operation_suggestion': parts[4]
                    } 
    return rules_data

def validate_document(docx_file, rules_path: str) -> DocumentReviewResult:
    """验证文档并返回问题列表，docx_file 可为文件路径或 file-like 对象"""

    if hasattr(docx_file, 'read'):
        file_bytes = docx_file.read()
        doc = Document(io.BytesIO(file_bytes))
    else:
        doc = Document(docx_file)
    
    rules = load_rules(rules_path)
    issues: List[Issue] = []

    # Chinese numeral conversion helpers (ensure these are correctly defined and comprehensive)
    def simple_chinese_to_int(s: str) -> Optional[int]:
        mapping = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
                   '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15, '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十':20}
        # Extend this mapping as needed for numbers like 二十一, 三十, etc.
        return mapping.get(s)

    def int_to_simple_chinese(n: int) -> Optional[str]:
        mapping = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九', 10: '十',
                   11: '十一', 12: '十二', 13: '十三', 14: '十四', 15: '十五', 16: '十六', 17: '十七', 18: '十八', 19: '十九', 20:'二十'}
        # Extend this mapping as needed
        return mapping.get(n)

    heading_configs = [
        {
            'level': 0,
            'pattern': re.compile(r"^爱你在心口.*?"),
            'jump_order_rule_id': '',
            'font_style_rule_id': '15-02',
            'font_rule': ['仿宋', 'FangSong', 'FangSong_GB2312'],
            'bold': False,
            'numeral_type': 'arabic'
        },
        {
            'level': 1,
            'pattern': re.compile(r"^([一二三四五六七八九十]+(?:[一二三四五六七八九])?)、"),
            'jump_order_rule_id': '05-05', # Corresponds to '标序问题-一级标题跳序问题' (from user's original code for L1)
                                         # If rules_p1.md has '05-05' for L1, update this ID.
            'font_style_rule_id': '06-05',
            'font_rule': ['黑体', 'SimHei'],
            'bold': True,
            'numeral_type': 'chinese'
        },
        {
            'level': 2,
            'pattern': re.compile(r"^[（(]([一二三四五六七八九十]+(?:[一二三四五六七八九])?)[)）]"),
            'jump_order_rule_id': '05-03', # '标序问题-二级标题跳序问题'
            'font_style_rule_id': '06-03',
            'font_rule': ['楷体', 'KaiTi', 'STKaiti'],
            'bold': False,
            'numeral_type': 'chinese'
        },
        {
            'level': 3,
            'pattern': re.compile(r"^(\d+)\."),
            'jump_order_rule_id': '05-04', # '标序问题-三级标题跳序问题'
            'font_style_rule_id': '06-04',
            'font_rule': ['仿宋', 'FangSong', 'FangSong_GB2312'],
            'bold': False,
            'numeral_type': 'arabic'
        },
        {
            'level': 4,
            'pattern': re.compile(r"^[（(](\d+)[)）]"),
            'jump_order_rule_id': '05-02', # Assuming a rule like '05-05' exists for L4 jump order based on previous structure.
                                         # Please verify/add this rule to rules_p1.md if it's different or missing.
                                         # The original code had a '05-05' for L4 in the `heading_configs` example.
            'font_style_rule_id': '06-02',
            'font_rule': ['仿宋', 'FangSong', 'FangSong_GB2312'],
            'bold': False,
            'numeral_type': 'arabic'
        }
    ]

    collected_headings_by_level: List[List[Dict[str, Any]]] = [[] for _ in heading_configs]

    # Standard two-character indent (value from original code)
    TWO_CHAR_INDENT_EMU = 2 
    INDENT_TOLERANCE = 0

    def get_dominant_font_properties(paragraph) -> tuple[Optional[str], Optional[bool]]:
        run_details = []
        total_text_len = 0
        for run in paragraph.runs:
            run_text = run.text
            if run_text.strip():
                run_len = len(run_text)
                run_details.append({
                    'len': run_len,
                    'font_name': run.font.name,
                    'bold': run.font.bold,
                })
                total_text_len += run_len
        
        if not run_details or total_text_len == 0: return None, None

        font_counts = {}
        for rd in run_details:
            key = (rd['font_name'], rd['bold'])
            font_counts[key] = font_counts.get(key, 0) + rd['len']
        
        if not font_counts: return None, None
        
        dominant_font_key = max(font_counts, key=font_counts.get)
        return dominant_font_key[0], dominant_font_key[1]

    def check_font_style(p_obj, p_text_content: str, rule_id: str, expected_fonts: List[str], expected_bold: Optional[bool]):
        actual_font_name, actual_bold = get_dominant_font_properties(p_obj)
        rule_info = rules.get(rule_id, {})
        
        font_ok = False
        if actual_font_name:
            for ef in expected_fonts:
                if ef.lower() in actual_font_name.lower(): # Case-insensitive check for font name
                    font_ok = True
                    break
        # If actual_font_name is None, font_ok remains False
        
        bold_ok = True 
        if expected_bold is True and not actual_bold:
            bold_ok = False
        elif expected_bold is False and actual_bold is True:
            bold_ok = False
        
        # SimHei/黑体 is often inherently bold or its 'bold' flag is set, so if expected bold, it's ok.
        if actual_font_name and actual_font_name.lower() in ["simhei", "黑体"] and expected_bold is True:
             bold_ok = True 

        if not font_ok or not bold_ok:
            notes = []
            if not font_ok:
                notes.append(f"字体应为 '{'/'.join(expected_fonts)}' 系列, 实际主要字体为 '{actual_font_name if actual_font_name else '未知'}'")
            if not bold_ok:
                expected_bold_str = "加粗" if expected_bold is True else ("不加粗" if expected_bold is False else "未指定")
                actual_bold_str = "加粗" if actual_bold is True else ("不加粗" if actual_bold is False else "未明确")
                notes.append(f"字体应 {expected_bold_str}, 实际 {actual_bold_str}")
            
            heading_prefix = p_text_content# Get ALL part as potential heading text
            if len(heading_prefix) > 20: heading_prefix = heading_prefix[:WORD_LENGTH_THRESHOLD] + "..." # Truncate if too long

            issues.append(Issue(
                issueType=rule_info.get('type_scene', f"格式问题-{rule_id}"),
                specificWord=heading_prefix,
                sentence=p_text_content,
                suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
                rule_id=rule_id,
                additionalNotes=f"{rule_info.get('description', '标题格式问题')}: {'; '.join(notes)}"
            ))
        FONT_SIZE=16
        notes = []
        crt_size=""
        # 15-02  
        rule_info = rules.get("15-02", {})
        if p_obj.size != FONT_SIZE and p_obj.size is not None:
            # crt_size=p_obj.font.size 
            notes.append(f"字号大小应为{FONT_SIZE}磅，实际为{p_obj.size}磅")
            issues.append(Issue(
                issueType=rule_info.get('type_scene', f"正文字号应为16磅"),
                specificWord=p_obj.text[:WORD_LENGTH_THRESHOLD],
                sentence=p_obj.text,
                suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
                rule_id=rule_id,
                additionalNotes=f"{rule_info.get('description', '字号错误')}: {'; '.join(notes)}"
            ))

        elif p_obj.size is None:
            for run in p_obj.runs:
                if run.font.size != FONT_SIZE:
                    notes.append(f"字号大小应为{FONT_SIZE}磅，实际为{run.font.size}磅")
                    issues.append(Issue(
                    issueType=rule_info.get('type_scene', f"正文字号应为16磅"),
                    specificWord=run.text[:WORD_LENGTH_THRESHOLD],
                    sentence=run.text,
                    suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
                    rule_id=rule_id,
                    additionalNotes=f"{rule_info.get('description', '字号错误')}: {'; '.join(notes)}"
                ))

    # Main loop to process paragraphs
    for p_idx, p in enumerate(doc.paragraphs):
        with open("debug.csv", "a") as f:
            f.write(f"{p_idx},{p.info}\n")  # Debug print from original, commented out
           
        text_content = p.text # Use raw text for matching, strip later if needed for specificWord
        # html_content = p.html # If needed
        # first_line_indent = p.paragraph_format.first_line_indent # If needed
        # print(f"{p_idx+1}: {first_line_indent},{text_content[:50]}\n") # Debug print from original, commented out

        if not text_content.strip(): # Skip empty or whitespace-only paragraphs
            continue

        is_any_known_heading = False
        for level_idx, config in enumerate(heading_configs):
            match = config['pattern'].match(text_content)
            if match and level_idx!=0:
                is_any_known_heading = True
                numeral_part_str = match.group(1)
                full_heading_prefix = match.group(0)
                
                collected_headings_by_level[level_idx].append({
                    'num_str': numeral_part_str,
                    'text': text_content.strip(), # Store stripped text for sentence context
                    'paragraph_obj': p,
                    'prefix': full_heading_prefix
                })
                
                # Check font style for this heading
                check_font_style(p, text_content.strip(), 
                                 config['font_style_rule_id'], 
                                 config['font_rule'], 
                                 config['bold'])
                break # Paragraph is classified as one heading type, move to next paragraph

        # Paragraph indentation and hanging indent checks (Rules 14-02, 14-03)
        if not is_any_known_heading: 
            stripped_text_content = text_content.strip()
            # Rule 14-02: 段落-自然段左空两字
            first_line_indent = p.paragraph_format.first_line_indent
            if first_line_indent is None or first_line_indent != TWO_CHAR_INDENT_EMU: 
                rule_info = rules.get("14-02", {})
                issues.append(Issue(
                    issueType=rule_info.get('type_scene', "段落-自然段左空两字"),
                    specificWord=stripped_text_content[:WORD_LENGTH_THRESHOLD],
                    sentence=stripped_text_content,
                    suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
                    rule_id="14-02",
                    additionalNotes=rule_info.get('description', "段落首行应当左空二字")
                ))
            
            # Rule 14-03: 段落- 回行顶格
            # left_indent = p.paragraph_format.left_indent
            # if left_indent is not None and left_indent > INDENT_TOLERANCE: 
            #     rule_info = rules.get("14-03", {})
            #     issues.append(Issue(
            #         issueType=rule_info.get('type_scene', "段落- 回行顶格"),
            #         specificWord=stripped_text_content[:WORD_LENGTH_THRESHOLD],
            #         sentence=stripped_text_content,
            #         suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
            #         rule_id="14-03",
            #         additionalNotes=rule_info.get('description', "段落回行应顶格（段落整体左侧不应有额外缩进）")
            #     ))

    # Check for jump-order issues in collected headings
    for level_idx, headings_at_this_level in enumerate(collected_headings_by_level):
        config = heading_configs[level_idx]
        numeral_type = config['numeral_type']
        jump_order_rule_id = config['jump_order_rule_id']

        for i in range(len(headings_at_this_level) - 1):
            current_heading = headings_at_this_level[i]
            next_heading = headings_at_this_level[i+1]

            current_val: Optional[int] = None
            next_actual_val: Optional[int] = None
            expected_next_num_str_func = str # Default for arabic if not chinese

            if numeral_type == 'chinese':
                current_val = simple_chinese_to_int(current_heading['num_str'])
                next_actual_val = simple_chinese_to_int(next_heading['num_str'])
                expected_next_num_str_func = int_to_simple_chinese
            elif numeral_type == 'arabic':
                try:
                    current_val = int(current_heading['num_str'])
                    next_actual_val = int(next_heading['num_str'])
                except ValueError:
                    # Log or handle non-integer numerals if they are not expected
                    # print(f"Warning: Could not convert numeral '{current_heading['num_str']}' or '{next_heading['num_str']}' to int for level {config['level']}")
                    continue 
            
            if current_val is not None and next_actual_val is not None:
                if next_actual_val != current_val + 1:
                    rule_info = rules.get(jump_order_rule_id, {})
                    expected_num_val = current_val + 1
                    expected_num_val_str = expected_next_num_str_func(expected_num_val)
                    if expected_num_val_str is None and numeral_type == 'chinese': # Fallback if int_to_simple_chinese returns None
                        expected_num_val_str = str(expected_num_val) 
                    elif expected_num_val_str is None: # Should not happen for arabic if conversion was successful
                         expected_num_val_str = "?"

                    issues.append(Issue(
                        issueType=rule_info.get('type_scene', f"标序问题-跳序问题 (L{config['level']})"),
                        specificWord=next_heading['prefix'], 
                        sentence=next_heading['text'],
                        suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
                        rule_id=jump_order_rule_id,
                        additionalNotes=rule_info.get('description', 
                            f"标题序号不连续。在 '{current_heading['prefix']}' 之后，期望序号为 '{expected_num_val_str}'，实际为 '{next_heading['num_str']}'.")
                    ))
            # else: Consider logging if numeral conversion failed for a detected heading

    return DocumentReviewResult(issues=issues)

def validate_and_output_json(docx_file, rules_path: str) -> dict:
    """验证文档并输出JSON结果，适配 Flask 文件对象输入，直接返回 dict"""
    # docx_file 可以是文件路径或 file-like 对象
    # 直接将 docx_file 传递给 validate_document，由它处理路径或文件对象
    result = validate_document(docx_file, rules_path)
    result_dict = result.to_dict()
    for issue in result_dict.get('issues', []):
        if 'suggestion' in issue and 'after' not in issue['suggestion']:
            issue['suggestion']['after'] = ""
    output_path = "./validation_result.json"
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result_dict, f, ensure_ascii=False, indent=4)
            # f.write(str(result_dict).replace("'", '"'))
            print(f"校验完成，结果已保存到 {output_path}")
    # 
    return result_dict
    
    # The following lines for writing to output_path and returning json_str are now effectively dead code
    # as the function returns result_dict earlier. This might need further review based on requirements.
    # json_str = json.dumps(result_dict, ensure_ascii=False, indent=4)
    # 
    # if output_path:
    #     with open(output_path, 'w', encoding='utf-8') as f:
    #         f.write(json_str)
    #         print(f"校验完成，结果已保存到 {output_path}")
    # 
    # return json_str

if __name__ == "__main__":
    # 使用示例文件进行测试
    docx_file="./p1.txt"
    # docx_file="./转完p1.docx"
    rules_file = "./rules_p1.md"
    output_file = "./validation_result.json"
    
    print(f"开始校验文档: {docx_file}")
    print(f"使用规则文件: {rules_file}")
    
    # 执行校验并输出到文件
    validate_and_output_json(docx_file, rules_file)
