from typing import Dict, List, Optional
from dataclasses import dataclass
from wpsdoc import Document
import json
import re
import io

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
    # Load 
    rules = load_rules(rules_path)
    issues: List[Issue] = []

    # Chinese numeral conversion helpers
    def simple_chinese_to_int(s: str) -> Optional[int]:
        mapping = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
                   '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15, '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十':20}
        # Add more if needed, e.g., 二十一, etc.
        return mapping.get(s)

    def int_to_simple_chinese(n: int) -> Optional[str]:
        mapping = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九', 10: '十',
                   11: '十一', 12: '十二', 13: '十三', 14: '十四', 15: '十五', 16: '十六', 17: '十七', 18: '十八', 19: '十九', 20:'二十'}
        # Add more if needed
        return mapping.get(n)

    # Rule 05-02: 标序问题-跳序问题 (一级标题 "一、")
    level1_headings_text_parts = [] # Stores the numeral part like "一", "二"
    level1_headings_full_text = [] # Stores full paragraph text like "一、XXXX"
    
    # Regex for "一、", "二、", ..., "十、", "十一、", etc. (simplified for now)
    level1_pattern = re.compile(r"^([一二三四五六七八九十]+(?:[一二三四五六七八九])?)、")
    i=0
    for p in doc.paragraphs:
        text_content = ''  # 初始化文本内容
        # for run in p.runs:
        #     text_content += run.text  # 拼接所有run的文本
        text_content = p.text#.strip()
        html_content=p.html
        first_line_indent = p.paragraph_format.first_line_indent
        i+=1
        # with open("debug.csv", "a") as f:
        print(f"{i}: {first_line_indent},{html_content},{text_content}\n")
            
        if not text_content: continue
        match = level1_pattern.match(text_content)
        if match:
            level1_headings_text_parts.append(match.group(1)) 
            level1_headings_full_text.append(text_content)

    for i in range(len(level1_headings_text_parts) - 1):
        current_num_str = level1_headings_text_parts[i]
        next_actual_num_str = level1_headings_text_parts[i+1]
        current_p_full_text = level1_headings_full_text[i+1]

        current_val = simple_chinese_to_int(current_num_str)
        next_actual_val = simple_chinese_to_int(next_actual_num_str)

        if current_val is not None and next_actual_val is not None:
            if next_actual_val != current_val + 1:
                rule_info = rules.get("05-02", {})
                expected_next_str = int_to_simple_chinese(current_val + 1)
                issues.append(Issue(
                    issueType=rule_info.get('type_scene', "标序问题-跳序问题"),
                    specificWord=f"{next_actual_num_str}、",
                    sentence=current_p_full_text,
                    suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
                    rule_id="05-02",
                    additionalNotes=rule_info.get('description', f"标题序号不连续，期望是 {expected_next_str}、，实际是 {next_actual_num_str}、")
                ))
        # else: Consider logging a warning if conversion fails for a detected heading

    # Standard two-character indent in EMUs (approx 0.85cm, common in Word for '2 char' indent)
    # 1 inch = 914400 EMUs. 2 chars (e.g. 12pt SongTi, each char is 12pt wide) ~ 24pt.
    # 24pt * 12700 EMU/pt = 304800 EMU.
    TWO_CHAR_INDENT_EMU = 2 
    INDENT_TOLERANCE = 0# A tolerance for indent comparison (approx 0.5mm)

    for p_idx, p in enumerate(doc.paragraphs):
        text_content = p.text.strip()
        if not text_content:
            continue

        # Define heading patterns - ensure they are mutually exclusive or checked in order
        is_heading_level1 = bool(level1_pattern.match(text_content))
        
        level2_pattern_text = r"([一二三四五六七八九十]+(?:[一二三四五六七八九])?)"
        level2_pattern_full = re.compile(f"^[（(]{level2_pattern_text}[)）]")
        is_heading_level2 = bool(level2_pattern_full.match(text_content))
        
        level3_pattern_text = r"(\d+)"
        level3_pattern = re.compile(f"^{level3_pattern_text}\.")
        is_heading_level3 = bool(level3_pattern.match(text_content))
        
        level4_pattern_text = r"(\d+)"
        level4_pattern_full = re.compile(f"^[（(]{level4_pattern_text}[)）]")
        is_heading_level4 = bool(level4_pattern_full.match(text_content))
        
        is_any_known_heading = is_heading_level1 or is_heading_level2 or is_heading_level3 or is_heading_level4

        # Rule 14-02: 段落-自然段左空两字
        if not is_any_known_heading: 
            first_line_indent = p.paragraph_format.first_line_indent
            # Check if first_line_indent is None or less than expected (with tolerance)
            if first_line_indent is None or first_line_indent != TWO_CHAR_INDENT_EMU:
                rule_info = rules.get("14-02", {})
                issues.append(Issue(
                    issueType=rule_info.get('type_scene', "段落-自然段左空两字"),
                    specificWord=text_content[:20] ,
                    sentence=text_content,
                    suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
                    rule_id="14-02",
                    additionalNotes=rule_info.get('description', "段落首行应当左空二字")
                ))
            
        # Rule 14-03: 段落- 回行顶格 (applies to non-heading paragraphs)
        if not is_any_known_heading:
            left_indent = p.paragraph_format.left_indent
            # Check if left_indent is not None and greater than tolerance (i.e., has some left indent)
            if left_indent is not None and left_indent > INDENT_TOLERANCE: 
                rule_info = rules.get("14-03", {})
                issues.append(Issue(
                    issueType=rule_info.get('type_scene', "段落- 回行顶格"),
                    specificWord=text_content[:20] ,
                    sentence=text_content,
                    suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
                    rule_id="14-03",
                    additionalNotes=rule_info.get('description', "段落回行应顶格（段落整体左侧不应有额外缩进）")
                ))

        def get_dominant_font_properties(paragraph):
            run_details = []
            total_text_len = 0
            for run in paragraph.runs:
                run_text = run.text
                if run_text.strip(): # Consider only runs with actual text
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
                    if ef.lower() in actual_font_name.lower():
                        font_ok = True
                        break
            else: # No dominant font found, likely an issue or empty runs with formatting
                font_ok = False 
            
            # Handling bold: If expected_bold is True, actual_bold must be True.
            # If expected_bold is False, actual_bold must be False or None.
            # If expected_bold is None, actual_bold can be anything (not checked).
            bold_ok = True # Assume ok unless a specific check fails
            if expected_bold is True and not actual_bold:
                bold_ok = False
            elif expected_bold is False and actual_bold is True:
                bold_ok = False
            
            # Special case for 黑体/SimHei which is often inherently bold or its 'bold' flag is set
            if actual_font_name and actual_font_name.lower() in ["simhei", "黑体"] and expected_bold is True:
                 bold_ok = True 

            if not font_ok or not bold_ok:
                notes = []
                if not font_ok:
                    notes.append(f"字体应为 '{'/'.join(expected_fonts)}' 系列, 实际主要字体为 '{actual_font_name if actual_font_name else '未知'}'")
                if not bold_ok:
                    expected_bold_str = "加粗" if expected_bold is True else ("不加粗" if expected_bold is False else "未指定")
                    actual_bold_str = "加粗" if actual_bold is True else ("不加粗" if actual_bold is False else "未明确") # actual_bold can be None
                    notes.append(f"字体应 {expected_bold_str}, 实际 {actual_bold_str}")
                
                # Try to get the heading prefix for specificWord
                heading_prefix = p_text_content.split(' ')[0]
                if len(heading_prefix) > 15: heading_prefix = heading_prefix[:15] 

                issues.append(Issue(
                    issueType=rule_info.get('type_scene', f"格式问题-{rule_id}"),
                    specificWord=heading_prefix,
                    sentence=p_text_content,
                    suggestion=Suggestion(operation=rule_info.get('operation_suggestion', "提醒"), after=""),
                    rule_id=rule_id,
                    additionalNotes=f"{rule_info.get('description', '标题格式问题')}: {'; '.join(notes)}"
                ))

        # Rule 06-05: 结构-一级标题格式 (一、，黑体)
        if is_heading_level1:
            check_font_style(p, text_content, "06-05", ["黑体", "SimHei"], expected_bold=True)

        # Rule 06-03: 结构-二级标题格式 ((一)，楷体)
        if is_heading_level2:
            check_font_style(p, text_content, "06-03", ["楷体", "KaiTi", "STKaiti"], expected_bold=False) # Typically not bold

        # Rule 06-04: 结构-三级标题格式 (1.，仿宋_GB2312)
        if is_heading_level3:
            check_font_style(p, text_content, "06-04", ["仿宋", "FangSong", "FangSong_GB2312"], expected_bold=False)

        # Rule 06-02: 结构-四级标题格式 ((1)，仿宋_GB2312)
        if is_heading_level4:
            check_font_style(p, text_content, "06-02", ["仿宋", "FangSong", "FangSong_GB2312"], expected_bold=False)

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
    docx_file="./标准测试html.txt"
    # docx_file="./转完p1.docx"
    rules_file = "./rules_p1.md"
    output_file = "./validation_result.json"
    
    print(f"开始校验文档: {docx_file}")
    print(f"使用规则文件: {rules_file}")
    
    # 执行校验并输出到文件
    validate_and_output_json(docx_file, rules_file)