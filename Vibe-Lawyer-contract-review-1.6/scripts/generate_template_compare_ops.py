import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W_NS}


CATALOG = [
    {
        'name': '付款',
        'keywords': ['付款', '支付', '价款', '结算', '款项', '发票'],
        'risk': '付款条件不清时，容易引发价款争议。',
        'suggestion': '参考自有模板补强付款触发条件和付款期限。',
    },
    {
        'name': '验收',
        'keywords': ['验收', '交付', '测试', '交付标准', '交付成果'],
        'risk': '验收标准不清时，容易引发交付争议。',
        'suggestion': '参考自有模板明确交付标准和验收机制。',
    },
    {
        'name': '违约责任',
        'keywords': ['违约', '赔偿', '违约责任', '损失', '违约金'],
        'risk': '违约责任偏弱时，守约方索赔空间会受限。',
        'suggestion': '参考自有模板补强违约责任和赔偿范围。',
    },
    {
        'name': '解除',
        'keywords': ['解除', '终止', '解约'],
        'risk': '解除条件不清时，退出机制容易失衡。',
        'suggestion': '参考自有模板明确解除条件和解除后责任。',
    },
    {
        'name': '责任限制',
        'keywords': ['责任限制', '责任上限', '最高责任', '间接损失', '损失上限'],
        'risk': '责任限制失衡时，风险承担可能明显偏离。',
        'suggestion': '参考自有模板校准责任上限和例外责任。',
    },
    {
        'name': '知识产权',
        'keywords': ['知识产权', '著作权', '专利', '技术成果', '成果归属'],
        'risk': '知识产权归属不清时，容易影响成果控制权。',
        'suggestion': '参考自有模板明确成果归属和侵权责任。',
    },
    {
        'name': '保密',
        'keywords': ['保密', '秘密信息', '商业秘密', '披露'],
        'risk': '保密义务不足时，容易造成信息泄露风险。',
        'suggestion': '参考自有模板补强保密范围和期限。',
    },
    {
        'name': '数据合规',
        'keywords': ['数据', '个人信息', '隐私', '网络安全', '信息安全'],
        'risk': '数据义务缺失时，容易产生合规风险。',
        'suggestion': '参考自有模板补强数据处理和安全责任。',
    },
    {
        'name': '争议解决',
        'keywords': ['争议解决', '争议', '争议处理', '仲裁', '管辖', '法院', '法律适用'],
        'risk': '争议解决约定偏弱时，争议处理成本会提高。',
        'suggestion': '参考自有模板明确争议解决路径和管辖。',
    },
]


HEADING_PATTERNS = [
    re.compile(r'^(第[一二三四五六七八九十百零〇\d]+条)\s*'),
    re.compile(r'^[一二三四五六七八九十百零〇]+、\s*'),
    re.compile(r'^\(?\d+(?:\.\d+){0,3}\)?[\.\s、]*'),
    re.compile(r'^（[一二三四五六七八九十百零〇]+）\s*'),
]

WEAK_HEADING_HINTS = ['条款', '方式', '机制', '责任', '义务', '安排', '规则', '标准', '要求']


def norm(text: str | None) -> str:
    if not text:
        return ''
    return re.sub(r'\s+', ' ', text).strip()


def compare_key(text: str | None) -> str:
    return re.sub(r'[^\w\u4e00-\u9fff]+', '', norm(text).lower())


def bigrams(text: str | None) -> set[str]:
    key = compare_key(text)
    if not key:
        return set()
    if len(key) == 1:
        return {key}
    return {key[i:i + 2] for i in range(len(key) - 1)}


def similarity(left: str | None, right: str | None) -> float:
    left_set = bigrams(left)
    right_set = bigrams(right)
    if not left_set or not right_set:
        return 0.0
    intersection = len(left_set & right_set)
    union = len(left_set | right_set)
    return intersection / union if union else 0.0


def paragraph_text(paragraph: ET.Element) -> str:
    return ''.join(node.text or '' for node in paragraph.findall('.//w:t', NS))


def paragraph_style(paragraph: ET.Element) -> str:
    style = paragraph.find('./w:pPr/w:pStyle', NS)
    return style.get(f'{{{W_NS}}}val', '') if style is not None else ''


def read_docx_paragraphs(path: Path) -> list[dict]:
    with zipfile.ZipFile(path) as archive:
        document_xml = archive.read('word/document.xml')
    root = ET.fromstring(document_xml)
    occurrences: dict[str, int] = {}
    paragraphs: list[dict] = []
    for index, paragraph in enumerate(root.findall('.//w:p', NS)):
        text = paragraph_text(paragraph)
        normalized = norm(text)
        occurrences[normalized] = occurrences.get(normalized, 0) + 1
        paragraphs.append({
            'index': index,
            'text': text,
            'normalized': normalized,
            'occurrence': occurrences[normalized],
            'style_id': paragraph_style(paragraph),
        })
    return paragraphs


def is_heading(paragraph: dict) -> bool:
    text = paragraph['normalized']
    style_id = paragraph['style_id'] or ''
    if not text:
        return False
    if re.search(r'heading|title', style_id, re.IGNORECASE):
        return True
    return any(pattern.search(text) for pattern in HEADING_PATTERNS)


def weak_heading_score(paragraph: dict, next_paragraph: dict | None) -> int:
    text = paragraph['normalized']
    if not text or is_heading(paragraph):
        return 0
    if any(mark in text for mark in ['。', '；', '，', ',']):
        return 0
    score = 0
    if len(text) <= 18:
        score += 1
    if any(hint in text for hint in WEAK_HEADING_HINTS):
        score += 1
    if resolve_topic(text):
        score += 2
    if next_paragraph and next_paragraph.get('normalized') and len(next_paragraph['normalized']) > len(text):
        score += 1
    if text.endswith('：') or text.endswith(':'):
        score += 1
    return score


def is_weak_heading(paragraph: dict, next_paragraph: dict | None) -> bool:
    return weak_heading_score(paragraph, next_paragraph) >= 3


def clause_number(text: str) -> str:
    for pattern in HEADING_PATTERNS:
        match = pattern.search(norm(text))
        if match:
            return match.group(0).strip()
    return ''


def clause_title(text: str) -> str:
    title = norm(text)
    for pattern in HEADING_PATTERNS:
        title = pattern.sub('', title)
    return title.strip()


def resolve_topic(text: str) -> dict | None:
    best_topic = None
    best_score = 0
    for topic in CATALOG:
        score = sum(1 for keyword in topic['keywords'] if keyword and keyword in text)
        if score > best_score:
            best_score = score
            best_topic = topic
    return best_topic if best_score > 0 else None


def finalize_clause(clause: dict) -> dict:
    clause['body_text'] = '\n'.join(item['text'] for item in clause['paragraphs'][1:]).strip()
    clause['full_text'] = '\n'.join(item['text'] for item in clause['paragraphs']).strip()
    topic = resolve_topic(clause['full_text'])
    clause['topic_name'] = topic['name'] if topic else ''
    return clause


def build_clauses(paragraphs: list[dict]) -> list[dict]:
    clauses: list[dict] = []
    current = None
    for index, paragraph in enumerate(paragraphs):
        next_paragraph = paragraphs[index + 1] if index + 1 < len(paragraphs) else None
        detection = 'strong' if is_heading(paragraph) else ('weak' if is_weak_heading(paragraph, next_paragraph) else '')
        if current is None or detection:
            if current is not None:
                clauses.append(finalize_clause(current))
            current = {
                'start_index': paragraph['index'],
                'end_index': paragraph['index'],
                'heading_text': paragraph['text'],
                'title_text': clause_title(paragraph['text']),
                'number_key': clause_number(paragraph['text']),
                'paragraphs': [paragraph],
                'clause_detection': detection or 'fallback',
            }
        else:
            current['end_index'] = paragraph['index']
            current['paragraphs'].append(paragraph)
    if current is not None:
        clauses.append(finalize_clause(current))
    return clauses


def representative_paragraph(clause: dict) -> dict:
    body_paragraphs = [item for item in clause['paragraphs'] if item['index'] != clause['start_index']]
    if body_paragraphs:
        return max(body_paragraphs, key=lambda item: len(item['normalized']))
    return clause['paragraphs'][0]


def alignment(source_clause: dict, template_clause: dict, topic_name: str) -> dict:
    title_score = similarity(source_clause['title_text'], template_clause['title_text'])
    body_score = similarity(source_clause['body_text'], template_clause['body_text'])
    full_score = similarity(source_clause['full_text'], template_clause['full_text'])
    number_score = 1.0 if source_clause['number_key'] and source_clause['number_key'] == template_clause['number_key'] else 0.0
    topic_score = 1.0 if source_clause.get('topic_name') == topic_name and template_clause.get('topic_name') == topic_name else 0.0
    confidence = title_score * 0.25 + body_score * 0.35 + full_score * 0.2 + number_score * 0.1 + topic_score * 0.1
    reasons = []
    if number_score:
        reasons.append('条号一致')
    if title_score >= 0.6:
        reasons.append('标题接近')
    if body_score >= 0.45:
        reasons.append('正文接近')
    if topic_score:
        reasons.append('主题命中')
    return {
        'confidence': round(confidence, 4),
        'title_score': round(title_score, 4),
        'body_score': round(body_score, 4),
        'full_score': round(full_score, 4),
        'reason': '、'.join(reasons) if reasons else '主题回退匹配',
    }


def find_topic(name: str) -> dict | None:
    for topic in CATALOG:
        if topic['name'] == name:
            return topic
    return None


def build_comment(topic: dict, mode: str, template_text: str, reason: str | None = None) -> dict:
    comment = {
        '问题': f'对方合同在{topic["name"]}条款上的表述与自有模板不一致，当前文本未采用自有模板的风控口径。',
        '风险': topic['risk'],
        '修改建议': '该条款与自有模板差异较大，建议人工确认交易口径后按自有模板补强。' if mode == 'comment' else topic['suggestion'],
        '建议条款': template_text,
        '修改依据': '依据自有模板同主题条款比对结果。' if not reason else f'依据自有模板同主题条款比对结果。对齐依据：{reason}',
    }
    return comment


def source_clause_for_topic(source_clauses: list[dict], source_paragraphs: list[dict], topic: dict, topic_name: str) -> tuple[dict | None, str]:
    direct_clause = next((item for item in source_clauses if item.get('topic_name') == topic_name), None)
    if direct_clause is not None:
        return direct_clause, direct_clause.get('clause_detection', 'strong')
    fallback_clause = fallback_clause_for_topic(source_paragraphs, topic)
    if fallback_clause is not None:
        return fallback_clause, 'fallback'
    return None, 'none'


def template_clause_for_topic(template_clauses: list[dict], template_paragraphs: list[dict], topic: dict, topic_name: str) -> tuple[dict | None, str]:
    direct_clause = next((item for item in template_clauses if item.get('topic_name') == topic_name), None)
    if direct_clause is not None:
        return direct_clause, direct_clause.get('clause_detection', 'strong')
    fallback_clause = fallback_clause_for_topic(template_paragraphs, topic)
    if fallback_clause is not None:
        return fallback_clause, 'fallback'
    return None, 'none'


def clause_detection_label(value: str | None) -> str:
    return {
        'strong': '强标题识别',
        'weak': '弱标题识别',
        'fallback': '正文块兜底识别',
        'none': '未识别到条款块',
    }.get(value or '', '未识别到条款块')


def build_review_summary(alignment_reasons_summary: list[dict], missing_reasons_summary: list[dict]) -> tuple[list[str], str]:
    lines: list[str] = []
    lines.append(f'本次模板比对共识别 {len(alignment_reasons_summary)} 处已对齐条款，{len(missing_reasons_summary)} 个缺失主题。')
    for item in alignment_reasons_summary:
        lines.append(
            f'{item["topic_name"]}：已按{item["alignment_reason"]}完成{item["mode"]}，置信度 {item["alignment_confidence"]}；'
            f'对方合同识别方式为{clause_detection_label(item.get("source_clause_detection"))}。'
        )
    for item in missing_reasons_summary:
        action = '已计划自动插入' if item.get('insertion_planned') else '当前仅提示风险'
        placement = item.get('insertion_placement')
        placement_text = ''
        if placement == 'before':
            placement_text = '，插入位置在已匹配条款前'
        elif placement == 'after':
            placement_text = '，插入位置在已匹配条款后'
        lines.append(
            f'{item["topic_name"]}：缺失原因={item["missing_reason"]}，模板识别方式为{clause_detection_label(item.get("template_clause_detection"))}，'
            f'{action}{placement_text}。'
        )
    return lines, '\n'.join(lines)


def clause_requires_manual_review(source_clause: dict, template_clause: dict) -> bool:
    return len(source_clause.get('paragraphs', [])) > 2 or len(template_clause.get('paragraphs', [])) > 2


def template_topic_sequence(template_clauses: list[dict], focus_topics: list[str]) -> list[dict]:
    sequence: list[dict] = []
    seen: set[str] = set()
    for clause in template_clauses:
        topic_name = clause.get('topic_name') or ''
        if not topic_name or topic_name not in focus_topics or topic_name in seen:
            continue
        seen.add(topic_name)
        sequence.append({
            'topic_name': topic_name,
            'template_clause': clause,
        })
    for topic_name in focus_topics:
        if topic_name in seen:
            continue
        topic = find_topic(topic_name)
        if topic is None:
            continue
        fallback_clause = fallback_clause_for_topic([item for clause in template_clauses for item in clause['paragraphs']], topic)
        if fallback_clause is None:
            continue
        seen.add(topic_name)
        sequence.append({
            'topic_name': topic_name,
            'template_clause': fallback_clause,
        })
    return sequence


def matched_topic_lookup(source_clauses: list[dict], matched_topics: list[str]) -> dict[str, dict]:
    lookup: dict[str, dict] = {}
    for clause in source_clauses:
        topic_name = clause.get('topic_name') or ''
        if not topic_name or topic_name not in matched_topics or topic_name in lookup:
            continue
        lookup[topic_name] = clause
    return lookup


def insertion_anchor_for_missing(
    topic_name: str,
    template_sequence: list[dict],
    matched_lookup: dict[str, dict],
    source_clauses: list[dict],
) -> dict:
    topic_names = [item['topic_name'] for item in template_sequence]
    if topic_name not in topic_names:
        if source_clauses:
            return {'paragraph_index': source_clauses[-1]['end_index'], 'placement': 'after', 'template_index': len(template_sequence)}
        return {'paragraph_index': 0, 'placement': 'after', 'template_index': len(template_sequence)}
    template_index = topic_names.index(topic_name)

    previous_match = None
    for index in range(template_index - 1, -1, -1):
        candidate = matched_lookup.get(topic_names[index])
        if candidate is not None:
            previous_match = candidate
            break

    next_match = None
    for index in range(template_index + 1, len(topic_names)):
        candidate = matched_lookup.get(topic_names[index])
        if candidate is not None:
            next_match = candidate
            break

    if next_match is not None:
        return {
            'paragraph_index': next_match['start_index'],
            'placement': 'before',
            'template_index': template_index,
        }
    if previous_match is not None:
        return {
            'paragraph_index': previous_match['end_index'],
            'placement': 'after',
            'template_index': template_index,
        }
    if source_clauses:
        return {
            'paragraph_index': source_clauses[-1]['end_index'],
            'placement': 'after',
            'template_index': template_index,
        }
    return {'paragraph_index': 0, 'placement': 'after', 'template_index': template_index}


def paragraph_matches_topic(paragraph: dict, topic: dict) -> bool:
    text = norm(paragraph.get('text'))
    return any(keyword in text for keyword in topic['keywords'])


def paragraph_switches_topic(paragraph: dict, topic: dict) -> bool:
    detected_topic = resolve_topic(norm(paragraph.get('text')))
    return detected_topic is not None and detected_topic['name'] != topic['name']


def fallback_clause_for_topic(paragraphs: list[dict], topic: dict) -> dict | None:
    for index, paragraph in enumerate(paragraphs):
        if not paragraph_matches_topic(paragraph, topic):
            continue
        block = [paragraph]
        cursor = index + 1
        while cursor < len(paragraphs):
            current = paragraphs[cursor]
            next_paragraph = paragraphs[cursor + 1] if cursor + 1 < len(paragraphs) else None
            if not current.get('normalized'):
                break
            if is_heading(current) or is_weak_heading(current, next_paragraph):
                break
            if len(block) >= 1 and paragraph_switches_topic(current, topic) and not paragraph_matches_topic(current, topic):
                break
            block.append(current)
            cursor += 1
            if len(block) >= 4:
                break
        return finalize_clause({
            'start_index': block[0]['index'],
            'end_index': block[-1]['index'],
            'heading_text': block[0]['text'],
            'title_text': clause_title(block[0]['text']),
            'number_key': clause_number(block[0]['text']),
            'paragraphs': block,
            'clause_detection': 'fallback',
        })
    return None


def load_config(path: Path) -> dict:
    return json.loads(path.read_text(encoding='utf-8'))


def resolve_focus_topics(compare_config: dict) -> list[str]:
    topics = compare_config.get('focus_topics') or compare_config.get('topics')
    if not topics:
        return [item['name'] for item in CATALOG]
    return [str(item).strip() for item in topics if str(item).strip()]


def generate(compare_config: dict, source_path: Path, instructions_path: Path, default_mode: str) -> dict:
    template_path = Path(compare_config.get('template_path') or compare_config.get('template') or '')
    if not template_path:
        raise ValueError('template_compare 缺少 template_path')
    if not template_path.is_absolute():
        template_path = (instructions_path.parent / template_path).resolve()
    focus_topics = resolve_focus_topics(compare_config)
    max_operations = int(compare_config.get('max_operations', 8))
    revision_mode = str(compare_config.get('mode') or default_mode).lower()
    min_alignment_confidence = float(compare_config.get('min_alignment_confidence', 0.45))
    min_similarity_for_revision = float(compare_config.get('min_similarity_for_revision', 0.35))
    allow_missing_clause_insert = bool(compare_config.get('allow_missing_clause_insert', False))
    missing_clause_mode = str(compare_config.get('missing_clause_mode') or ('revision_comment' if allow_missing_clause_insert else 'comment')).lower()

    source_paragraphs = read_docx_paragraphs(source_path)
    template_paragraphs = read_docx_paragraphs(template_path)
    source_clauses = build_clauses(source_paragraphs)
    template_clauses = build_clauses(template_paragraphs)
    template_sequence = template_topic_sequence(template_clauses, focus_topics)

    operations: list[dict] = []
    matched_topics: list[str] = []
    missing_topics: list[str] = []
    matched_pairs: list[dict] = []
    used_source: set[int] = set()
    alignment_reasons_summary: list[dict] = []
    missing_reasons_summary: list[dict] = []

    for topic_name in focus_topics:
        if len(operations) >= max_operations:
            break
        topic = find_topic(topic_name)
        if topic is None:
            continue
        source_candidates = sorted([item for item in source_clauses if item.get('topic_name') == topic_name], key=lambda item: item['start_index'])
        template_candidates = sorted([item for item in template_clauses if item.get('topic_name') == topic_name], key=lambda item: item['start_index'])
        if not template_candidates:
            continue
        if not source_candidates:
            missing_topics.append(topic_name)
            continue
        used_template: set[int] = set()
        for source_candidate in source_candidates:
            if len(operations) >= max_operations:
                break
            if source_candidate['start_index'] in used_source:
                continue
            best_alignment = None
            best_template = None
            best_slot = None
            for slot, template_candidate in enumerate(template_candidates):
                if slot in used_template:
                    continue
                current_alignment = alignment(source_candidate, template_candidate, topic_name)
                if best_alignment is None or current_alignment['confidence'] > best_alignment['confidence']:
                    best_alignment = current_alignment
                    best_template = template_candidate
                    best_slot = slot
            if best_template is None or best_alignment is None:
                continue
            used_source.add(source_candidate['start_index'])
            used_template.add(best_slot)
            if norm(source_candidate['full_text']) == norm(best_template['full_text']):
                continue
            source_para = representative_paragraph(source_candidate)
            template_para = representative_paragraph(best_template)
            requires_manual_review = clause_requires_manual_review(source_candidate, best_template)
            alignment_reason = best_alignment['reason']
            if requires_manual_review:
                alignment_reason = f'{alignment_reason}；多段条款自动降级人工确认'
            mode = 'comment' if (
                revision_mode == 'comment'
                or best_alignment['confidence'] < min_alignment_confidence
                or best_alignment['body_score'] < min_similarity_for_revision
                or requires_manual_review
            ) else revision_mode
            comment = build_comment(topic, mode, best_template['full_text'], f'{alignment_reason}；置信度：{best_alignment["confidence"]}')
            operations.append({
                'location': {'paragraph_index': source_para['index']},
                'mode': mode,
                'replacement_text': template_para['text'] if mode in ('revision', 'revision_comment') else None,
                'comment': comment,
                'match_type': 'exact',
                'occurrence': 1,
                'topic_name': topic_name,
                'alignment_reason': alignment_reason,
                'alignment_confidence': best_alignment['confidence'],
                'source_clause_detection': source_candidate.get('clause_detection', 'strong'),
                'template_clause_detection': best_template.get('clause_detection', 'strong'),
            })
            alignment_reasons_summary.append({
                'topic_name': topic_name,
                'mode': mode,
                'alignment_reason': alignment_reason,
                'alignment_confidence': best_alignment['confidence'],
                'source_paragraph_index': source_para['index'],
                'template_paragraph_index': template_para['index'],
                'source_clause_detection': source_candidate.get('clause_detection', 'strong'),
                'template_clause_detection': best_template.get('clause_detection', 'strong'),
            })
            matched_pairs.append({'topic_name': topic_name, 'source_end_index': source_candidate['end_index']})
            if topic_name not in matched_topics:
                matched_topics.append(topic_name)

    for topic_name in focus_topics:
        if len(operations) >= max_operations:
            break
        if topic_name in matched_topics or topic_name in missing_topics:
            continue
        topic = find_topic(topic_name)
        if topic is None:
            continue
        source_clause, source_detection = source_clause_for_topic(source_clauses, source_paragraphs, topic, topic_name)
        template_clause, template_detection = template_clause_for_topic(template_clauses, template_paragraphs, topic, topic_name)
        if source_clause is not None or template_clause is None:
            continue
        missing_topics.append(topic_name)
        missing_reasons_summary.append({
            'topic_name': topic_name,
            'missing_reason': 'source_topic_not_found',
            'source_clause_detection': source_detection,
            'template_clause_detection': template_detection,
            'insertion_planned': allow_missing_clause_insert and missing_clause_mode in ('revision', 'revision_comment'),
        })

    matched_lookup = matched_topic_lookup(source_clauses, matched_topics)
    planned_missing: list[dict] = []
    for topic_name in missing_topics:
        topic = find_topic(topic_name)
        if topic is None:
            continue
        template_clause, template_detection = template_clause_for_topic(template_clauses, template_paragraphs, topic, topic_name)
        if topic is None or template_clause is None:
            continue
        anchor = insertion_anchor_for_missing(topic_name, template_sequence, matched_lookup, source_clauses)
        planned_missing.append({
            'topic_name': topic_name,
            'topic': topic,
            'template_clause': template_clause,
            'anchor': anchor,
            'template_clause_detection': template_detection,
        })

    planned_missing.sort(
        key=lambda item: (
            item['anchor']['paragraph_index'],
            0 if item['anchor']['placement'] == 'before' else 1,
            item['anchor']['template_index'] if item['anchor']['placement'] == 'before' else -item['anchor']['template_index'],
        )
    )

    for item in planned_missing:
        if len(operations) >= max_operations:
            break
        topic_name = item['topic_name']
        topic = item['topic']
        template_clause = item['template_clause']
        anchor = item['anchor']
        comment = {
            '问题': f'对方合同缺少自有模板中的{topic_name}条款。',
            '风险': topic['risk'],
            '修改建议': '按自有模板新增该条款，并保留修订痕迹。' if allow_missing_clause_insert else '建议结合交易背景补充该条款，当前先提示风险。',
            '建议条款': template_clause['full_text'],
            '修改依据': '依据自有模板同主题条款比对结果。',
        }
        missing_reason = 'template_topic_only'
        if allow_missing_clause_insert and missing_clause_mode in ('revision', 'revision_comment'):
            operations.append({
                'location': {'paragraph_index': anchor['paragraph_index']},
                'mode': missing_clause_mode,
                'comment': comment,
                'match_type': 'exact',
                'occurrence': 1,
                'insertion_texts': [paragraph['text'] for paragraph in template_clause['paragraphs']],
                'insertion_placement': anchor['placement'],
                'topic_name': topic_name,
                'missing_reason': missing_reason,
            })
        elif source_paragraphs:
            anchor_paragraph = source_paragraphs[-1]
            operations.append({
                'location': {'paragraph_index': anchor_paragraph['index']},
                'mode': 'comment',
                'comment': comment,
                'match_type': 'exact',
                'occurrence': 1,
                'topic_name': topic_name,
                'missing_reason': missing_reason,
            })
        missing_reasons_summary.append({
            'topic_name': topic_name,
            'missing_reason': missing_reason,
            'template_clause_detection': item['template_clause_detection'],
            'insertion_planned': allow_missing_clause_insert and missing_clause_mode in ('revision', 'revision_comment'),
            'insertion_placement': anchor['placement'],
            'anchor_paragraph_index': anchor['paragraph_index'],
        })

    if not operations:
        pair_count = min(len(source_paragraphs), len(template_paragraphs))
        for index in range(pair_count):
            if len(operations) >= max_operations:
                break
            source_candidate = source_paragraphs[index]
            template_candidate = template_paragraphs[index]
            if norm(source_candidate['text']) == norm(template_candidate['text']):
                continue
            topic = resolve_topic(f'{source_candidate["text"]} {template_candidate["text"]}') or {
                'name': '对应条款',
                'risk': '该条款与自有模板存在差异，可能导致风险控制口径不一致。',
                'suggestion': '参考自有模板对对应条款进行补强。',
            }
            fallback_mode = revision_mode
            if revision_mode != 'comment' and similarity(source_candidate['text'], template_candidate['text']) < min_similarity_for_revision:
                fallback_mode = 'comment'
            comment = build_comment(topic, fallback_mode, template_candidate['text'])
            operations.append({
                'location': {'paragraph_index': source_candidate['index']},
                'mode': fallback_mode,
                'replacement_text': template_candidate['text'] if fallback_mode in ('revision', 'revision_comment') else None,
                'comment': comment,
                'match_type': 'exact',
                'occurrence': 1,
                'topic_name': topic['name'],
                'alignment_reason': '顺序回退匹配',
                'alignment_confidence': 0.0,
                'source_clause_detection': 'fallback',
                'template_clause_detection': 'fallback',
            })
            alignment_reasons_summary.append({
                'topic_name': topic['name'],
                'mode': fallback_mode,
                'alignment_reason': '顺序回退匹配',
                'alignment_confidence': 0.0,
                'source_paragraph_index': source_candidate['index'],
                'template_paragraph_index': template_candidate['index'],
                'source_clause_detection': 'fallback',
                'template_clause_detection': 'fallback',
            })

    review_summary_lines, review_summary_text = build_review_summary(alignment_reasons_summary, missing_reasons_summary)

    return {
        'template_path': str(template_path),
        'operations': operations,
        'focus_topics': focus_topics,
        'matched_topics': matched_topics,
        'missing_topics': missing_topics,
        'alignment_confidence': min_alignment_confidence,
        'missing_clause_insertion_enabled': allow_missing_clause_insert,
        'alignment_reasons_summary': alignment_reasons_summary,
        'missing_reasons_summary': missing_reasons_summary,
        'review_summary_lines': review_summary_lines,
        'review_summary_text': review_summary_text,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--source', required=True)
    parser.add_argument('--instructions', required=True)
    parser.add_argument('--default-mode', required=True)
    parser.add_argument('--output-json')
    args = parser.parse_args()

    instructions_path = Path(args.instructions).resolve()
    config = load_config(instructions_path)
    compare_config = config.get('template_compare') or {}
    result = generate(compare_config, Path(args.source).resolve(), instructions_path, args.default_mode)
    if args.output_json:
        Path(args.output_json).write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
    else:
        json.dump(result, sys.stdout, ensure_ascii=False)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
