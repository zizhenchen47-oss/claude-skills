import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from shutil import which

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from test_xml_revision import collect_docx_summary


def resolve_powershell_executable() -> str:
    for candidate in ('pwsh', 'powershell', 'powershell.exe'):
        executable = which(candidate)
        if executable:
            return executable
    raise RuntimeError('未找到可用的 PowerShell 可执行文件（pwsh / powershell / powershell.exe）')


def create_docx_with_paragraphs(output_path: str, paragraphs: list[str]) -> None:
    temp_dir = tempfile.mkdtemp()
    try:
        word_dir = os.path.join(temp_dir, 'word')
        rels_dir = os.path.join(temp_dir, '_rels')
        word_rels_dir = os.path.join(word_dir, '_rels')
        os.makedirs(word_dir)
        os.makedirs(rels_dir)
        os.makedirs(word_rels_dir)

        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

        paragraph_xml = []
        for text in paragraphs:
            escaped = (
                text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
            )
            paragraph_xml.append(
                '    <w:p><w:r><w:t>{}</w:t></w:r></w:p>'.format(escaped)
            )
        document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
{}
  </w:body>
</w:document>'''.format('\n'.join(paragraph_xml))

        settings_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat/>
</w:settings>'''

        root_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

        document_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

        with open(os.path.join(temp_dir, '[Content_Types].xml'), 'w', encoding='utf-8') as handle:
            handle.write(content_types)
        with open(os.path.join(word_dir, 'document.xml'), 'w', encoding='utf-8') as handle:
            handle.write(document_xml)
        with open(os.path.join(word_dir, 'settings.xml'), 'w', encoding='utf-8') as handle:
            handle.write(settings_xml)
        with open(os.path.join(rels_dir, '.rels'), 'w', encoding='utf-8') as handle:
            handle.write(root_rels_xml)
        with open(os.path.join(word_rels_dir, 'document.xml.rels'), 'w', encoding='utf-8') as handle:
            handle.write(document_rels_xml)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as archive:
            for root, _, files in os.walk(temp_dir):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    archive_name = os.path.relpath(file_path, temp_dir).replace(os.sep, '/')
                    archive.write(file_path, archive_name)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def build_python3_only_environment() -> tuple[dict[str, str], str]:
    shim_dir = tempfile.mkdtemp(prefix='python3-only-')
    proxy_script = os.path.join(shim_dir, 'python3_proxy.py')
    proxy_command = os.path.join(shim_dir, 'python3.cmd')
    with open(proxy_script, 'w', encoding='utf-8') as handle:
        handle.write(
            'import subprocess\n'
            'import sys\n'
            '\n'
            f'REAL_PYTHON = {json.dumps(os.path.abspath(sys.executable))}\n'
            'args = sys.argv[1:]\n'
            'for index, arg in enumerate(args):\n'
            "    if arg == '--instructions' and index + 1 < len(args):\n"
            '        value = args[index + 1]\n'
            '        try:\n'
            "            value.encode('ascii')\n"
            '        except UnicodeEncodeError:\n'
            "            print('non-ascii instructions path blocked', file=sys.stderr)\n"
            '            sys.exit(91)\n'
            '        break\n'
            'sys.exit(subprocess.run([REAL_PYTHON, *args]).returncode)\n'
        )
    with open(proxy_command, 'w', encoding='utf-8', newline='\r\n') as handle:
        handle.write(
            '@echo off\r\n'
            f'"{os.path.abspath(sys.executable)}" "%~dp0python3_proxy.py" %*\r\n'
        )

    filtered_path_entries = []
    for entry in os.environ.get('PATH', '').split(os.pathsep):
        normalized = entry.strip()
        lowered = normalized.lower()
        if not normalized:
            continue
        if 'python' in lowered or 'windowsapps' in lowered:
            continue
        filtered_path_entries.append(normalized)

    env = os.environ.copy()
    env['PATH'] = os.pathsep.join([shim_dir, *filtered_path_entries])
    return env, shim_dir


def run_template_compare_case(
    repo_root: str,
    powershell_executable: str,
    source_paragraphs: list[str],
    template_paragraphs: list[str],
    template_compare: dict,
    operations: list[dict] | None = None,
    env: dict[str, str] | None = None,
    case_subdir: str | None = None,
) -> tuple[dict, dict]:
    generate_script = os.path.join(repo_root, 'scripts', 'run_generate_review_docx.ps1')
    temp_dir = tempfile.mkdtemp(prefix='template-compare-case-')
    try:
        case_dir = os.path.join(temp_dir, case_subdir) if case_subdir else temp_dir
        os.makedirs(case_dir, exist_ok=True)
        source_docx = os.path.join(case_dir, 'counterparty-contract.docx')
        template_docx = os.path.join(case_dir, 'company-template.docx')
        instructions_json = os.path.join(case_dir, 'template-compare.json')
        output_docx = os.path.join(case_dir, 'counterparty-reviewed.docx')
        create_docx_with_paragraphs(source_docx, source_paragraphs)
        create_docx_with_paragraphs(template_docx, template_paragraphs)
        payload = {
            'author': '模板比对测试',
            'date': '2026-04-02T10:00:00Z',
            'template_compare': {
                'template_path': './company-template.docx',
                **template_compare,
            },
            'operations': operations or []
        }
        with open(instructions_json, 'w', encoding='utf-8') as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
        command = [
            powershell_executable,
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            generate_script,
            '-Source',
            source_docx,
            '-Instructions',
            instructions_json,
            '-Output',
            output_docx,
        ]
        result = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            cwd=repo_root,
            env=env
        )
        if result.returncode != 0:
            raise RuntimeError(
                f'模板比对 MVP 执行失败\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}'
            )
        return json.loads(result.stdout), collect_docx_summary(output_docx)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main() -> int:
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    powershell_executable = resolve_powershell_executable()
    try:
        summary, docx_summary = run_template_compare_case(
            repo_root,
            powershell_executable,
            [
                '第一条 付款条款',
                '乙方应在验收后30日内付款。',
                '第二条 争议解决',
                '争议由双方协商解决，协商不成的，向原告所在地人民法院起诉。'
            ],
            [
                '第一条 付款条款',
                '乙方应在验收合格且收到合法有效发票后10个工作日内付款。',
                '第二条 争议解决',
                '争议由双方协商解决，协商不成的，任一方可向合同签订地有管辖权的人民法院起诉。'
            ],
            {
                'focus_topics': ['付款', '争议解决'],
                'max_operations': 4,
                'mode': 'revision_comment',
                'min_similarity_for_revision': 0.2
            }
        )
        assert summary['template_compare_applied'] is True, json.dumps(summary, ensure_ascii=False, indent=2)
        assert summary['template_generated_operations'] == 2, json.dumps(summary, ensure_ascii=False, indent=2)
        assert len(summary['template_compare_focus_topics']) == 2, json.dumps(summary, ensure_ascii=False, indent=2)
        assert len(summary['alignment_reasons_summary']) == 2, json.dumps(summary, ensure_ascii=False, indent=2)
        assert 'review_summary_text' in summary and 'revision_comment' in summary['review_summary_text'], json.dumps(summary, ensure_ascii=False, indent=2)
        assert docx_summary['comment_range_start_count'] == 2, json.dumps(docx_summary, ensure_ascii=False, indent=2)
        assert docx_summary['comment_reference_count'] == 2, json.dumps(docx_summary, ensure_ascii=False, indent=2)
        assert docx_summary['ins_count'] >= 2, json.dumps(docx_summary, ensure_ascii=False, indent=2)
        assert len(docx_summary['visible_paragraph_texts']) == 4, json.dumps(docx_summary, ensure_ascii=False, indent=2)
        assert any('依据自有模板同主题条款比对结果。' in text for text in docx_summary['comment_texts']), json.dumps(docx_summary, ensure_ascii=False, indent=2)

        phase2_summary, phase2_docx_summary = run_template_compare_case(
            repo_root,
            powershell_executable,
            [
                '第一条 争议解决',
                '争议由双方协商解决，协商不成的，向原告所在地人民法院起诉。',
                '第二条 付款条款',
                '乙方应在验收后30日内付款。'
            ],
            [
                '第一条 付款条款',
                '乙方应在验收合格且收到合法有效发票后10个工作日内付款。',
                '第二条 争议解决',
                '争议由双方协商解决，协商不成的，任一方可向合同签订地有管辖权的人民法院起诉。'
            ],
            {
                'focus_topics': ['付款', '争议解决'],
                'max_operations': 6,
                'mode': 'revision_comment',
                'min_similarity_for_revision': 0.2,
                'min_alignment_confidence': 0.2
            }
        )
        assert phase2_summary['template_compare_applied'] is True
        assert phase2_summary['template_generated_operations'] == 2, json.dumps(phase2_summary, ensure_ascii=False, indent=2)
        assert phase2_docx_summary['comment_range_start_count'] == 2, json.dumps(phase2_docx_summary, ensure_ascii=False, indent=2)
        assert phase2_docx_summary['ins_count'] >= 2, json.dumps(phase2_docx_summary, ensure_ascii=False, indent=2)
        assert len(phase2_docx_summary['visible_paragraph_texts']) == 4, json.dumps(phase2_docx_summary, ensure_ascii=False, indent=2)

        insertion_summary, insertion_docx_summary = run_template_compare_case(
            repo_root,
            powershell_executable,
            [
                '第一条 付款条款',
                '乙方应在验收后30日内付款。'
            ],
            [
                '第一条 付款条款',
                '乙方应在验收后30日内付款。'
            ],
            {
                'focus_topics': [],
                'max_operations': 1,
                'mode': 'comment'
            },
            operations=[
                {
                    'mode': 'revision_comment',
                    'location': {'paragraph': 2},
                    'comment': '对方合同缺少保密条款。',
                    'insertion_texts': [
                        '第二条 保密条款',
                        '双方应对在合作过程中获悉的商业秘密承担保密义务。'
                    ]
                }
            ]
        )
        assert insertion_summary['operations_applied'] == 1, json.dumps(insertion_summary, ensure_ascii=False, indent=2)
        assert insertion_docx_summary['comment_range_start_count'] == 2, json.dumps(insertion_docx_summary, ensure_ascii=False, indent=2)
        assert len(insertion_docx_summary['visible_paragraph_texts']) >= 4, json.dumps(insertion_docx_summary, ensure_ascii=False, indent=2)

        phase3_summary, phase3_docx_summary = run_template_compare_case(
            repo_root,
            powershell_executable,
            [
                '第一条 付款条款',
                '乙方应在验收后30日内付款。',
                '第二条 争议解决',
                '争议由双方协商解决，协商不成的，向原告所在地人民法院起诉。'
            ],
            [
                '第一条 保密条款',
                '双方应对在合作过程中获悉的商业秘密承担保密义务。',
                '第二条 付款条款',
                '乙方应在验收合格且收到合法有效发票后10个工作日内付款。',
                '第三条 违约责任',
                '任何一方违约的，应赔偿守约方因此遭受的全部损失。',
                '第四条 争议解决',
                '争议由双方协商解决，协商不成的，任一方可向合同签订地有管辖权的人民法院起诉。'
            ],
            {
                'focus_topics': ['保密', '付款', '违约责任', '争议解决'],
                'max_operations': 8,
                'mode': 'revision_comment',
                'min_similarity_for_revision': 0.2,
                'min_alignment_confidence': 0.2,
                'allow_missing_clause_insert': True,
                'missing_clause_mode': 'revision_comment'
            }
        )
        assert phase3_summary['template_compare_applied'] is True, json.dumps(phase3_summary, ensure_ascii=False, indent=2)
        assert phase3_summary['template_generated_operations'] >= 4, json.dumps(phase3_summary, ensure_ascii=False, indent=2)
        assert len(phase3_docx_summary['visible_paragraph_texts']) >= 8, json.dumps(phase3_docx_summary, ensure_ascii=False, indent=2)
        assert phase3_docx_summary['comment_range_start_count'] >= 6, json.dumps(phase3_docx_summary, ensure_ascii=False, indent=2)
        assert len(phase3_summary['missing_reasons_summary']) == 2, json.dumps(phase3_summary, ensure_ascii=False, indent=2)

        phase4_summary, phase4_docx_summary = run_template_compare_case(
            repo_root,
            powershell_executable,
            [
                '付款方式',
                '乙方应在验收后30日内付款。',
                '争议解决',
                '争议由双方协商解决，协商不成的，向原告所在地人民法院起诉。'
            ],
            [
                '付款方式',
                '乙方应在验收合格且收到合法有效发票后10个工作日内付款。',
                '保密',
                '双方应对在合作过程中获悉的商业秘密承担保密义务。',
                '争议解决',
                '争议由双方协商解决，协商不成的，任一方可向合同签订地有管辖权的人民法院起诉。'
            ],
            {
                'focus_topics': ['付款', '保密', '争议解决'],
                'max_operations': 6,
                'mode': 'revision_comment',
                'min_similarity_for_revision': 0.2,
                'min_alignment_confidence': 0.2,
                'allow_missing_clause_insert': True,
                'missing_clause_mode': 'revision_comment'
            }
        )
        assert phase4_summary['template_compare_applied'] is True, json.dumps(phase4_summary, ensure_ascii=False, indent=2)
        assert len(phase4_summary['alignment_reasons_summary']) >= 2, json.dumps(phase4_summary, ensure_ascii=False, indent=2)
        assert len(phase4_summary['missing_reasons_summary']) >= 1, json.dumps(phase4_summary, ensure_ascii=False, indent=2)
        assert any(item['source_clause_detection'] in ('weak', 'fallback') for item in phase4_summary['alignment_reasons_summary']), json.dumps(phase4_summary, ensure_ascii=False, indent=2)
        assert any(item['missing_reason'] == 'template_topic_only' for item in phase4_summary['missing_reasons_summary']), json.dumps(phase4_summary, ensure_ascii=False, indent=2)
        assert 'review_summary_text' in phase4_summary and len(phase4_summary['review_summary_lines']) >= 3, json.dumps(phase4_summary, ensure_ascii=False, indent=2)
        assert len(phase4_docx_summary['visible_paragraph_texts']) >= 4, json.dumps(phase4_docx_summary, ensure_ascii=False, indent=2)
        assert phase4_docx_summary['comment_range_start_count'] >= 3, json.dumps(phase4_docx_summary, ensure_ascii=False, indent=2)

        python3_only_env, shim_dir = build_python3_only_environment()
        try:
            linux_style_summary, linux_style_docx_summary = run_template_compare_case(
                repo_root,
                powershell_executable,
                [
                    '第一条 付款条款',
                    '乙方应在验收后30日内付款。'
                ],
                [
                    '第一条 付款条款',
                    '乙方应在验收合格且收到合法有效发票后10个工作日内付款。'
                ],
                {
                    'focus_topics': ['付款'],
                    'max_operations': 2,
                    'mode': 'revision_comment',
                    'min_similarity_for_revision': 0.2
                },
                env=python3_only_env,
                case_subdir='中文目录'
            )
        finally:
            shutil.rmtree(shim_dir, ignore_errors=True)
        assert linux_style_summary['template_compare_applied'] is True, json.dumps(linux_style_summary, ensure_ascii=False, indent=2)
        assert linux_style_summary['template_generated_operations'] == 1, json.dumps(linux_style_summary, ensure_ascii=False, indent=2)
        assert linux_style_docx_summary['comment_range_start_count'] == 1, json.dumps(linux_style_docx_summary, ensure_ascii=False, indent=2)
        assert len(linux_style_docx_summary['visible_paragraph_texts']) == 2, json.dumps(linux_style_docx_summary, ensure_ascii=False, indent=2)
        assert 'review_summary_lines' in linux_style_summary and len(linux_style_summary['review_summary_lines']) >= 1, json.dumps(linux_style_summary, ensure_ascii=False, indent=2)

        long_source_paragraphs = [
            '合同目的',
            '双方拟就项目合作建立长期服务关系。',
            '项目范围',
            '服务范围包括实施、培训、运维支持。',
            '付款安排',
            '甲方应在项目验收后30日内支付服务费。',
            '双方应保证所提供资料真实、完整、有效。',
            '服务期间，乙方应采取合理措施确保系统稳定运行。',
            '争议处理',
            '争议由双方协商解决，协商不成的，向原告所在地人民法院起诉。',
            '通知与送达',
            '双方确认本合同载明地址为有效送达地址。'
        ]
        long_template_paragraphs = [
            '合同目的',
            '双方拟就项目合作建立长期服务关系。',
            '项目范围',
            '服务范围包括实施、培训、运维支持。',
            '付款安排',
            '甲方应在项目验收合格且收到合法有效发票后10个工作日内支付服务费。',
            '保密',
            '双方应对在合作过程中获悉的商业秘密承担保密义务，除法律法规另有规定外不得向第三方披露。',
            '双方应保证所提供资料真实、完整、有效。',
            '服务期间，乙方应采取合理措施确保系统稳定运行，并建立持续监控机制。',
            '争议处理',
            '争议由双方协商解决，协商不成的，任一方可向合同签订地有管辖权的人民法院起诉。',
            '通知与送达',
            '双方确认本合同载明地址与电子邮箱均为有效送达地址。'
        ]
        phase5_summary, phase5_docx_summary = run_template_compare_case(
            repo_root,
            powershell_executable,
            long_source_paragraphs,
            long_template_paragraphs,
            {
                'focus_topics': ['付款', '保密', '争议解决'],
                'max_operations': 8,
                'mode': 'revision_comment',
                'min_similarity_for_revision': 0.2,
                'min_alignment_confidence': 0.2,
                'allow_missing_clause_insert': True,
                'missing_clause_mode': 'revision_comment'
            }
        )
        assert phase5_summary['template_compare_applied'] is True, json.dumps(phase5_summary, ensure_ascii=False, indent=2)
        assert phase5_summary['template_generated_operations'] >= 3, json.dumps(phase5_summary, ensure_ascii=False, indent=2)
        assert len(phase5_summary['alignment_reasons_summary']) >= 2, json.dumps(phase5_summary, ensure_ascii=False, indent=2)
        assert len(phase5_summary['missing_reasons_summary']) >= 1, json.dumps(phase5_summary, ensure_ascii=False, indent=2)
        assert 'review_summary_text' in phase5_summary and len(phase5_summary['review_summary_lines']) >= 4, json.dumps(phase5_summary, ensure_ascii=False, indent=2)
        assert len(phase5_docx_summary['visible_paragraph_texts']) >= 10, json.dumps(phase5_docx_summary, ensure_ascii=False, indent=2)

        similarity_guard_summary, similarity_guard_docx_summary = run_template_compare_case(
            repo_root,
            powershell_executable,
            [
                '第一条 付款条款',
                '甲方应在每月月底前付款。'
            ],
            [
                '第一条 付款条款',
                '甲方应在合同签署后三个工作日内一次性支付全部合同价款，并另行承担保证金与税费义务。'
            ],
            {
                'focus_topics': ['付款'],
                'max_operations': 2,
                'mode': 'revision_comment',
                'min_similarity_for_revision': 0.95,
                'min_alignment_confidence': 0.1
            }
        )
        assert similarity_guard_summary['template_generated_operations'] == 1, json.dumps(similarity_guard_summary, ensure_ascii=False, indent=2)
        assert similarity_guard_summary['revisions_applied'] == 0, json.dumps(similarity_guard_summary, ensure_ascii=False, indent=2)
        assert similarity_guard_summary['comments_applied'] == 1, json.dumps(similarity_guard_summary, ensure_ascii=False, indent=2)
        assert similarity_guard_docx_summary['ins_count'] == 0, json.dumps(similarity_guard_docx_summary, ensure_ascii=False, indent=2)
        assert similarity_guard_docx_summary['comment_range_start_count'] == 1, json.dumps(similarity_guard_docx_summary, ensure_ascii=False, indent=2)
        assert len(similarity_guard_docx_summary['visible_paragraph_texts']) == 2, json.dumps(similarity_guard_docx_summary, ensure_ascii=False, indent=2)

        multi_missing_summary, multi_missing_docx_summary = run_template_compare_case(
            repo_root,
            powershell_executable,
            [
                '第一条 付款条款',
                '甲方应在验收后30日内付款。'
            ],
            [
                '第一条 付款条款',
                '甲方应在验收合格且收到合法有效发票后10个工作日内付款。',
                '第二条 保密条款',
                '双方应对合作中获悉的商业秘密承担保密义务。',
                '第三条 违约责任',
                '违约方应赔偿守约方因此遭受的全部损失。'
            ],
            {
                'focus_topics': ['付款', '保密', '违约责任'],
                'max_operations': 6,
                'mode': 'revision_comment',
                'min_similarity_for_revision': 0.2,
                'min_alignment_confidence': 0.2,
                'allow_missing_clause_insert': True,
                'missing_clause_mode': 'revision_comment'
            }
        )
        assert multi_missing_summary['template_generated_operations'] >= 3, json.dumps(multi_missing_summary, ensure_ascii=False, indent=2)
        assert len(multi_missing_summary['missing_reasons_summary']) == 2, json.dumps(multi_missing_summary, ensure_ascii=False, indent=2)
        assert len(multi_missing_docx_summary['visible_paragraph_texts']) >= 6, json.dumps(multi_missing_docx_summary, ensure_ascii=False, indent=2)

        multiparagraph_guard_summary, multiparagraph_guard_docx_summary = run_template_compare_case(
            repo_root,
            powershell_executable,
            [
                '第一条 付款条款',
                '甲方应在验收后30日内付款。',
                '逾期付款的，违约金按日万分之五计算。'
            ],
            [
                '第一条 付款条款',
                '甲方应在验收合格且收到合法有效发票后10个工作日内付款。',
                '逾期付款的，违约金按中国人民银行同期贷款市场报价利率上浮50%计算。'
            ],
            {
                'focus_topics': ['付款'],
                'max_operations': 2,
                'mode': 'revision_comment',
                'min_similarity_for_revision': 0.2,
                'min_alignment_confidence': 0.2
            }
        )
        assert multiparagraph_guard_summary['template_generated_operations'] >= 1, json.dumps(multiparagraph_guard_summary, ensure_ascii=False, indent=2)
        assert multiparagraph_guard_summary['revisions_applied'] == 0, json.dumps(multiparagraph_guard_summary, ensure_ascii=False, indent=2)
        assert multiparagraph_guard_summary['comments_applied'] >= 1, json.dumps(multiparagraph_guard_summary, ensure_ascii=False, indent=2)
        assert multiparagraph_guard_docx_summary['ins_count'] == 0, json.dumps(multiparagraph_guard_docx_summary, ensure_ascii=False, indent=2)
        assert multiparagraph_guard_docx_summary['comment_range_start_count'] >= 1, json.dumps(multiparagraph_guard_docx_summary, ensure_ascii=False, indent=2)
        assert len(multiparagraph_guard_docx_summary['visible_paragraph_texts']) >= 2, json.dumps(multiparagraph_guard_docx_summary, ensure_ascii=False, indent=2)

        print('template compare tests passed')
        print(json.dumps({
            'linux_style_generated_operations': linux_style_summary['template_generated_operations'],
            'multi_missing_generated_operations': multi_missing_summary['template_generated_operations'],
            'phase5_generated_operations': phase5_summary['template_generated_operations'],
            'similarity_guard_mode': 'comment_only' if similarity_guard_summary['revisions_applied'] == 0 else 'revision',
            'template_compare_applied': summary['template_compare_applied'],
            'template_generated_operations': summary['template_generated_operations'],
            'phase2_generated_operations': phase2_summary['template_generated_operations'],
            'phase3_generated_operations': phase3_summary['template_generated_operations'],
            'phase4_alignment_summary_size': len(phase4_summary['alignment_reasons_summary']),
            'phase4_missing_summary_size': len(phase4_summary['missing_reasons_summary']),
            'phase5_alignment_summary_size': len(phase5_summary['alignment_reasons_summary']),
            'phase5_missing_summary_size': len(phase5_summary['missing_reasons_summary']),
            'multiparagraph_guard_mode': 'comment_only' if multiparagraph_guard_summary['revisions_applied'] == 0 else 'revision',
        }, ensure_ascii=True, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        print(repr(exc))
        return 1


if __name__ == '__main__':
    sys.exit(main())
