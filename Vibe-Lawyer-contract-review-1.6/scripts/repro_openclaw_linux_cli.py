#!/usr/bin/env python3
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from shutil import which
from xml.etree import ElementTree as ET


W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W_NS, 'pr': PR_NS}


def resolve_repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


def resolve_powershell_executable() -> str:
    for candidate in ('pwsh', 'powershell', 'powershell.exe'):
        executable = which(candidate)
        if executable:
            return executable
    raise RuntimeError('未找到可用的 PowerShell 可执行文件（pwsh / powershell / powershell.exe）')


def create_docx_with_paragraphs(output_path: Path, paragraphs: list[str]) -> None:
    temp_dir = Path(tempfile.mkdtemp(prefix='openclaw-docx-'))
    try:
        word_dir = temp_dir / 'word'
        rels_dir = temp_dir / '_rels'
        word_rels_dir = word_dir / '_rels'
        word_dir.mkdir(parents=True, exist_ok=True)
        rels_dir.mkdir(parents=True, exist_ok=True)
        word_rels_dir.mkdir(parents=True, exist_ok=True)

        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

        paragraph_xml: list[str] = []
        for text in paragraphs:
            escaped = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            paragraph_xml.append(f'    <w:p><w:r><w:t>{escaped}</w:t></w:r></w:p>')
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

        (temp_dir / '[Content_Types].xml').write_text(content_types, encoding='utf-8')
        (word_dir / 'document.xml').write_text(document_xml, encoding='utf-8')
        (word_dir / 'settings.xml').write_text(settings_xml, encoding='utf-8')
        (rels_dir / '.rels').write_text(root_rels_xml, encoding='utf-8')
        (word_rels_dir / 'document.xml.rels').write_text(document_rels_xml, encoding='utf-8')

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as archive:
            for file_path in temp_dir.rglob('*'):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(temp_dir).as_posix())
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def build_python3_only_environment() -> tuple[dict[str, str], Path]:
    shim_dir = Path(tempfile.mkdtemp(prefix='python3-only-'))
    proxy_script = shim_dir / 'python3_proxy.py'
    proxy_command = shim_dir / 'python3.cmd'

    proxy_script.write_text(
        '\n'.join([
            'import subprocess',
            'import sys',
            '',
            f'REAL_PYTHON = {json.dumps(str(Path(sys.executable).resolve()))}',
            'args = sys.argv[1:]',
            'for index, arg in enumerate(args):',
            "    if arg == '--instructions' and index + 1 < len(args):",
            '        value = args[index + 1]',
            '        try:',
            "            value.encode('ascii')",
            '        except UnicodeEncodeError:',
            "            print('non-ascii instructions path blocked', file=sys.stderr)",
            '            sys.exit(91)',
            '        break',
            'sys.exit(subprocess.run([REAL_PYTHON, *args]).returncode)',
            ''
        ]),
        encoding='utf-8'
    )
    proxy_command.write_text(
        '\r\n'.join([
            '@echo off',
            f'"{Path(sys.executable).resolve()}" "%~dp0python3_proxy.py" %*',
            ''
        ]),
        encoding='utf-8'
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
    env['PATH'] = os.pathsep.join([str(shim_dir), *filtered_path_entries])
    return env, shim_dir


def load_json_file(path: Path) -> dict:
    return json.loads(path.read_text(encoding='utf-8'))


def build_case_files(case_dir: Path) -> tuple[Path, Path, Path]:
    source_docx = case_dir / '对方合同.docx'
    template_docx = case_dir / '我方模板.docx'
    instructions_json = case_dir / 'openclaw-instructions.json'

    create_docx_with_paragraphs(
        source_docx,
        [
            '第一条 付款条款',
            '乙方应在验收后30日内付款。',
            '第二条 争议解决',
            '争议由双方协商解决，协商不成的，向原告所在地人民法院起诉。'
        ]
    )
    create_docx_with_paragraphs(
        template_docx,
        [
            '第一条 付款条款',
            '乙方应在验收合格且收到合法有效发票后10个工作日内付款。',
            '第二条 争议解决',
            '争议由双方协商解决，协商不成的，任一方可向合同签订地有管辖权的人民法院起诉。'
        ]
    )

    payload = {
        'author': 'OpenClaw Linux Repro',
        'date': '2026-04-02T10:00:00Z',
        'template_compare': {
            'template_path': './我方模板.docx',
            'focus_topics': ['付款', '争议解决'],
            'max_operations': 4,
            'mode': 'revision_comment',
            'min_similarity_for_revision': 0.2
        },
        'operations': []
    }
    instructions_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    return source_docx, template_docx, instructions_json


def resolve_template_path(instructions_path: Path, payload: dict) -> str | None:
    compare_config = payload.get('template_compare') or {}
    template_value = compare_config.get('template_path') or compare_config.get('template')
    if not template_value:
        return None
    candidate = Path(template_value)
    if candidate.is_absolute():
        return str(candidate)
    return str((instructions_path.parent / candidate).resolve())


def infer_expected_comments(payload: dict) -> bool:
    compare_config = payload.get('template_compare') or {}
    if compare_config:
        compare_mode = str(compare_config.get('mode') or 'revision_comment').lower()
        if compare_mode in {'comment', 'revision_comment'}:
            return True
    for operation in payload.get('operations') or []:
        mode = str(operation.get('mode') or '').lower()
        if mode in {'comment', 'revision_comment'}:
            return True
    return False


def resolve_output_path(source_path: Path, output_path: str | None) -> Path:
    if output_path:
        return Path(output_path).resolve()
    return source_path.with_name(f'{source_path.stem}-openclaw-replayed.docx')


def prepare_replay_inputs(
    source_path: str | None,
    instructions_path: str | None,
    output_path: str | None
) -> tuple[str, Path, Path, Path | None, dict, Path | None]:
    if (source_path is None) != (instructions_path is None):
        raise RuntimeError('外部复放模式下必须同时提供 --source 和 --instructions')

    if source_path is None and instructions_path is None:
        work_root = Path(tempfile.mkdtemp(prefix='openclaw-linux-repro-'))
        case_dir = work_root / '中文目录'
        case_dir.mkdir(parents=True, exist_ok=True)
        source_docx, template_docx, instructions_json = build_case_files(case_dir)
        output_docx = resolve_output_path(source_docx, output_path)
        payload = load_json_file(instructions_json)
        return 'sample', source_docx, instructions_json, output_docx, payload, work_root

    source_docx = Path(source_path).resolve()
    instructions_json = Path(instructions_path).resolve()
    if not source_docx.exists():
        raise RuntimeError(f'源文件不存在: {source_docx}')
    if not instructions_json.exists():
        raise RuntimeError(f'instructions 文件不存在: {instructions_json}')
    payload = load_json_file(instructions_json)
    resolved_output = resolve_output_path(source_docx, output_path)
    template_path = resolve_template_path(instructions_json, payload)
    if template_path and not Path(template_path).exists():
        raise RuntimeError(f'template_compare 模板文件不存在: {template_path}')
    return 'replay', source_docx, instructions_json, resolved_output, payload, None


def inspect_reviewed_docx(docx_path: Path) -> dict[str, object]:
    with zipfile.ZipFile(docx_path, 'r') as archive:
        members = set(archive.namelist())
        has_comments_xml = 'word/comments.xml' in members
        has_settings_xml = 'word/settings.xml' in members
        has_document_rels = 'word/_rels/document.xml.rels' in members
        has_track_revisions = False
        comment_count = 0
        has_comment_relationship = False

        if has_settings_xml:
            settings_root = ET.fromstring(archive.read('word/settings.xml'))
            has_track_revisions = settings_root.find('./w:trackRevisions', NS) is not None

        if has_comments_xml:
            comments_root = ET.fromstring(archive.read('word/comments.xml'))
            comment_count = len(comments_root.findall('./w:comment', NS))

        if has_document_rels:
            rels_root = ET.fromstring(archive.read('word/_rels/document.xml.rels'))
            for rel in rels_root.findall('./pr:Relationship', NS):
                if rel.attrib.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments':
                    has_comment_relationship = True
                    break

    return {
        'has_comments_xml': has_comments_xml,
        'has_settings_xml': has_settings_xml,
        'has_track_revisions': has_track_revisions,
        'has_comment_relationship': has_comment_relationship,
        'comment_count': comment_count
    }


def run_repro(
    repo_root: Path,
    keep_temp: bool,
    source_path: str | None,
    instructions_path: str | None,
    output_path: str | None,
    python_mode: str,
    expect_comments: str
) -> int:
    powershell_executable = resolve_powershell_executable()
    mode, source_docx, instructions_json, output_docx, payload, work_root = prepare_replay_inputs(
        source_path=source_path,
        instructions_path=instructions_path,
        output_path=output_path
    )
    template_docx = resolve_template_path(instructions_json, payload)
    if python_mode == 'python3-only':
        env, shim_dir = build_python3_only_environment()
    else:
        env = os.environ.copy()
        shim_dir = None
    generate_script = repo_root / 'scripts' / 'run_generate_review_docx.ps1'
    expected_comments = infer_expected_comments(payload) if expect_comments == 'auto' else expect_comments == 'yes'

    command = [
        powershell_executable,
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        str(generate_script),
        '-Source',
        str(source_docx),
        '-Instructions',
        str(instructions_json),
        '-Output',
        str(output_docx),
    ]

    print('=== OpenClaw Linux CLI Repro ===')
    print(f'mode: {mode}')
    print(f'repo_root: {repo_root}')
    print(f'working_dir: {source_docx.parent}')
    print(f'source_docx: {source_docx}')
    print(f'template_docx: {template_docx}')
    print(f'instructions_json: {instructions_json}')
    print(f'output_docx: {output_docx}')
    print(f'powershell: {powershell_executable}')
    print(f'python_mode: {python_mode}')
    if shim_dir is not None:
        print(f'python_shim_dir: {shim_dir}')
    print(f'expected_comments: {expected_comments}')
    print('command:')
    print(' '.join(json.dumps(part, ensure_ascii=False) for part in command))

    try:
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
        print('\n--- STDOUT ---')
        print(result.stdout.strip() or '<empty>')
        print('\n--- STDERR ---')
        print(result.stderr.strip() or '<empty>')

        if result.returncode != 0:
            print(f'\n复现失败：命令退出码 {result.returncode}')
            return result.returncode

        if not output_docx.exists():
            print('\n复现失败：未生成输出 DOCX')
            return 2

        summary = json.loads(result.stdout)
        docx_summary = inspect_reviewed_docx(output_docx)
        report = {
            'template_compare_applied': summary.get('template_compare_applied'),
            'template_generated_operations': summary.get('template_generated_operations'),
            'staging_reason': summary.get('staging_reason'),
            'template_compare_template': summary.get('template_compare_template'),
            'review_summary_text': summary.get('review_summary_text'),
            'reviewed_docx_checks': docx_summary
        }
        print('\n--- Parsed Summary ---')
        print(json.dumps(report, ensure_ascii=False, indent=2))

        if payload.get('template_compare') and summary.get('template_compare_applied') is not True:
            print('\n复现失败：template_compare 未生效')
            return 3
        if payload.get('template_compare') and int(summary.get('template_generated_operations') or 0) < 1:
            print('\n复现失败：未生成任何 template_compare operation')
            return 4
        if expected_comments and not docx_summary['has_comments_xml']:
            print('\n复现失败：输出 docx 中没有 comments.xml')
            return 5
        if expected_comments and docx_summary['comment_count'] < 1:
            print('\n复现失败：输出 docx 中没有任何批注节点')
            return 6

        print('\n复现成功：已完成 openclaw instructions 复放，并输出了可检查的 reviewed docx。')
        return 0
    finally:
        if shim_dir is not None:
            shutil.rmtree(shim_dir, ignore_errors=True)
        if keep_temp and work_root is not None:
            print(f'\n已保留工作目录：{work_root}')
        elif work_root is not None:
            shutil.rmtree(work_root, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(description='模拟 Linux 云环境下 openclaw 的 CLI 落版调用，并支持复放外部 instructions.json。')
    parser.add_argument('--repo-root', default=str(resolve_repo_root()))
    parser.add_argument('--keep-temp', action='store_true')
    parser.add_argument('--source')
    parser.add_argument('--instructions')
    parser.add_argument('--output')
    parser.add_argument('--python-mode', choices=('python3-only', 'actual'), default='python3-only')
    parser.add_argument('--expect-comments', choices=('auto', 'yes', 'no'), default='auto')
    args = parser.parse_args()
    return run_repro(
        Path(args.repo_root).resolve(),
        keep_temp=args.keep_temp,
        source_path=args.source,
        instructions_path=args.instructions,
        output_path=args.output,
        python_mode=args.python_mode,
        expect_comments=args.expect_comments
    )


if __name__ == '__main__':
    raise SystemExit(main())
