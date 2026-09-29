"""Validate and create a revision bundle without overwriting the original submission."""
from pathlib import Path
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'deliverables'
QA = ROOT / 'qa' / 'revision'
EXCLUDED = {'.venv', 'venv', 'qa', 'deliverables', '__pycache__', '.git', 'node_modules'}
EXTENSIONS = {'.py', '.ps1', '.md', '.txt', '.csv', '.json', '.html', '.png'}
PRESENTATION = '课堂汇报_修订版.pptx'
TALK_VIDEO = 'PPT汇报视频_修订版.mp4'
TALK_SCRIPT = '汇报讲稿_修订版.md'
REPORT = '课堂报告_修订版.docx'


def validate_report():
    """Bind report approval to a local record instead of one fixed document hash."""
    review = json.loads((QA / 'report-visual-qa.json').read_text(encoding='utf-8-sig'))
    count = review.get('page_count')
    if type(count) is not int or count < 1 or review.get('pages_reviewed') != list(range(1, count + 1)):
        raise RuntimeError('Every report page requires visual review')
    digest = hashlib.sha256((OUT / REPORT).read_bytes()).hexdigest()
    if review.get('report_sha256') != digest or not review.get('structure_passed'):
        raise RuntimeError('Report changed after visual QA; re-render and inspect before packaging')
    return review


def validate_presentation():
    """Require local review evidence tied to the exact revised deck and video."""
    folder = ROOT / 'qa' / 'ppt-revision'
    visual = json.loads((folder / 'visual-qa.json').read_text(encoding='utf-8-sig'))
    video = json.loads((folder / 'video-manifest.json').read_text(encoding='utf-8-sig'))
    sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
    if visual.get('slides_reviewed') != list(range(1, 10)) or not visual.get('structure_passed'):
        raise RuntimeError('All nine revised slides require visual and structural review')
    if sha(OUT / PRESENTATION) != visual['deck_sha256'] or sha(OUT / TALK_SCRIPT) != visual['script_sha256']:
        raise RuntimeError('Presentation or narration changed after QA; review again')
    if video['deck_sha256'] != visual['deck_sha256'] or sha(OUT / TALK_VIDEO) != video['video_sha256']:
        raise RuntimeError('Video does not match the reviewed deck; rebuild it')
    if [entry['slide'] for entry in visual.get('video_frame_checks', [])] != list(range(1, 10)):
        raise RuntimeError('All nine video segments require slide comparison')
    if sha(folder / 'content.json') != visual.get('content_sha256'):
        raise RuntimeError('Content plan changed after visual QA; review again')
    if sha(folder / 'content.json') != video['content_sha256']:
        raise RuntimeError('Narration plan changed after video generation; rebuild it')
    return 'Nine revised slides visually reviewed; editable structure and source notes verified; video hashes matched'


def ffmpeg_path():
    found = shutil.which('ffmpeg')
    if found:
        return found
    from imageio_ffmpeg import get_ffmpeg_exe
    return get_ffmpeg_exe()


def inspect_video(path):
    ffmpeg = ffmpeg_path()
    probe = subprocess.run([ffmpeg, '-hide_banner', '-i', str(path)], capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=30)
    text = probe.stderr
    match = re.search(r'Duration: (\d+):(\d+):(\d+\.\d+)', text)
    if not match or 'Video:' not in text:
        raise RuntimeError(f'Cannot read video metadata: {path.name}')
    h, m, s = map(float, match.groups())
    info = {'seconds': h * 3600 + m * 60 + s, 'audio': 'Audio:' in text, 'video': True}
    decoded = subprocess.run([ffmpeg, '-v', 'error', '-xerror', '-i', str(path), '-f', 'null', '-'], capture_output=True, text=True, timeout=240)
    if decoded.returncode or decoded.stderr.strip():
        raise RuntimeError(f'Video decode failed: {path.name}: {decoded.stderr}')
    info['full_decode_passed'] = True
    return info


def source_files():
    for directory, folders, files in os.walk(ROOT, followlinks=False):
        folders[:] = sorted(name for name in folders if name not in EXCLUDED and not (Path(directory) / name).is_symlink())
        for name in sorted(files):
            p = Path(directory) / name
            rel = p.relative_to(ROOT)
            if not p.is_symlink() and p.suffix in EXTENSIONS and p.name != 'kg_demo.png':
                yield p, (Path('source') / rel).as_posix()


def main():
    QA.mkdir(parents=True, exist_ok=True)
    report_qa = validate_report()
    presentation_qa = validate_presentation()
    qa = {name: inspect_video(OUT / name) for name in (TALK_VIDEO, 'system_live_recording.mp4')}
    if not (0 < qa[TALK_VIDEO]['seconds'] < 180 and qa[TALK_VIDEO]['audio']):
        raise RuntimeError('PPT video must have audio and be under three minutes')
    test = subprocess.run([sys.executable, '-X', 'utf8', '-m', 'unittest', 'test_project', '-v'], cwd=ROOT, capture_output=True, text=True, encoding='utf-8', timeout=120)
    (QA / 'package-tests.txt').write_text(test.stdout + test.stderr, encoding='utf-8')
    if test.returncode:
        raise RuntimeError('Tests failed; no package created')
    qa['report_visual_qa'] = f"PASSED: {report_qa['page_count']} pages rendered using canonical render_docx.py with Windows URI adapter; all pages visually inspected"
    qa['member_information_filled'] = report_qa.get('member_information_filled', False)
    qa['slides'] = presentation_qa
    (QA / 'final_validation.json').write_text(json.dumps(qa, ensure_ascii=False, indent=2), encoding='utf-8')
    member_status = '组员信息与分工已填写' if qa['member_information_filled'] else '个人信息待填写'
    next_step = '请核对组员信息与分工；内容如继续修改，须复查分页并同步更新PPT视频。' if qa['member_information_filled'] else '填写姓名、学号、分工；填写后复查分页；PPT如继续修改，需要同步重录配音视频。'
    notes = f'''修订预览包说明（{member_status}）

{REPORT}：保留模板十部分；区分历史性能实验与本次功能测试。已通过隔离的LibreOffice与标准渲染脚本输出{report_qa['page_count']}页校验图，并逐页检查，未发现文字截断、图文重叠。
{PRESENTATION}：9页可编辑修订副本，保留原稿版式，补充规则编号、来源备注，区分原基准与功能复测；已逐页检查。
{TALK_VIDEO}：与修订PPT及讲稿对应，中文合成配音，{qa[TALK_VIDEO]['seconds']:.2f}秒，小于3分钟；已完整解码验证。
系统运行实录.mp4：{qa['system_live_recording.mp4']['seconds']:.2f}秒，真实运行main.py --step并录制专用窗口；自动发送回车推进，无配音，末尾展示生成的图谱。
{TALK_SCRIPT}：修订PPT的逐页配套讲稿。
source/：代码和复现资料，不含虚拟环境、QA目录或依赖缓存。

提交前：{next_step}
数据全部为教学合成数据，风险标签不构成现实风控或法律结论。性能数值为仓库原基准，不是本次机器的新测量。
'''
    entries = [(OUT / '课堂报告_修订版.docx', '课堂报告_修订版.docx'), (OUT / PRESENTATION, PRESENTATION), (OUT / TALK_VIDEO, TALK_VIDEO), (OUT / 'system_live_recording.mp4', '系统运行实录.mp4'), (OUT / TALK_SCRIPT, TALK_SCRIPT)]
    entries += list(source_files())
    manifest = {name: hashlib.sha256(path.read_bytes()).hexdigest() for path, name in entries}
    dest = ROOT.parent / '第6组_知识表示与推理_修订预览包.zip'
    temporary = dest.with_suffix('.zip.tmp')
    try:
        with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as z:
            for path, name in entries:
                z.write(path, name)
            z.writestr('提交前必读.txt', notes.encode('utf-8-sig'))
            z.writestr('SHA256.json', json.dumps(manifest, ensure_ascii=False, indent=2))
        with zipfile.ZipFile(temporary) as z:
            if z.testzip() is not None:
                raise RuntimeError('ZIP integrity failure')
            for name, digest in manifest.items():
                if hashlib.sha256(z.read(name)).hexdigest() != digest:
                    raise RuntimeError(f'Hash mismatch: {name}')
            if any(any(part in EXCLUDED for part in Path(name).parts) for name in z.namelist()):
                raise RuntimeError('Excluded directory in archive')
        temporary.replace(dest)
    finally:
        if temporary.exists():
            temporary.unlink()
    print(json.dumps(qa, ensure_ascii=False, indent=2))
    print(f'Preview bundle: {dest} ({dest.stat().st_size:,} bytes)')


if __name__ == '__main__':
    main()
