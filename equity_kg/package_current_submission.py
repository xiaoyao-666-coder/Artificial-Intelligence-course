"""Package the current first-assignment files without rebuilding reviewed media.

Uses the standard library and an explicit directory/extension allowlist.
The historical package_submission.py remains available for the old revision.
"""
from hashlib import sha256
import json
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parent
DESTINATION = ROOT.parent / "第6组_知识表示与推理_完整交付包.zip"
DIRECTORIES = ("data", "kr", "reasoning", "output", "deliverables")
EXTENSIONS = {".py", ".ps1", ".md", ".txt", ".csv", ".json", ".html",
              ".png", ".docx", ".pptx", ".pdf", ".mp4", ".mp3", ".srt"}
EXCLUDED = {"qa", ".git", ".venv", "venv", "__pycache__", "node_modules"}
REQUIRED = (
    "课堂报告.docx", "企业股权知识表示与风险推理.pptx",
    "企业股权知识表示与风险推理.mp4",
    "企业股权知识表示与风险推理_本人音色讲解.mp3",
    "企业股权知识表示与风险推理_本人音色讲解.srt",
    "企业股权知识表示与风险推理_本人音色讲解_讲稿.md",
    "system_live_recording_字幕版.mp4", "system_live_recording_字幕版.srt",
)


def main():
    for name in REQUIRED:
        if not (ROOT / "deliverables" / name).is_file():
            raise FileNotFoundError(name)
    candidates = list(ROOT.iterdir())
    for directory in DIRECTORIES:
        candidates.extend((ROOT / directory).rglob("*"))
    files = []
    for path in candidates:
        relative = path.relative_to(ROOT)
        if (not path.is_file() or path.suffix.lower() not in EXTENSIONS
                or any(part in EXCLUDED or part.startswith(".") for part in relative.parts)
                or path.name.startswith("~$")):
            continue
        if path.is_symlink() or not path.resolve().is_relative_to(ROOT):
            raise ValueError(f"Unexpected file outside project: {relative}")
        files.append(path)
    files = sorted(set(files))
    entries = {"equity_kg/" + p.relative_to(ROOT).as_posix(): p.read_bytes() for p in files}
    entries["阅读说明.md"] = (
        "# 第一次作业：企业股权知识表示与风险推理\n\n"
        "当前推荐交付：课堂报告.docx、企业股权知识表示与风险推理.pptx（15页）、"
        "同名讲解视频（151.30秒）及本人音色讲解配套讲稿、MP3、SRT，"
        "以及 system_live_recording_字幕版.mp4（72.27秒）和同名SRT。\n\n"
        "成品均在 equity_kg/deliverables/；早期9页修订版、14页PDF与旧讲稿仅供历史参考。"
        "代码运行说明见 [项目README](equity_kg/README.md)，业务复现不依赖Office或配音模型。\n\n"
        "SHA256.json记录本包除自身外的每个文件摘要。"
        "本次仅打包已有成品，不重新渲染，不变更实验数据。"
        "QA、依赖、模型、声音参考样本和第二次作业不在包内。\n"
    ).encode("utf-8")
    manifest = {name: sha256(data).hexdigest() for name, data in entries.items()}
    entries["SHA256.json"] = (json.dumps(manifest, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    temporary = DESTINATION.with_suffix(".zip.tmp")
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in entries.items():
                archive.writestr(name, data)
        with zipfile.ZipFile(temporary) as archive:
            if archive.testzip() is not None or set(archive.namelist()) != set(entries):
                raise ValueError("ZIP integrity or inventory check failed")
            for name, digest in manifest.items():
                if sha256(archive.read(name)).hexdigest() != digest:
                    raise ValueError(f"SHA256 mismatch: {name}")
        temporary.replace(DESTINATION)
    finally:
        temporary.unlink(missing_ok=True)
    print(json.dumps({"archive": str(DESTINATION), "files_verified": len(manifest),
                      "bytes": DESTINATION.stat().st_size,
                      "sha256": sha256(DESTINATION.read_bytes()).hexdigest(),
                      "all_hashes_match": True}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
