# 人工智能课程：企业股权知识表示与风险推理

第6组课堂作业，以虚构企业股权数据演示知识图谱、产生式规则与框架表示，完成控制关系推理、股权穿透和风险线索核查。

## 成果下载

| 材料 | 文件 |
| --- | --- |
| 最新课堂报告 | [课堂报告.docx](equity_kg/deliverables/课堂报告.docx) |
| 最新课堂PPT（15页） | [企业股权知识表示与风险推理.pptx](equity_kg/deliverables/企业股权知识表示与风险推理.pptx) |
| PPT讲解视频（本人音色合成配音，151.30秒，含中文字幕） | [企业股权知识表示与风险推理.mp4](equity_kg/deliverables/企业股权知识表示与风险推理.mp4) |
| 配套讲稿 | [本人音色讲解_讲稿.md](equity_kg/deliverables/企业股权知识表示与风险推理_本人音色讲解_讲稿.md) |
| 配套音频与字幕 | [MP3](equity_kg/deliverables/企业股权知识表示与风险推理_本人音色讲解.mp3) · [SRT](equity_kg/deliverables/企业股权知识表示与风险推理_本人音色讲解.srt) |
| 系统运行实录（烧录中文字幕，72.27秒） | [system_live_recording_字幕版.mp4](equity_kg/deliverables/system_live_recording_字幕版.mp4) · [SRT](equity_kg/deliverables/system_live_recording_字幕版.srt) |
| 最新完整交付包（含代码、数据和历史参考材料） | [第6组_知识表示与推理_完整交付包.zip](第6组_知识表示与推理_完整交付包.zip) |

修订报告已填写五位组员的姓名、完整学号及分工，梁镇贤为组长；PPT首页列出组员，备注含完整学号与分工。全组共同参与开发、测试和汇报，工作量大致均衡。系统演示视频为真实程序窗口实录及图谱展示。

2026-10-12 同步最新版交付物。15页PPT与151.30秒讲解视频对应；系统实录另提供21条烧录中文字幕，保留原始实录。早期9页修订PPT及视频、14页暖色增强版PDF和旧讲稿仍在 `equity_kg/deliverables/` 保留作历史参考，不作为最新版配套预览。根目录旧“作业包”和“修订预览包”也仅供历史对照。

## 代码与复现

主要代码、虚构数据、实测结果和运行说明位于 [equity_kg](equity_kg/README.md)。建议使用 Python 3.13：

```powershell
cd equity_kg
python -m pip install -r requirements.txt
python main.py
python -m unittest test_project -v
python benchmark.py
```

上方链接指向最新版报告、15页PPT、配套讲稿与讲解视频，以及带字幕的真实系统窗口实录。交付版本和生成流程见 [项目说明](equity_kg/README.md)。成品保留组员姓名、学号和分工；QA记录、声音参考样本、模型文件及临时产物仅保留本地。完整交付包由 `python equity_kg/package_current_submission.py` 生成，包含 SHA256 文件清单。

重建文档时，先运行 `build_deliverables.py`，再运行 `revise_ppt.py` 应用补充业务背景的PPT版本；Office导出、配音与视频生成依赖本地Windows环境，详见项目说明。

演示数据均为虚构数据，推理规则为教学简化规则，风险标签只表示待核查线索。
