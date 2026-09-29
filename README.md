# 人工智能课程：企业股权知识表示与风险推理

第6组课堂作业，以虚构企业股权数据演示知识图谱、产生式规则与框架表示，完成控制关系推理、股权穿透和风险线索核查。

## 成果下载

| 材料 | 文件 |
| --- | --- |
| 课堂报告 | [课堂报告_修订版.docx](equity_kg/deliverables/课堂报告_修订版.docx) |
| 课堂PPT（9页，含背景、案例与讲稿备注） | [课堂汇报_修订版.pptx](equity_kg/deliverables/课堂汇报_修订版.pptx) |
| PPT汇报视频（中文合成配音，3分钟内） | [PPT汇报视频_修订版.mp4](equity_kg/deliverables/PPT汇报视频_修订版.mp4) |
| 系统演示视频 | [system_live_recording.mp4](equity_kg/deliverables/system_live_recording.mp4) |
| 汇报讲稿 | [汇报讲稿_修订版.md](equity_kg/deliverables/汇报讲稿_修订版.md) |
| 完整作业压缩包 | [第6组修订预览包.zip](第6组_知识表示与推理_修订预览包.zip) |

修订报告已填写五位组员的姓名、完整学号及分工，梁镇贤为组长；PPT首页列出组员，备注含完整学号与分工。全组共同参与开发、测试和汇报，工作量大致均衡。系统演示视频为真实程序窗口实录及图谱展示。

## 代码与复现

主要代码、虚构数据、实测结果和运行说明位于 [equity_kg](equity_kg/README.md)。建议使用 Python 3.13：

```powershell
cd equity_kg
python -m pip install -r requirements.txt
python main.py
python -m unittest test_project -v
python benchmark.py
```

上方下载链接指向已填写组员信息的修订报告、9页修订PPT、配套讲稿与汇报视频，以及真实系统窗口实录；生成与验证流程见 [修订交付说明](equity_kg/README.md#修订交付与验证边界)。修订成品与完整预览包随仓库同步，姓名、学号和分工不脱敏；QA记录与临时产物仅保留本地。`equity_kg/deliverables/` 保存本地交付成果；`第6组_知识表示与推理_作业包/` 为原作业压缩包的展开副本。根目录另保留课程参考PPT、报告模板与作业要求图片。临时渲染、音频片段及Python缓存不纳入版本控制。

重建文档时，先运行 `build_deliverables.py`，再运行 `revise_ppt.py` 应用补充业务背景的PPT版本；Office导出、配音与视频生成依赖本地Windows环境，详见项目说明。

演示数据均为虚构数据，推理规则为教学简化规则，风险标签只表示待核查线索。
