# 人工智能课程：企业股权知识表示与风险推理

第6组课堂作业，以虚构企业股权数据演示知识图谱、产生式规则与框架表示，完成控制关系推理、股权穿透和风险线索核查。

## 成果下载

| 材料 | 文件 |
| --- | --- |
| 课堂报告 | [课堂报告.docx](equity_kg/deliverables/课堂报告.docx) |
| 课堂PPT（9页，含背景、案例与讲稿备注） | [课堂汇报.pptx](equity_kg/deliverables/课堂汇报.pptx) |
| PPT汇报视频（中文合成配音，3分钟内） | [PPT汇报视频.mp4](equity_kg/deliverables/PPT汇报视频.mp4) |
| 系统演示视频 | [系统运行demo.mp4](equity_kg/deliverables/系统运行demo.mp4) |
| 汇报讲稿 | [汇报讲稿.md](equity_kg/deliverables/汇报讲稿.md) |
| 完整作业压缩包 | [第6组作业包.zip](第6组_知识表示与推理_作业包.zip) |

报告、PPT中的姓名、学号与分工保留占位，提交前自行填写。系统演示视频为真实程序输出的分页回放及图谱展示。

## 代码与复现

主要代码、虚构数据、实测结果和运行说明位于 [equity_kg](equity_kg/README.md)。建议使用 Python 3.13：

```powershell
cd equity_kg
python -m pip install -r requirements.txt
python main.py
python -m unittest test_project -v
python benchmark.py
```

`equity_kg/deliverables/` 保存修订后的交付成果；`第6组_知识表示与推理_作业包/` 为压缩包的展开副本。根目录另保留课程参考PPT、报告模板与作业要求图片。临时渲染、音频片段及Python缓存不纳入版本控制。

重建文档时，先运行 `build_deliverables.py`，再运行 `revise_ppt.py` 应用补充业务背景的PPT版本；Office导出、配音与视频生成依赖本地Windows环境，详见项目说明。

演示数据均为虚构数据，推理规则为教学简化规则，风险标签只表示待核查线索。
