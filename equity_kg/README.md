# 企业股权知识表示与风险推理

## 交付文件

`deliverables/` 包含课堂报告（按所附模板十部分填写）、9页可编辑PPT、中文合成配音的PPT汇报视频、系统运行demo视频及汇报讲稿。姓名、学号和分工按要求保留占位。

系统demo是一次真实Python进程输出的逐页回放，并展示实际生成的图谱，不是操作系统桌面录屏；完整运行日志见 `output/video_demo_run.txt`。PPT视频带合成中文配音，demo以程序文字为主。

## 本地运行

建议 Python 3.13，在本目录执行：

```powershell
python -m pip install -r requirements.txt
python main.py
python main.py --step
python -m unittest test_project -v
python benchmark.py
```

`--step` 每一步按回车继续，适合本人另行录制操作演示。`output/kg_reasoned.html` 可直接在浏览器打开和拖动节点。PyVis缺失时仍生成PNG，Windows字体默认微软雅黑。

## 原有工作与续作

接手时已有 `dataset.py`、`main.py`、`kr/`、`reasoning/`、`viz.py` 及图谱输出。已有内容包括虚构股权数据、图谱与谓词转换、框架继承、15条规则、解释链、路径和环路算法，以及SQLite对照接口。未发现报告、汇报PPT、视频、测试与实际效率实验结果。

本次补充 `benchmark.py` 与等价性检查、11项自动化测试、报告、PPT、两个视频及复现说明。修订共同控制交易预警的措辞，将亲属规则限定为明确的配偶教学假设，并防止深度截断后的企业被当作自然人筛查候选。

## 可核对结果

- 16实体、25条原始关系、53条初始事实，15条规则运行17周期后得到135条事实。
- 王建国对恒通科技经济持股34.2%；教学控制规则下可推导控制关系。
- 9条自然人控制关系回写，6家企业带风险标签；循环持股与担保圈各涉及3家，恒通科技出现待核查交易线索。
- 11项测试通过；430实体、792条持股边的DAG中，对20个目标和4个深度逐一确认DFS、递归CTE、逐层JOIN结果多重集合一致。
- `output/benchmark.csv` 为耗时摘要，JSON含每轮样本、版本和单独记录的构建耗时。性能数据只支持本次实现和数据规模，不代表全部图数据库。

## 语义范围

演示数据全部虚构；真实行业案例参考ICIJ Offshore Leaks公开说明页，未导入其真实人员数据，也未声称复现其内部推理系统。共同控制交易不等于未披露交易。配偶一致行动、持股大于50%的控制传递、风险定级都是教学简化规则。25%仅为教学筛查阈值。框架的默认低风险不代表安全。循环股权按简单路径截断，未求解无限路径权益；SQL逐层JOIN只在无环基准中保证与图查询一致。

## 文档和视频重建

`build_deliverables.py` 用python-docx与python-pptx生成文件；按当前环境使用Codex bundled Python运行。`office_qa.ps1` 通过本机Word和PowerPoint导出PDF及页面图，`make_audio.ps1` 使用Windows Huihui语音，`build_videos.py` 使用FFmpeg合成视频。重建视频需修改脚本中的本机Python路径（当前为D:/python/python.exe）。QA文件位于 `qa/`，不作为作业正文提交。运行 `build_deliverables.py` 会重写报告中的个人信息占位，请先备份自行填写的版本。
