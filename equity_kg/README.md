# 企业股权知识表示与风险推理

## 交付文件

`deliverables/` 包含课堂报告（按所附模板十部分填写）、9页可编辑PPT、中文合成配音的PPT汇报视频、系统运行demo视频及汇报讲稿。公开原稿保留姓名、学号和分工占位，本地修订副本可按组员资料填写并重新验证。

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

`build_deliverables.py` 用python-docx与python-pptx生成文件；按当前环境使用Codex bundled Python运行。`office_qa.ps1` 通过本机Word和PowerPoint导出PDF及页面图，`make_audio.ps1` 使用Windows Huihui语音，`build_videos.py` 使用FFmpeg合成视频。视频脚本使用当前 Python 解释器；FFmpeg 优先取 PATH，缺失时使用 imageio-ffmpeg 提供的可执行文件。QA文件位于 `qa/`，不作为作业正文提交。运行 `build_deliverables.py` 会重写报告中的个人信息占位，请先备份自行填写的版本。

## 修订交付与验证边界

保留原报告、PPT和视频；`课堂报告_修订版.docx` 是报告修订副本，明确区分历史基准结果与本次功能复测。修订报告已通过隔离提取的 LibreOffice 25.8.7 转换，使用标准 render_docx.py 加 Windows 文件 URI 适配器渲染为5页图片，并逐页检查通过。此路径避开了本机 Word 导出挂起问题，不修改系统 Office 安装或默认关联。填写个人信息后仍应复查分页。

`deliverables/system_live_recording.mp4` 替代旧分页回放作为推荐系统演示：录制专用窗口内真实 `main.py --step` 进程的输出，自动发送回车推进，并展示新生成的结果图谱；无配音。原 `系统运行demo.mp4` 保留作参考。使用 `python record_live_demo.py` 可重录（Windows 图形会话，录制期间不要遮挡窗口）。

`package_submission.py` 现在使用当前解释器，验证两段选定视频的完整解码、PPT视频音轨和三分钟限制，并重跑单元测试。采用明确文件清单，排除 `.venv`、QA 和依赖目录，生成带 SHA256 清单的独立“修订预览包”，不覆盖原提交ZIP。运行：

```powershell
python -m pip install -r requirements-deliverables.txt
python package_submission.py
```

修订预览包现在选用 `课堂汇报_修订版.pptx`、`汇报讲稿_修订版.md` 和 `PPT汇报视频_修订版.mp4`，替换包内原PPT及其配音。修订副本沿用原稿9页结构和字体，补充规则编号与来源备注，明确区分控制关系、经济持股、原基准性能数据和功能复测，保留原文件供对照。

新增 `edit_presentation.ps1` 通过本机PowerPoint精确替换既有文本并保存修订副本、逐页PNG和演讲者备注；可选的 `notes_extra` 保存组员分工等补充备注，保持配音正文独立。`make_presentation_audio.ps1` 用Windows Huihui生成对应配音；`build_revision_video.py` 用FFmpeg逐页合成带音轨的视频。编辑时保留原文本框的自动适应设置，因此文字变化可能改变文本框高度；校验限定为已编辑文本框，位置和宽度仍须一致，最终页面需人工确认无碰撞或截断。脚本保留原字体和字号，并显式固定中文字体以防Office主题替换。

该修订路径依赖本机PowerPoint、Windows SAPI和FFmpeg，以及本地 `qa/ppt-revision/content.json` 与审核记录；不需要重新生成报告或运行旧版全量生成脚本。默认执行顺序如下：

```powershell
.\edit_presentation.ps1 -ContentPath .\qa\ppt-revision\content.json
.\make_presentation_audio.ps1 -ContentPath .\qa\ppt-revision\content.json
python build_revision_video.py
# 逐页人工检查导出的PNG后，再执行本地结构及视频对应校验：
python qa/ppt-revision/verify_final.py
python package_submission.py
```

打包前检查9页结构、备注来源、逐页视频画面、讲稿与内容计划的SHA256关联，以及两段视频完整解码、汇报视频音轨和三分钟限制；同时运行11项功能测试并核对ZIP内各文件摘要。报告审核原先绑定单个固定文件摘要，填写成员信息后会失效；现在改读本地 `qa/revision/report-visual-qa.json`，核对 `report_sha256`、`page_count`、连续完整的 `pages_reviewed` 和 `structure_passed`，仍要求每次内容变更后重新渲染检查。可选的 `member_information_filled` 决定包内提示是否仍要求填写成员信息；个人资料仅保存在本地报告、PPT及QA中。验证覆盖重新填写后的报告、补充备注、两段视频和过期审核拒绝；不包含真人口述、异机字体兼容性或真实业务准确率。改动PPT或讲稿后须重建相应配音、视频和审核记录。
公开同步范围：代码、合成示例和复现说明；本地修订报告、系统实录、修订预览ZIP及QA记录不随本次代码提交上传。打包脚本依赖这些本地产物，纯代码克隆后不能直接生成相同预览包。
