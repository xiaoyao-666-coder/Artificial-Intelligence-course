from pathlib import Path
import zipfile, json, subprocess
ROOT=Path(__file__).parent
OUT=ROOT/'deliverables'
qa={}
for file in OUT.glob('*.mp4'):
    info=json.loads(subprocess.check_output(['ffprobe','-v','error','-show_entries','format=duration:stream=codec_name,codec_type,width,height','-of','json',str(file)],text=True))
    qa[file.name]=info
    if file.name.startswith('PPT'): assert float(info['format']['duration'])<180
    subprocess.run(['ffmpeg','-v','error','-i',str(file),'-f','null','-'],check=True)
test=subprocess.run(['D:/python/python.exe','-X','utf8','-m','unittest','test_project','-v'],cwd=ROOT,capture_output=True,text=True,encoding='utf-8')
(ROOT/'output/test_results.txt').write_text(test.stdout+test.stderr,encoding='utf-8')
assert test.returncode==0
(ROOT/'qa/final_validation.json').write_text(json.dumps(qa,ensure_ascii=False,indent=2),encoding='utf-8')
readme='''提交材料说明

1. 课堂报告.docx：5页，按老师模板十部分完成；第6组，姓名、学号、分工保留占位。
2. 课堂汇报.pptx：9页，可编辑，演讲者备注中有讲稿。
3. PPT汇报视频.mp4：带中文合成配音，符合3分钟内要求；已按背景、问题、案例、方法和验证重写。
4. 系统运行demo.mp4：1分32秒，真实程序输出分页回放及结果图谱，无配音。
5. 汇报讲稿.md：可自行修改并重录。
6. source/：代码、数据、运行说明、实测数据与完整输出日志。

请先填写个人信息。
系统demo为程序输出可视化回放，不是桌面操作录屏；如需本人实录，可运行python main.py --step。
'''
(OUT/'提交说明.txt').write_text(readme,encoding='utf-8-sig')
dest=ROOT.parent/'第6组_知识表示与推理_作业包.zip'
with zipfile.ZipFile(dest,'w',zipfile.ZIP_DEFLATED) as z:
    for p in OUT.iterdir():
        if p.is_file():z.write(p,p.name)
    for p in ROOT.rglob('*'):
        rel=p.relative_to(ROOT)
        if not p.is_file() or any(x in rel.parts for x in ('qa','deliverables','__pycache__')):continue
        if p.suffix in ('.py','.ps1','.md','.txt','.csv','.json','.html','.png') and p.name!='kg_demo.png':z.write(p,str(Path('source')/rel))
with zipfile.ZipFile(dest) as z:assert z.testzip() is None
print(dest, dest.stat().st_size)
