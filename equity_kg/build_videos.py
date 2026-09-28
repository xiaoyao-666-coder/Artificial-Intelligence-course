"""将真实程序输出分页回放，并把PowerPoint导出的页面与讲稿配音合成视频。"""
from pathlib import Path
import subprocess, wave, re, json, sys, shutil

def find_ffmpeg():
    executable = shutil.which("ffmpeg")
    if executable:
        return executable
    try:
        from imageio_ffmpeg import get_ffmpeg_exe
        return get_ffmpeg_exe()
    except ImportError as error:
        raise RuntimeError("Install FFmpeg or imageio-ffmpeg before rebuilding videos") from error

FFMPEG = find_ffmpeg()
from PIL import Image, ImageDraw, ImageFont
ROOT=Path(__file__).parent; QA=ROOT/'qa'; OUT=ROOT/'deliverables'
FONT='C:/Windows/Fonts/msyh.ttc'; BOLD='C:/Windows/Fonts/msyhbd.ttc'
def run(args): subprocess.run(args,check=True,stdout=subprocess.DEVNULL,stderr=subprocess.PIPE)
def video(image,audio,duration,target):
    args=[FFMPEG,'-y','-loop','1','-i',str(image)]
    if audio:args+=['-i',str(audio)]
    args+=['-t',str(duration),'-vf','scale=1600:900:force_original_aspect_ratio=decrease,pad=1600:900:(ow-iw)/2:(oh-ih)/2,format=yuv420p','-r','12','-c:v','libx264','-preset','fast','-crf','24']
    if audio:args+=['-af','atempo=1.08,apad','-c:a','aac','-b:a','128k']
    args+=['-movflags','+faststart',str(target)];run(args)
def combine(parts,target):
    listing=QA/(target.stem+'-concat.txt')
    listing.write_text('\n'.join("file '"+p.as_posix()+"'" for p in parts),encoding='utf-8')
    run([FFMPEG,'-y','-f','concat','-safe','0','-i',str(listing),'-c','copy','-movflags','+faststart',str(target)])

slides=sorted((QA/'slides').glob('*.PNG'),key=lambda p:int(re.search(r'\d+',p.stem).group()))
if not slides:slides=sorted((QA/'slides').glob('*.png'),key=lambda p:int(re.search(r'\d+',p.stem).group()))
assert len(slides)==9,len(slides)
parts=[];durations=[]
for i,slide in enumerate(slides,1):
    audio=QA/f'voice-{i}.wav'
    with wave.open(str(audio)) as w:seconds=w.getnframes()/w.getframerate()/1.08+.8
    part=QA/f'talk-{i}.mp4'
    if '--demo-only' not in sys.argv or not part.exists(): video(slide,audio,seconds,part)
    parts.append(part);durations.append(seconds)
assert sum(durations)<180,sum(durations)
combine(parts,OUT/'PPT汇报视频.mp4')
if '--talk-only' in sys.argv:
    print('PPT video seconds:',sum(durations))
    sys.exit(0)

# Capture a fresh process execution. The movie is explicitly a readable replay,
# not claimed to be an operating-system screen recording.
proc=subprocess.run([sys.executable,'-X','utf8','-u',str(ROOT/'main.py')],cwd=ROOT,capture_output=True,text=True,encoding='utf-8',check=True)
(ROOT/'output/video_demo_run.txt').write_text(proc.stdout,encoding='utf-8')
chunks=re.split(r'={20,}\n',proc.stdout)
blocks=[]
for j in range(1,len(chunks)-1,2):
    head=chunks[j].strip();lines=[x.rstrip() for x in chunks[j+1].splitlines() if x.strip()]
    # Long proof trees are paginated without deleting their content.
    for start in range(0,len(lines),15): blocks.append((head,lines[start:start+15]))
parts=[]
for idx,(head,lines) in enumerate(blocks):
    im=Image.new('RGB',(1600,900),'#101C2C');d=ImageDraw.Draw(im)
    d.text((48,30),'系统代码运行演示  |  程序真实输出回放',font=ImageFont.truetype(BOLD,32),fill='#76DECD')
    d.text((48,88),'python -X utf8 -u main.py   ·   exit code 0',font=ImageFont.truetype(FONT,22),fill='#A7B5CA')
    d.text((48,145),head,font=ImageFont.truetype(BOLD,28),fill='white')
    for row,line in enumerate(lines):
        line=line.replace('⇐','<=')
        font_size=22
        while ImageFont.truetype(FONT,font_size).getlength(line)>1500 and font_size>15:font_size-=1
        d.text((48,213+row*37),line,font=ImageFont.truetype(FONT,font_size),fill='#E4EBF4')
    d.text((48,840),f'虚构教学数据；风险线索不等于法律认定。   {idx+1}/{len(blocks)+2}',font=ImageFont.truetype(FONT,21),fill='#A7B5CA')
    frame=QA/f'demo-{idx+1}.png';im.save(frame)
    part=QA/f'demo-{idx+1}.mp4';video(frame,None,8,part);parts.append(part)
for image in ('kg_raw.png','kg_reasoned.png'):
    part=QA/f'{image}.mp4';video(ROOT/'output'/image,None,10,part);parts.append(part)
combine(parts,OUT/'系统运行demo.mp4')
(QA/'video_info.json').write_text(json.dumps({'talk_seconds':sum(durations),'demo_seconds':len(blocks)*8+20,'talk_slides':len(slides),'demo_process_exit':proc.returncode},indent=2),encoding='utf-8')
print('Video durations:',sum(durations),len(blocks)*8+20)
