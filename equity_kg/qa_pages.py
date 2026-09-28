from pathlib import Path
import pypdfium2 as pdf
from PIL import Image,ImageDraw
root=Path(__file__).parent/'qa'
d=pdf.PdfDocument(str(root/'report.pdf'))
for i,p in enumerate(d): p.render(scale=1.6).to_pil().save(root/f'report-{i+1}.png')
print('report pages',len(d))
files=list(root.glob('report-*.png'))
files+=list((root/'slides').glob('*.PNG'))+list((root/'slides').glob('*.png'))
for batch in range(0,len(files),6):
    canvas=Image.new('RGB',(1200,1100),'#cccccc');draw=ImageDraw.Draw(canvas)
    for j,path in enumerate(files[batch:batch+6]):
        im=Image.open(path);im.thumbnail((390,510));x=(j%3)*400;y=(j//3)*550
        canvas.paste(im,(x,y+25));draw.text((x+5,y+3),path.stem,fill='black')
    canvas.save(root/f'contact-{batch//6}.png')
