# MD Doc Link icons - Author: MD (Mayank Dholariya)
from PIL import Image, ImageDraw, ImageFont
BLUE=(28,96,160); WHITE=(255,255,255); ORANGE=(245,158,40); PAPER=(250,252,255); LINE=(150,175,205)

def draw(S):
    k=4; s=S*k
    im=Image.new('RGBA',(s,s),(0,0,0,0)); d=ImageDraw.Draw(im)
    d.rounded_rectangle([0,0,s-1,s-1],radius=int(s*0.18),fill=BLUE)
    # document sheet with folded corner
    x0,y0,x1,y1=int(s*.20),int(s*.13),int(s*.66),int(s*.80); f=int(s*.14)
    d.polygon([(x0,y0),(x1-f,y0),(x1,y0+f),(x1,y1),(x0,y1)],fill=PAPER)
    d.polygon([(x1-f,y0),(x1-f,y0+f),(x1,y0+f)],fill=LINE)
    if S>=32:
        for i in range(3):
            yy=int(y0+s*(.24+.11*i)); d.rectangle([x0+int(s*.07),yy,x1-int(s*.08),yy+max(k,int(s*.035))],fill=LINE)
    # chain link (two rounded rings) bottom-right
    w=max(k*2,int(s*.075))
    a=[int(s*.46),int(s*.56),int(s*.76),int(s*.76)]
    b=[int(s*.62),int(s*.68),int(s*.90),int(s*.88)]
    for r in (a,b):
        d.rounded_rectangle(r,radius=int((r[3]-r[1])/2),outline=ORANGE,width=w)
    return im.resize((S,S),Image.LANCZOS)

for S in (16,32,256):
    draw(S).save('MDDocLink_%d.png'%S)
draw(256).save('MDDocLink.ico',sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])

def font(sz,bold=False):
    for p in ['/usr/share/fonts/truetype/dejavu/DejaVuSans%s.ttf'%('-Bold' if bold else '')]:
        try: return ImageFont.truetype(p,sz)
        except Exception: pass
    return ImageFont.load_default()

L=Image.new('RGB',(164,314),BLUE); d=ImageDraw.Draw(L)
for y in range(314):
    c=tuple(int(BLUE[i]*(1-y/314*.45)) for i in range(3)); d.line([(0,y),(164,y)],fill=c)
L.paste(draw(110),(27,40),draw(110))
d.text((82,180),"MD Doc Link",fill=WHITE,font=font(17,True),anchor='mm')
d.text((82,204),"Excel → hyperlinks",fill=(215,230,245),font=font(11),anchor='mm')
d.text((82,222),"for Navisworks",fill=(215,230,245),font=font(11),anchor='mm')
d.text((82,296),"MD (Mayank Dholariya)",fill=(190,210,235),font=font(9),anchor='mm')
L.save('wizard_large.bmp')
Sm=Image.new('RGB',(55,55),WHITE); Sm.paste(draw(48),(3,3),draw(48)); Sm.save('wizard_small.bmp')
