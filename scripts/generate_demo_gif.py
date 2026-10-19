#!/usr/bin/env python3
"""Generate synthetic ShiftGuard README demo animation. Requires Pillow."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
out=Path(__file__).resolve().parents[1]/"assets"/"shiftguard-demo.gif"
out.parent.mkdir(parents=True,exist_ok=True)
W,H=1000,420
def font(n):
    p="/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    return ImageFont.truetype(p,n) if Path(p).exists() else ImageFont.load_default()
frames=[]
stages=[("01  COLLECT NAMES + ROLES","General-shift priority comes first"),("02  PLAN SOC ROTATIONS","Assign Morning / Afternoon / Night"),("03  VALIDATE COVERAGE","Disclose uncovered role and shift slots"),("04  REVIEW + EXPORT","Excel, CSV and warning summary")]
for step,(title,subtitle) in enumerate(stages):
    im=Image.new("RGB",(W,H),"#091226");d=ImageDraw.Draw(im)
    d.rounded_rectangle((20,20,980,400),radius=20,fill="#14263f",outline="#536b85",width=2)
    d.text((45,45),"SHIFTGUARD 4.0  /  AI-READY STUDIO",font=font(26),fill="#e4c17a")
    d.text((45,100),title,font=font(31),fill="white")
    d.text((45,149),subtitle,font=font(19),fill="#bacce0")
    for r in range(3):
        d.text((50,216+r*50),f"SOC L{1+r//2}   {r+1:02}",font=font(17),fill="#ffffff")
        for c in range(5):
            colors=["#b7d8f8","#bfe9d0","#dfcdfa","#53647b","#53647b"]
            x=235+c*135;y=204+r*50
            color=colors[(r+c)%5]
            if step==1 and r==0 and c==2:color="#e8bc65"
            if step==2 and r==0 and c==2:color="#e88585"
            d.rounded_rectangle((x,y,x+115,y+36),radius=7,fill=color)
            d.text((x+8,y+8),["Morning","Afternoon","Night","Off","Off"][(r+c)%5],font=font(12),fill="#14213b" if color!="#53647b" else "#ffffff")
    d.text((45,371),"Illustrative synthetic data — not a real staffing assessment",font=font(14),fill="#afbdd0")
    frames.append(im)
frames[0].save(out,save_all=True,append_images=frames[1:],duration=900,loop=0,optimize=True)
assert Image.open(out).n_frames==4
print(out,out.stat().st_size)
