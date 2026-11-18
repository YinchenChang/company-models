import math
def pack(w,h,R=147.0,step=0.5):
    best=0
    for ox in [i*step for i in range(int(w/step))]:
        for oy in [j*step for j in range(int(h/step))]:
            n=0
            x=-R-w+ox
            while x<R:
                y=-R-h+oy
                while y<R:
                    if all((cx*cx+cy*cy)<=R*R for cx in (x,x+w) for cy in (y,y+h)): n+=1
                    y+=h
                x+=w
            best=max(best,n)
    return best
def f(S,r=150): return math.pi*r*r/S - 2*math.pi*r/math.sqrt(2*S)
for name,S,ar in [("H100-ish",2500,1.0),("BW3.3x",2831,1.25),("RB4x",3432,1.25),("RB4.5x",3861,1.25),("RB5x",4290,1.25),("RU8x",6864,1.25),("RU9x",7722,1.25)]:
    h=math.sqrt(S/ar); w=S/h
    print(name,S,round(w,1),round(h,1),pack(w,h,step=2.0),round(f(S),1), round(f(S,147),1))
