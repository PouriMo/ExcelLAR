import sys,json; sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from refine import *
sides=(-1,1,-1,1)
def val(a,b):
    if not (phi0>a>b>phi1): return None
    lam,ph=kin(a,b,sides)
    sg=[np.sign(p*s) for p,s in zip(ph,sides)]
    return lam if len(set(sg))==1 and lam>0 else None
c=[math.radians(133),math.radians(80)]; h=[0.3,0.3]; best=(1e9,None)
for it in range(14):
    for a in np.linspace(c[0]-h[0],c[0]+h[0],81):
        for b in np.linspace(c[1]-h[1],c[1]+h[1],81):
            v=val(a,b)
            if v is not None and v<best[0]: best=(v,(a,b))
    c=list(best[1]); h=[h[0]/3,h[1]/3]
lam,(a,b)=best
print('exact ring kinematic lam =',lam,'phi',np.degrees([a,b]),'H',[hp(phi0,sides[0]).round(2).tolist(),hp(a,sides[1]).round(2).tolist(),hp(b,sides[2]).round(2).tolist(),hp(phi1,sides[3]).round(2).tolist()], 'xP',xP)
json.dump(dict(lam=lam,H=[hp(phi0,sides[0]).tolist(),hp(a,sides[1]).tolist(),hp(b,sides[2]).tolist(),hp(phi1,sides[3]).tolist()]),open('ref.json','w'))
