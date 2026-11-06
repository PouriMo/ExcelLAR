"""Minimum thickness of a semicircular arch for radial and vertical joints (self-weight only)."""
import sys,os; sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from lbla import *
from models import write_csv
from scipy.optimize import linprog
R=1000.;D=1000.;gam=2e-5
def pt(r,a): return np.array([r*math.cos(a),r*math.sin(a)])
def radial(t,n):
    ri,re=R-t/2,R+t/2; ph=np.linspace(math.pi,0,n+1); P=[]
    for j in range(n):
        a,b=ph[j],ph[j+1]; P.append(np.array([pt(ri,a),pt(re,a),pt(re,b),pt(ri,b)]))
    return P
def vertical(t,n,m=60):
    ri,re=R-t/2,R+t/2; xmax=0.995*ri
    xs=np.linspace(-xmax,xmax,n-1)
    yi=lambda x:math.sqrt(ri*ri-x*x); ye=lambda x:math.sqrt(re*re-x*x)
    P=[]
    # left end block: base, extrados polyline up to the first joint, joint, intrados polyline down
    ae=math.acos(xs[0]/re); ai=math.acos(xs[0]/ri)
    ext=[pt(re,a) for a in np.linspace(math.pi,ae,m)]; inn=[pt(ri,a) for a in np.linspace(ai,math.pi,m)]
    P.append(np.array(ext+inn))
    for j in range(n-2):
        x0,x1=xs[j],xs[j+1]
        P.append(np.array([[x0,yi(x0)],[x0,ye(x0)],[x1,ye(x1)],[x1,yi(x1)]]))
    q=P[0].copy(); q[:,0]*=-1; P.append(q[::-1])
    return P
def fix(P):
    # numerical identity of shared vertices
    return P
def feasible(P):
    n=len(P); w=np.array([gam*D*centroid(p)[1] for p in P])
    m=dict(polys=P,w=w,fx=np.zeros(n),fy=np.zeros(n),nb=n); s=build(m)
    k=s['G'].shape[1]
    r=linprog(np.zeros(k),A_ub=s['Gin'],b_ub=np.zeros(s['Gin'].shape[0]),A_eq=s['G'],b_eq=s['Fw'],bounds=[(None,None)]*k,method='highs')
    return r.status==0
def bis(fn,n,lo=0.08,hi=0.13):
    for _ in range(14):
        c=(lo+hi)/2
        if feasible(fn(c*R,n)): hi=c
        else: lo=c
    return hi
if __name__=='__main__':
    for n in (20,50,100,200,400):
        print('radial',n,round(bis(radial,n),5), 'vertical',round(bis(vertical,n),5),flush=True)
