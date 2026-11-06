"""Writes ExcelLAR data sheets (CSV) and generates the refinement models of case study b."""
import os,sys; sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from lbla import *
HEAD='Block,X [mm],Y [mm],,Interface,Len [mm],Ds [mm],Fi [deg],Depth [mm],Weight [N],,Fx_lam [N],Fy_lam [N],Backfill_Added'
def write_csv(path,polys,w,fx,fy,depth,head=HEAD,extra=None):
    m=dict(polys=polys,w=np.array(w),fx=np.array(fx),fy=np.array(fy),nb=len(polys))
    s=build(m); nb=len(polys); ni=nb+1
    rows=[]
    for j,p in enumerate(polys):
        q=list(map(tuple,p))+[tuple(p[0])]
        for k,(x,y) in enumerate(q):
            fxv=f"{fx[j]:.6f}" if k==0 else ''; fyv=f"{fy[j]:.6f}" if k==0 else ''
            rows.append([f'Block {j+1}',f'{x:.6f}',f'{y:.6f}',fxv,fyv])
    out=[head]
    for i in range(max(len(rows),ni)):
        r=rows[i] if i<len(rows) else ['','','','','']
        if i<ni:
            ds=np.linalg.norm(s['mid'][i+1]-s['mid'][i]) if i<nb else ''
            I=[f'Interface {i+1}',f"{s['t'][i]:.6f}",f'{ds:.6f}' if i<nb else '',f"{math.degrees(s['th'][i]):.6f}",f'{depth:.6f}',f'{w[i]:.9f}' if i<nb else '']
        else: I=['']*6
        fill='No' if i==0 else ''
        out.append(','.join([r[0],r[1],r[2],'']+I+['',r[3],r[4],fill]))
    open(path,'w',newline='\r\n').write('\n'.join(out)+'\n')

def corners(m,j):
    """Block j quadrilateral as (a0,b0,b1,a1): joint j (a0-b0) and joint j+1 (a1-b1), a = one boundary."""
    P=m['polys']; p=P[j]
    return p

def nested(m,k,P0=None):
    """Each block divided into k sub-blocks by straight joints joining equal fractions of its two chords."""
    P=m['polys']; nb=m['nb']
    s=build(m)
    # joints as ordered end points: a = intrados side, b = extrados side
    A=[];B=[]
    C=s['C']
    # use circle centre side: order each joint so that a is closer to a far point below the arch
    O=np.array([np.mean(C[:,0]),C[:,1].min()-1000])
    for i in range(nb+1):
        if i==0: u,v=_ab(P,0,'L')
        elif i==nb: u,v=_ab(P,nb-1,'R')
        else: u,v=shared_edge(P[i-1],P[i])
        if np.linalg.norm(u-O)>np.linalg.norm(v-O): u,v=v,u
        A.append(u);B.append(v)
    polys=[];w=[]
    for j in range(nb):
        area=centroid(P[j])[1]
        for q in range(k):
            s0,s1=q/k,(q+1)/k
            a0,a1=A[j]+(A[j+1]-A[j])*s0,A[j]+(A[j+1]-A[j])*s1
            b0,b1=B[j]+(B[j+1]-B[j])*s0,B[j]+(B[j+1]-B[j])*s1
            p=np.array([a0,a1,b1,b0]); polys.append(p)
            w.append(m['w'][j]*centroid(p)[1]/area)
    return polys,np.array(w)

def _ab(P,j,side):
    return lbla_abut(P,j)
def lbla_abut(P,j):
    import lbla
    return lbla._abut(P[j],P[j+1] if j==0 else P[j-1])

def split_vertical(polys,xP,fy=-1.0):
    """Vertical load on the line x = xP shared by the two sub-blocks whose centroids bracket it."""
    n=len(polys); F=np.zeros(n); xc=np.array([centroid(p)[0][0] for p in polys])
    return xc,F

def load_on_line(polys,j0,j1,xP,F):
    xc=np.array([centroid(polys[i])[0][0] for i in range(j0,j1)])
    fy=np.zeros(len(polys))
    for i in range(len(xc)-1):
        lo,hi=sorted((xc[i],xc[i+1]))
        if lo<=xP<=hi:
            t=(xP-xc[i])/(xc[i+1]-xc[i]); fy[j0+i]=F*(1-t); fy[j0+i+1]=F*t
            return fy
    i=int(np.argmin(np.abs(xc-xP))); fy[j0+i]=F; return fy

