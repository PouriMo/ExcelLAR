"""Independent re-implementation of the ExcelLAR linear program (static limit analysis)."""
import csv, sys, math, re
import numpy as np
from scipy.optimize import linprog

def read_csv(path):
    rows=list(csv.reader(open(path,encoding='utf-8-sig')))
    hdr=rows[0]; rows=rows[1:]
    polys={}; fx={}; fy={}; fills={}
    for r in rows:
        r=r+['']*(14-len(r))
        s=r[0].strip()
        if s.startswith('Block'):
            b=int(s.split()[1]); polys.setdefault(b,[]).append((float(r[1]),float(r[2])))
            if r[11].strip(): fx[b]=float(r[11])
            if r[12].strip(): fy[b]=float(r[12])
        elif s:
            fills.setdefault(s,[]).append((r[1],r[2]))
    ifs=[]
    for r in rows:
        r=r+['']*(14-len(r))
        if r[4].strip().startswith('Interface'):
            ifs.append(dict(len=float(r[5]), ds=float(r[6]) if r[6].strip() else None, fi=float(r[7]), depth=float(r[8]) if r[8].strip() else None,
                            w=float(r[9]) if r[9].strip() else None))
    nb=max(polys)
    P=[]
    for b in range(1,nb+1):
        p=polys[b]
        if len(p)>1 and abs(p[0][0]-p[-1][0])<1e-9 and abs(p[0][1]-p[-1][1])<1e-9: p=p[:-1]
        P.append(np.array(p))
    return dict(polys=P, ifs=ifs, fx=np.array([fx.get(b,0.) for b in range(1,nb+1)]), fy=np.array([fy.get(b,0.) for b in range(1,nb+1)]), w=np.array([ifs[i]['w'] for i in range(nb)]), nb=nb, other=list(fills))

def centroid(p):
    x,y=p[:,0],p[:,1]; x1,y1=np.roll(x,-1),np.roll(y,-1)
    c=x*y1-x1*y; A=c.sum()/2
    return np.array([((x+x1)*c).sum()/(6*A),((y+y1)*c).sum()/(6*A)]),abs(A)

def shared_edge(a,b,tol=1e-3):
    idx=[]
    for i in range(len(a)):
        j=(i+1)%len(a)
        for k in range(len(b)):
            l=(k+1)%len(b)
            if (np.linalg.norm(a[i]-b[k])<tol and np.linalg.norm(a[j]-b[l])<tol) or (np.linalg.norm(a[i]-b[l])<tol and np.linalg.norm(a[j]-b[k])<tol):
                idx.append((np.linalg.norm(a[i]-a[j]),a[i],a[j]))
    idx.sort(key=lambda t:-t[0])
    return idx[0][1],idx[0][2]

def edge_by_normal(p,fi):
    best=None
    n=np.array([math.cos(math.radians(fi)),math.sin(math.radians(fi))])
    for i in range(len(p)):
        a,b=p[i],p[(i+1)%len(p)]
        L=np.linalg.norm(b-a)
        if L<1e-9: continue
        d=(b-a)/L; s=abs(d@n)
        if best is None or s<best[0]-1e-6: best=(s,a,b)
    return best[1],best[2]

def build(m):
    P=m['polys']; nb=m['nb']; ni=nb+1
    mid=np.zeros((ni,2)); th=np.zeros(ni); t=np.zeros(ni)
    C=np.array([centroid(p)[0] for p in P])
    for i in range(ni):
        if i==0:
            a,b=_abut(P[0],P[1])
            n=np.array([(b-a)[1],-(b-a)[0]]);n/=np.linalg.norm(n)
            if n@(C[0]-(a+b)/2)<0:n=-n
        elif i==nb:
            a,b=_abut(P[-1],P[-2])
            n=np.array([(b-a)[1],-(b-a)[0]]);n/=np.linalg.norm(n)
            if n@(C[-1]-(a+b)/2)>0:n=-n
        else:
            a,b=shared_edge(P[i-1],P[i])
            n=np.array([(b-a)[1],-(b-a)[0]]);n/=np.linalg.norm(n)
            if n@(C[i]-C[i-1])<0:n=-n
        mid[i]=(a+b)/2; th[i]=math.atan2(n[1],n[0]); t[i]=np.linalg.norm(b-a)
    G=np.zeros((3*nb,3*ni))
    for j in range(nb):
        d0=mid[j]-C[j]; d1=mid[j+1]-C[j]; c0,s0=math.cos(th[j]),math.sin(th[j]); c1,s1=math.cos(th[j+1]),math.sin(th[j+1])
        r=3*j;c=3*j
        G[r,c:c+6]=[-c0,-s0,0,c1,s1,0]
        G[r+1,c:c+6]=[-s0,c0,0,s1,-c1,0]
        G[r+2,c:c+6]=[-d0[0]*s0+d0[1]*c0,d0[0]*c0+d0[1]*s0,-1,d1[0]*s1-d1[1]*c1,-d1[0]*c1-d1[1]*s1,1]
    Gin=np.zeros((2*ni,3*ni))
    for i in range(ni):
        Gin[2*i,3*i]=t[i]/2;Gin[2*i,3*i+2]=1;Gin[2*i+1,3*i]=t[i]/2;Gin[2*i+1,3*i+2]=-1
    Fw=np.zeros(3*nb);Fl=np.zeros(3*nb)
    Fw[1::3]=m['w'];Fl[0::3]=-m['fx'];Fl[1::3]=-m['fy']
    return dict(G=G,Gin=Gin,Fw=Fw,Fl=Fl,mid=mid,th=th,t=t,C=C)

def _abut(p,nbr):
    a,b=shared_edge(p,nbr); best=None
    for i in range(len(p)):
        u,v=p[i],p[(i+1)%len(p)]
        if np.linalg.norm(v-u)<1e-6: continue
        if min(np.linalg.norm(u-a),np.linalg.norm(u-b),np.linalg.norm(v-a),np.linalg.norm(v-b))<1e-4: continue
        y=(u[1]+v[1])/2
        if best is None or y<best[0]: best=(y,u,v)
    return best[1],best[2]

def solve_full(s):
    G,Gin=s['G'],s['Gin']; n=G.shape[1]
    Aeq=np.hstack([G,-s['Fl'][:,None]])
    # scale rows
    Aub=np.hstack([Gin,np.zeros((Gin.shape[0],1))])
    r=linprog(np.r_[np.zeros(n),-1.],A_ub=Aub,b_ub=np.zeros(Aub.shape[0]),A_eq=Aeq,b_eq=s['Fw'],bounds=[(None,None)]*(n+1),method='highs')
    return r

def reduced(s,which='last'):
    G=s['G'];n=G.shape[1];
    keep=list(range(n-3)); last=list(range(n-3,n))
    G1=G[:,keep];G2=G[:,last]
    return G1,G2

def summary(s,r):
    X=r.x[:-1].reshape(-1,3); N,V,M=X.T
    e=np.where(np.abs(N)>1e-9*np.abs(N).max(),M/np.where(N==0,1,N),0)
    ratio=np.abs(e)/(s['t']/2)
    return dict(lam=r.x[-1],N=N,V=V,M=M,e=e,ratio=ratio,mu=(np.abs(V/N)).max())

if __name__=='__main__':
    import glob,os
    for f in sorted(glob.glob(sys.argv[1]+'/*.csv')):
        m=read_csv(f); s=build(m); r=solve_full(s)
        if r.status!=0: print(os.path.basename(f),'status',r.status,r.message); continue
        o=summary(s,r)
        h=[i+1 for i in np.where(o['ratio']>0.98)[0]]
        print(f"{os.path.basename(f):32s} nb={m['nb']:3d} lam={o['lam']:9.4f} W={m['w'].sum():9.4f} Wmax={m['w'].max():8.3f} Fref={np.hypot(m['fx'],m['fy']).max():.3f} load={[(i+1,m['fx'][i],m['fy'][i]) for i in np.nonzero(np.hypot(m['fx'],m['fy']))[0]]} mu={o['mu']:.3f} other={len(m['other'])} hinges={h}")
