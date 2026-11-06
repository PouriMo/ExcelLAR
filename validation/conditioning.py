import os,sys; sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from lbla import *
# semicircular arch R = 134 mm, radial joints; condition number of G1 for the last, first and crown interface as X2
for n in (11,55,100):
    R=134.;t=30.
    ph=np.linspace(math.pi,0,n+1)
    P=[np.array([[ (R-t/2)*math.cos(a),(R-t/2)*math.sin(a)],[(R+t/2)*math.cos(a),(R+t/2)*math.sin(a)],[(R+t/2)*math.cos(b),(R+t/2)*math.sin(b)],[(R-t/2)*math.cos(b),(R-t/2)*math.sin(b)]]) for a,b in zip(ph[:-1],ph[1:])]
    m=dict(polys=P,w=np.ones(n),fx=np.zeros(n),fy=np.zeros(n),nb=n); s=build(m); G=s['G']
    out=[]
    for k in (n, 0, n//2):
        idx=list(range(3*k,3*k+3)); keep=[i for i in range(G.shape[1]) if i not in idx]
        out.append(np.linalg.cond(G[:,keep]))
    # with row scaling of the moment rows
    print(n,['%.2e'%c for c in out],'det block', np.linalg.det(G[0:3,0:3]), 'ratio last/crown %.2f'%(out[0]/out[2]))
for f in ('11BlocksArch.csv','55BlocksArch.csv','BenchMark_CircleArch.csv','BenchMark 2_CircleArch.csv'):
    m=read_csv('../examples/'+f); s=build(m); G=s['G']; print(f,'cond %.2e'%np.linalg.cond(G[:,:-3]))
