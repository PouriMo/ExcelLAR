"""Case study b: radial refinement of the circular arch and exact four-hinge reference."""
import sys,os,json; sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from lbla import *
from models import write_csv, load_on_line
from scipy.optimize import minimize
O=np.array([160.462884,-69.197984]); ri,re=150.647984,180.647984
phi0=math.radians(90+62.655756); phi1=math.pi-phi0   # springing joints (left, right)
gd=4.459246/10797.184748   # weight per unit area: same material as the 11-block model
xP=42.43934786575517
yP=46.50911836
phiP=math.atan2(yP-O[1],xP-O[0])   # angular position of the load point
def arch(n):
    ph=np.linspace(phi0,phi1,n+1)
    P=[]
    for j in range(n):
        a,b=ph[j],ph[j+1]
        q=np.array([O+ri*np.array([math.cos(a),math.sin(a)]),O+re*np.array([math.cos(a),math.sin(a)]),
                    O+re*np.array([math.cos(b),math.sin(b)]),O+ri*np.array([math.cos(b),math.sin(b)])])
        P.append(q)
    w=np.array([gd*centroid(p)[1] for p in P])
    return P,w
def sector(a,b):
    """exact area and first moments (about the global origin) of the ring sector a>b"""
    A=(re**2-ri**2)/2*(a-b)
    Sx=(re**3-ri**3)/3*(math.sin(a)-math.sin(b))+O[0]*A   # int x dA
    return A,Sx
def hp(ph,side): 
    r=re if side>0 else ri
    return O+r*np.array([math.cos(ph),math.sin(ph)])
def kin(p2,p3,sides,P=1.0):
    H=[hp(phi0,sides[0]),hp(p2,sides[1]),hp(p3,sides[2]),hp(phi1,sides[3])]
    # rigid link 1 about H1 (w1=1), link 3 about H4, link 2 by closure
    a=H[2]-H[1]; b=H[3]-H[2]; r=-(H[1]-H[0])
    det=a[0]*b[1]-a[1]*b[0]
    w2=(r[0]*b[1]-r[1]*b[0])/det; w3=(a[0]*r[1]-a[1]*r[0])/det
    v0y=H[1][0]-H[0][0]         # vertical velocity of H2
    def Wd(a_,b_,w,oy_vel,ox):   # work of self-weight: -g*int(v_y dA), v_y = oy_vel + w (x-ox)
        A,Sx=sector(a_,b_); return -gd*(oy_vel*A+w*(Sx-ox*A))
    D=Wd(phi0,p2,1.0,0,H[0][0])+Wd(p2,p3,w2,v0y,H[1][0])+Wd(p3,phi1,w3,0,H[3][0])
    # vertical load -P on line xP inside link 1 or 2
    xh2=H[1][0]
    v=(xP-H[0][0]) if p2<phiP else v0y+w2*(xP-xh2)   # the load point decides the link
    L=-P*v
    phis=[1,w2-1,w3-w2,-w3]
    return -D/L,phis
if __name__=='__main__':
    E='../examples/'; os.makedirs(E+'convergence',exist_ok=True)
    rows=[]
    for n in (11,22,55,100,110,220,1100):
        P,w=arch(n)
        if n==11: 
            fy=np.array([0,-1.]+[0]*9)
        else: fy=load_on_line(P,0,n//3,xP,-1.0)
        mm=dict(polys=P,w=w,fx=np.zeros(n),fy=fy,nb=n)
        s=build(mm); r=solve_full(s); o=summary(s,r)
        ht=[]; 
        rat=o['ratio']; idx=[i for i in range(n+1) if rat[i]>0.9999]
        # hinge x on boundary
        xs=[]
        for i in idx:
            u=np.array([-math.sin(s['th'][i]),math.cos(s['th'][i])])
            xs.append(round(float(s['mid'][i][0]),1))
        print(n,'lam %.4f W %.4f load'%(o['lam'],w.sum()),[(int(i+1),round(float(fy[i]),4)) for i in np.nonzero(fy)[0]],'mu %.3f'%o['mu'],'i',[i+1 for i in idx][:8],'xmid',xs[:8],'sgn e',[int(np.sign(o['e'][i])) for i in idx][:8])
        if n<=220: write_csv(E+f'convergence/Arch_{n:03d}blocks.csv',P,w,np.zeros(n),fy,50.0)
    print('exact ring weight',gd*(re**2-ri**2)/2*(phi0-phi1))
