"""Figure 1 (block, interface actions, sign conventions) and Figure 2 (dimensionless strength domains)."""
import sys
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, FancyArrowPatch, Arc

OUT = sys.argv[1]
plt.rcParams.update({"font.family": "DejaVu Sans", "mathtext.fontset": "dejavusans", "font.size": 13})


# ------------------------------------------------------------------ Figure 1
def arrow(ax, p, q, color="#c0392b", lw=2.0, ms=15, style="-|>", zorder=6):
    ax.add_patch(FancyArrowPatch(p, q, arrowstyle=style, mutation_scale=ms, lw=lw, color=color,
                                 shrinkA=0, shrinkB=0, zorder=zorder))


def fig1():
    A0, B0, B1, A1 = map(np.array, ([-1.7, -1.1], [-2.2, 1.3], [2.0, 1.7], [1.7, -0.7]))
    g = 2.7                                     # gap of the exploded view
    C = np.array([A0, B0, B1, A1]).mean(0)
    P0, P1 = (A0 + B0) / 2, (A1 + B1) / 2
    def frame(a, b):
        d = (b - a) / np.linalg.norm(b - a)
        n = np.array([d[1], -d[0]])             # from block j to the next block
        return n, np.array([n[1], -n[0]])       # t = (sin theta, -cos theta)
    n0, t0 = frame(A0, B0); n1, t1 = frame(A1, B1)
    fig, ax = plt.subplots(figsize=(11.5, 5.6))
    left = np.array([A0 - [g + 2.9, -.75], B0 - [g + 2.5, -1.15], B0 - [g, 0], A0 - [g, 0]])
    right = np.array([A1 + [g, 0], B1 + [g, 0], B1 + [g + 2.9, -.95], A1 + [g + 2.7, -.75]])
    ax.add_patch(Polygon(left, closed=True, fc="#fbe0bd", ec="black", lw=1.1))
    ax.add_patch(Polygon(right, closed=True, fc="#fbe0bd", ec="black", lw=1.1))
    ax.add_patch(Polygon([A0, B0, B1, A1], closed=True, fc="#f6b26b", ec="black", lw=1.7))
    ax.text(C[0] + .1, C[1] + .68, "block $j$", ha="center", fontsize=15, fontweight="bold")
    ax.text(left[:, 0].mean() - .2, left[:, 1].mean() + .2, "block $j-1$", ha="center", fontsize=14)
    ax.text(right[:, 0].mean() + .25, right[:, 1].mean() + .15, "block $j+1$", ha="center", fontsize=14)
    ax.text(*(B0 + [-.15, .2]), "joint $j$", ha="center", fontsize=13, color="#444")
    ax.text(*(B1 + [.25, .2]), "joint $j+1$", ha="center", fontsize=13, color="#444")
    # centroid, midpoints, lever arms
    ax.plot(*C, "o", color="black", ms=7, zorder=7)
    ax.text(C[0] - .12, C[1] + .17, r"$C_j$", fontsize=15, ha="right")
    for P, lab, dx_ in ((P0, r"$P_j$", .1), (P1, r"$P_{j+1}$", -.1)):
        ax.plot(*P, "o", color="black", ms=6, zorder=7)
        ax.plot([C[0], P[0]], [C[1], P[1]], color="#333", lw=1, ls=":", zorder=3)
        ax.text(P[0] + dx_, P[1] + .14, lab, fontsize=14, ha="left" if dx_ > 0 else "right")
    L = 1.35
    red, purple, green, blue = "#c0392b", "#7d3c98", "#2e7d32", "#1f6fb2"
    # positive actions on block j
    arrow(ax, P0, P0 - L * n0, red); ax.text(*(P0 - L * n0 + [-.08, .14]), r"$N_j$", color=red, fontsize=15, ha="right")
    arrow(ax, P0, P0 - .95 * t0, red); ax.text(*(P0 - .95 * t0 + [-.14, .02]), r"$V_j$", color=red, fontsize=15, ha="right")
    arrow(ax, P1, P1 + L * n1, red); ax.text(*(P1 + L * n1 + [.08, .14]), r"$N_{j+1}$", color=red, fontsize=15)
    arrow(ax, P1, P1 + .95 * t1, red); ax.text(*(P1 + .95 * t1 + [.14, -.1]), r"$V_{j+1}$", color=red, fontsize=15)
    def moment(P, cw, lab, off):
        ang = np.radians([-150, -30]) if cw else np.radians([-30, -150])   # arc passes above the point
        A, B = [P + .62 * np.array([np.cos(q), np.sin(q)]) for q in ang]
        ax.add_patch(FancyArrowPatch(A, B, connectionstyle="arc3,rad=%s" % ("0.55" if not cw else "-0.55"),
                                     arrowstyle="-|>", mutation_scale=14, lw=1.9, color=purple, zorder=6))
        ax.text(*(P + off), lab, color=purple, fontsize=15, ha="center")
    moment(P0 - .75 * n0 + [0, .62], True, r"$M_j$", np.array([-.62, .62]))    # clockwise on the block
    moment(P1 + .75 * n1 + [0, .62], False, r"$M_{j+1}$", np.array([.82, .68]))    # counter-clockwise
    # loads at the centroid
    arrow(ax, C, C + [0, -1.05], "black"); ax.text(C[0] + .12, C[1] - 1.02, r"$w_j$", fontsize=15)
    arrow(ax, C, C + [.75, -.45], blue); ax.text(C[0] + .8, C[1] - .42, r"$\mathbf{F}_{\mathrm{ref},j}$", color=blue, fontsize=14, va="center")
    # orientation of joint j
    Q = np.array([-4.55, -2.2])
    ax.plot([Q[0], Q[0] + 1.25], [Q[1], Q[1]], color="#555", lw=1, ls="--")
    arrow(ax, Q, Q + 1.05 * n0, green, ms=13); arrow(ax, Q, Q + .9 * t0, green, ms=13)
    ax.text(*(Q + 1.05 * n0 + [.1, .03]), r"$\mathbf{n}_j$", color=green, fontsize=15)
    ax.text(*(Q + .9 * t0 + [.15, -.05]), r"$\mathbf{t}_j$", color=green, fontsize=15)
    th = np.degrees(np.arctan2(n0[1], n0[0]))
    ax.add_patch(Arc(Q, 1.5, 1.5, theta1=0, theta2=th, color=green, lw=1.2))
    ax.text(Q[0] + .85, Q[1] + .02, r"$\theta_j$", color=green, fontsize=14, va="bottom")
    # global axes
    G = np.array([-7.0, -2.75])
    arrow(ax, G, G + [.95, 0], "black", ms=12); arrow(ax, G, G + [0, .95], "black", ms=12)
    ax.text(G[0] + 1.03, G[1] - .08, "x", fontsize=14); ax.text(G[0] - .05, G[1] + 1.03, "y", fontsize=14)
    ax.set_aspect("equal"); ax.axis("off"); ax.set_xlim(-7.5, 7.5); ax.set_ylim(-3.1, 2.35)
    fig.savefig(OUT + "/fig1_block_equilibrium.png", dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)


# ------------------------------------------------------------------ Figure 2
def fig2():
    # rectangular section t x d with strengths fc, ft = 0.04 fc; N0 = fc t d, compression negative
    rt = 0.04
    n = np.linspace(-1, rt, 800)                  # N / N0
    # stress block: resultant -n between the strengths, moment about the mid-depth
    m_f = 0.5 * (n + 1) * (rt - n) / (1 + rt)
    m_i = 0.5 * np.maximum(-n, 0)
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(11.5, 4.3))
    a1.plot(n, m_f, color="black", lw=2.2, label=r"finite $f_c$, $f_t>0$")
    a1.plot(n[n <= 0], m_i[n <= 0], color="#c0392b", lw=2.2, ls="--", label=r"$f_t=0$, $f_c\rightarrow\infty$ (ExcelLAR)")
    a1.set_xlabel(r"$N/N_0$"); a1.set_ylabel(r"$M/(N_0\,t)$")
    a1.set_xlim(-1.02, 0.08); a1.set_ylim(0, .56); a1.grid(alpha=.3); a1.legend(frameon=False, fontsize=11, loc="upper right")
    a1.text(.02, .04, "(a)", transform=a1.transAxes, fontsize=13)
    nn = n[(n < -.005)]
    a2.plot(nn, (0.5 * (nn + 1) * (rt - nn) / (1 + rt)) / (-nn), color="black", lw=2.2, label=r"finite $f_c$, $f_t>0$")
    nf = nn
    a2.plot(nf, 0.5 * (nf + 1), color="#555", lw=1.8, ls="-.", label=r"finite $f_c$, $f_t=0$")
    a2.plot(nf, 0.5 * np.ones_like(nf), color="#c0392b", lw=2.2, ls="--", label=r"$f_t=0$, $f_c\rightarrow\infty$ (ExcelLAR)")
    a2.set_xlabel(r"$N/N_0$"); a2.set_ylabel(r"$e/t$")
    a2.set_xlim(-1.02, 0.02); a2.set_ylim(0, 1.2); a2.grid(alpha=.3); a2.legend(frameon=False, fontsize=11, loc="upper left")
    a2.text(.02, .04, "(b)", transform=a2.transAxes, fontsize=13)
    for a in (a1, a2):
        for s in ("top", "right"):
            a.spines[s].set_visible(False)
    fig.tight_layout()
    fig.savefig(OUT + "/fig2_strength_domains.png", dpi=300, facecolor="white")
    plt.close(fig)


fig1(); fig2()
