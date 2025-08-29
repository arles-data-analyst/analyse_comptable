import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Canvas carré transparent
fig, ax = plt.subplots(figsize=(2.56, 2.56), dpi=100)
ax.set_axis_off()
ax.set_xlim(0, 1); ax.set_ylim(0, 1)

# Barres
xs = [0.28, 0.45, 0.62]
hs = [0.45, 0.68, 0.90]
w  = 0.12
cols = ["#1f77b4", "#2ca02c", "#ff7f0e"]  # bleu, vert, orange

for x, h, c in zip(xs, hs, cols):
    ax.add_patch(
        patches.FancyBboxPatch(
            (x - w/2, 0.10), w, h,
            boxstyle="round,pad=0.02,rounding_size=0.02",
            facecolor=c, edgecolor="none"
        )
    )

# Anneaux (arcs)
ax.add_patch(patches.Arc((0.5, 0.5), 0.95, 0.95, theta1=210, theta2=20,  lw=8, color="#ff7f0e"))
ax.add_patch(patches.Arc((0.5, 0.5), 0.95, 0.95, theta1=200, theta2=330, lw=8, color="#1f77b4"))

# Flèche
ax.annotate("", xy=(0.78, 0.82), xytext=(0.35, 0.55),
            arrowprops=dict(arrowstyle="-|>", lw=8, color="#ff7f0e"))

fig.savefig("assets/logo.png", dpi=256, transparent=True, bbox_inches="tight", pad_inches=0)
print("✅ Logo enregistré dans assets/logo.png")
