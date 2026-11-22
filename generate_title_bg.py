import os
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFilter

base_dir = r"d:\UPES 2nd YEar\ML\ML PROJECT"
fig_dir = os.path.join(base_dir, "figures")
os.makedirs(fig_dir, exist_ok=True)

# Generate 1920x1080 Title Background
fig, ax = plt.subplots(figsize=(16, 9), dpi=120)
fig.patch.set_facecolor('#0B1728')
ax.set_facecolor('#0B1728')

# 1. Gradient Sky (Dusk in Doon Valley)
x = np.linspace(0, 10, 100)
for i in range(50):
    ax.axhspan(i * 0.1, (i + 1) * 0.1, facecolor=plt.cm.cividis(0.1 + i*0.012), alpha=0.35)

# 2. Foothills / Mountain Silhouettes (Dehradun Ghats)
x_m = np.linspace(0, 10, 500)
# Layer 1: Distant Mussoorie Ridge
y_m1 = 2.8 + 0.6 * np.sin(x_m * 1.2) + 0.3 * np.cos(x_m * 2.5) + 0.2 * np.sin(x_m * 4.0)
ax.fill_between(x_m, 0, y_m1, color='#122640', alpha=0.85)

# Layer 2: Mid Ridge (Bidholi / Prem Nagar hills)
y_m2 = 2.1 + 0.4 * np.sin(x_m * 1.8 + 1) + 0.25 * np.cos(x_m * 3.2)
ax.fill_between(x_m, 0, y_m2, color='#0D1D33', alpha=0.92)

# Layer 3: Foreground Hill
y_m3 = 1.4 + 0.3 * np.sin(x_m * 2.1 + 2) + 0.15 * np.cos(x_m * 4.5)
ax.fill_between(x_m, 0, y_m3, color='#081324', alpha=0.98)

# 3. Winding Highway & Light Streaks (Bus Corridor)
t = np.linspace(0, 1, 300)
# Curving road from bottom left through center to hills
road_x = 0.2 + 9.0 * t - 1.5 * np.sin(t * np.pi * 2.5)
road_y = 0.05 + 1.6 * t + 0.2 * np.cos(t * np.pi * 3)

# Road base
ax.plot(road_x, road_y, color='#1C3554', lw=22, alpha=0.8)
ax.plot(road_x, road_y, color='#2A4B75', lw=12, alpha=0.9)

# Golden Headlights & Red Taillight Streaks (Transit Telemetry)
ax.plot(road_x + 0.08, road_y + 0.02, color='#F6AD55', lw=3.5, alpha=0.95)
ax.plot(road_x - 0.08, road_y - 0.02, color='#E53E3E', lw=4.0, alpha=0.95)
ax.plot(road_x, road_y, color='#ECC94B', lw=1.5, alpha=0.8, linestyle='--')

# 4. Tech / Telemetry Nodes overlay (representing GPS beacons along the road)
sample_pts = np.linspace(0.1, 0.9, 12)
for sp in sample_pts:
    idx = int(sp * len(t))
    px, py = road_x[idx], road_y[idx]
    ax.scatter(px, py, s=90, color='#00E5FF', zorder=5, alpha=0.85, edgecolors='white', lw=1.5)
    # Ping circles
    ax.scatter(px, py, s=350, color='#00E5FF', zorder=4, alpha=0.2)

# 5. Dark Overlay Gradient for text readability
# Ensures white 40-44pt text on top is 100% readable
ax.fill_between([0, 10], 0, 5, color='#071220', alpha=0.68)

ax.set_xlim(0, 10)
ax.set_ylim(0, 5)
ax.axis('off')

plt.tight_layout()
out_bg = os.path.join(fig_dir, "title_bg.png")
plt.savefig(out_bg, dpi=120, bbox_inches='tight', pad_inches=0)
plt.close()
print(f"[SUCCESS] Created title background: {out_bg}")
