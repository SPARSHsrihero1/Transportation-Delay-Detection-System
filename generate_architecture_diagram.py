import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches

base_dir = r"d:\UPES 2nd YEar\ML\ML PROJECT"
fig_dir = os.path.join(base_dir, "figures")
os.makedirs(fig_dir, exist_ok=True)

fig, ax = plt.subplots(figsize=(12, 6.5), dpi=300)
fig.patch.set_facecolor('#F8FAFC')
ax.set_facecolor('#F8FAFC')

# Style helper
def draw_box(ax, x, y, w, h, title, subtitle, bg_col, edge_col, icon=""):
    rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.12",
                                  facecolor=bg_col, edgecolor=edge_col, lw=2.2, zorder=3)
    ax.add_patch(rect)
    ax.text(x + w/2, y + h*0.68, f"{icon} {title}", ha='center', va='center',
            fontsize=11.5, fontweight='bold', color=edge_col, zorder=4)
    ax.text(x + w/2, y + h*0.32, subtitle, ha='center', va='center',
            fontsize=9.2, color='#2D3748', zorder=4, linespacing=1.3)

# 1. Top Sequential Pipeline
draw_box(ax, 0.4, 4.6, 3.2, 1.4, "Raw Bus Telemetry", "3,000 trips on 6 real routes\nGPS coords, rain, chokepoints", "#FFFFFF", "#1B365D", "●")
draw_box(ax, 4.4, 4.6, 3.2, 1.4, "Data Cleaning", "Median GPS imputation\n99th %ile outlier clipping", "#FFFFFF", "#008080", "●")
draw_box(ax, 8.4, 4.6, 3.2, 1.4, "Feature Engineering", "StandardScaler (Z-score)\nOne-Hot & Ordinal Encoders", "#FFFFFF", "#2B6CB0", "●")

# Top Arrows
ax.annotate('', xy=(4.35, 5.3), xytext=(3.65, 5.3), arrowprops=dict(arrowstyle="->", lw=2.8, color="#A0AEC0"))
ax.annotate('', xy=(8.35, 5.3), xytext=(7.65, 5.3), arrowprops=dict(arrowstyle="->", lw=2.8, color="#A0AEC0"))

# Split Arrow Downwards from Feature Engineering
ax.annotate('', xy=(3.0, 3.5), xytext=(10.0, 4.55),
            arrowprops=dict(arrowstyle="->", lw=2.5, color="#1B365D", connectionstyle="arc3,rad=0.15"))
ax.annotate('', xy=(9.0, 3.5), xytext=(10.0, 4.55),
            arrowprops=dict(arrowstyle="->", lw=2.5, color="#E65100", connectionstyle="arc3,rad=-0.15"))

# Branch 1 (Left - Blue/Navy): Regression
draw_box(ax, 0.8, 2.0, 4.5, 1.3, "Loss & Model Selection", "MSE vs. MAE vs. Huber Loss (Robust)\nOLS Best Fit, Ridge (L2) & Lasso (L1)", "#FFFFFF", "#1B365D", "▶")
draw_box(ax, 0.8, 0.3, 4.5, 1.2, "TASK 1: Exact Travel Duration", "Accurate to within ~4.25 minutes\nR² = 0.913 | Adj R² = 0.910", "#EBF3FA", "#1B365D", "✓")

ax.annotate('', xy=(3.05, 1.55), xytext=(3.05, 1.95), arrowprops=dict(arrowstyle="->", lw=2.8, color="#1B365D"))

# Branch 2 (Right - Orange/Teal): Classification
draw_box(ax, 6.7, 2.0, 4.5, 1.3, "Margin & Optimizer Selection", "Hinge Loss (SVM) vs. Cross-Entropy\nAdam vs. RMSprop vs. Momentum SGD", "#FFFFFF", "#E65100", "▶")
draw_box(ax, 6.7, 0.3, 4.5, 1.2, "TASK 2: Severe Gridlock Alerts", "Flags severe delays (>65 min late)\nROC-AUC = 0.977 | Accuracy = 93.5%", "#FFF5EB", "#E65100", "★")

ax.annotate('', xy=(8.95, 1.55), xytext=(8.95, 1.95), arrowprops=dict(arrowstyle="->", lw=2.8, color="#E65100"))

ax.set_xlim(0, 12)
ax.set_ylim(0, 6.3)
ax.axis('off')
ax.set_title("UPES Transportation Delay System: Dual-Branch Machine Learning Architecture",
             fontsize=14, fontweight='bold', pad=12, color='#1B365D')

plt.tight_layout()
flow_path = os.path.join(fig_dir, "pipeline_architecture_diagram.png")
plt.savefig(flow_path, dpi=300, bbox_inches='tight')
plt.close()
print(f"[SUCCESS] Saved architecture diagram: {flow_path}")
