import matplotlib.pyplot as plt

fig, ax = plt.subplots(figsize=(6, 4))
ax.axis('off')

cell_text = [
    ["SPUL", "val 1", ""],
    ["Target Spul", "val 2", ""]
]
col_labels = ["", "Max", "Name"]

table = ax.table(cellText=cell_text, colLabels=col_labels, loc='center', cellLoc='center', bbox=[0, 0, 1, 1])

# add independent text
ax.text(0.833, 0.333, "GRAPH NAME\nLine 2\nLine 3", ha='center', va='center', fontweight='bold')

fig.savefig("test_table.png")
