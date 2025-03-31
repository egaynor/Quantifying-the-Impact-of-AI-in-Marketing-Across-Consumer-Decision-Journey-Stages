import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

file_path = "Data NEW Copy.xlsx"  
xls = pd.ExcelFile(file_path)
df = pd.read_excel(xls, sheet_name="Hard Data sorted")  

#WEIGHTING SYSTEM
weights = {
    "Case Study": 0.8,  
    "Estimate": 0.9,  
    "Survey": 0.95, 
    "Statistical analysis": 1.0,  
    None: 0.6  
}

#If multiple data entries from the same data source, average these and then weight to avoid swaying the dataset
df_avg = df.groupby(["SOURCE", "STAGE", "IMPACT", "COLLECTION METHOD"])["CHANGE %"].mean().reset_index()

#Apply weighting
df_avg["WEIGHT"] = df_avg["COLLECTION METHOD"].map(weights)
df_avg["WEIGHTED CHANGE %"] = df_avg["CHANGE %"] * df_avg["WEIGHT"]

summary = df_avg.groupby(["STAGE", "IMPACT"])["WEIGHTED CHANGE %"].mean().reset_index()

plt.figure(figsize=(12, 6))
plt.rcParams['figure.dpi'] = 360
sns.set_theme(style="whitegrid")

ax = sns.barplot(data=summary, x="STAGE", y="WEIGHTED CHANGE %", hue="IMPACT", palette=['#ffb2e1', '#adf0f2'])

for bar in ax.containers:
    ax.bar_label(bar, fmt="%.2f%%", padding=3, fontsize=16, color='black', weight='bold', fontfamily="Times New Roman")

plt.xticks(fontsize=16, fontname="Times New Roman", fontweight="bold")
plt.yticks(fontsize=16, fontname="Times New Roman", fontweight="bold")
plt.xlabel("")
plt.ylabel("Percentage Change", fontsize=16, fontname="Times New Roman", fontweight="bold")
plt.legend(fontsize=12, title_fontsize=14, prop={"family": "Times New Roman","weight": "bold"})
sns.despine(left=True)

plt.show()
print(summary)