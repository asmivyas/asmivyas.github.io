# ------------------------------------------------------
# 🎨 Sustainable Fashion Visuals Generator
# Creates 4 colorful charts from your Excel data
# ------------------------------------------------------

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from sklearn.linear_model import LinearRegression

# === Load Excel file ===
# Using the full path to make sure Python finds your file
file_path = r"C:\Users\Hp Elitebook\Desktop\sustainable-fashion-dashboard\Sustainable_Fashion_Europe_Data Final.xlsx"
xls = pd.ExcelFile(file_path)

# Load specific sheets
df_t = pd.read_excel(xls, "Transparency_2025")
df_b = pd.read_excel(xls, "BrandValue_2010_2024")
df_u = pd.read_excel(xls, "UK_Behaviors_2022_2024")

# Create output folder for images
out_dir = "new_visuals"
os.makedirs(out_dir, exist_ok=True)

# ------------------------------------------------------
# 5️⃣ Average Transparency by Brand Group
# ------------------------------------------------------
groups = {
    "H&M": "H&M Group",
    "Intimissimi": "Calzedonia Group",
    "Calzedonia": "Calzedonia Group",
    "Tezenis": "Calzedonia Group",
    "Zara": "Inditex",
    "Massimo Dutti": "Inditex"
}

df_t["Group"] = df_t["Brand"].map(groups).fillna("Other")
avg = df_t.groupby("Group")["Transparency_Index_2025"].mean()

plt.figure(figsize=(7, 5))
plt.bar(avg.index, avg.values, color=["#FF6F61", "#4DB6AC", "#FFD54F", "#7986CB"])
plt.title("Average Transparency by Brand Group (2025)", fontsize=13, weight='bold')
plt.ylabel("Transparency Index (%)")
plt.grid(axis="y", linestyle="--", alpha=0.4)
plt.tight_layout()
plt.savefig(f"{out_dir}/avg_transparency.png")
plt.close()

# ------------------------------------------------------
# 6️⃣ Year-over-Year Growth (Zara vs H&M)
# ------------------------------------------------------
df_b["Zara_Growth"] = df_b["Zara_BrandValue_USD_Million"].pct_change() * 100
df_b["HM_Growth"] = df_b["H&M_BrandValue_USD_Million"].pct_change() * 100

plt.figure(figsize=(8, 5))
plt.plot(df_b["Year"], df_b["Zara_Growth"], "o-", color="#42A5F5", label="Zara")
plt.plot(df_b["Year"], df_b["HM_Growth"], "o-", color="#AB47BC", label="H&M")
plt.title("Year-over-Year Brand Value Growth (%)", fontsize=13, weight='bold')
plt.xlabel("Year")
plt.ylabel("Growth (%)")
plt.legend()
plt.grid(True, linestyle="--", alpha=0.5)
plt.tight_layout()
plt.savefig(f"{out_dir}/brand_growth.png")
plt.close()

# ------------------------------------------------------
# 7️⃣ Change in UK Sustainable Behaviors (2022 → 2024)
# ------------------------------------------------------
df_u["Change_%"] = df_u["2024_%"] - df_u["2022_%"]
df_u = df_u.sort_values("Change_%", ascending=True)

plt.figure(figsize=(8, 5))
plt.barh(df_u["Behavior"], df_u["Change_%"], color="#66BB6A")
plt.title("Change in UK Sustainable Behaviors (2022–2024)", fontsize=13, weight='bold')
plt.xlabel("Change in % (2024 vs 2022)")
plt.tight_layout()
plt.savefig(f"{out_dir}/behavior_change.png")
plt.close()

# ------------------------------------------------------
# 8️⃣ Forecast Brand Values (2025–2026)
# ------------------------------------------------------
yrs = df_b["Year"].values.reshape(-1, 1)
future = np.array([[2025], [2026]])

plt.figure(figsize=(8, 5))
for brand, col in [("Zara", "#42A5F5"), ("H&M", "#AB47BC")]:
    model = LinearRegression()
    model.fit(yrs, df_b[f"{brand}_BrandValue_USD_Million"])
    pred = model.predict(future)
    plt.plot(df_b["Year"], df_b[f"{brand}_BrandValue_USD_Million"], "o-", label=f"{brand} Actual", color=col)
    plt.plot([2025, 2026], pred, "x--", color=col, label=f"{brand} Forecast")

plt.title("Forecasted Brand Values (2025–2026)", fontsize=13, weight='bold')
plt.xlabel("Year")
plt.ylabel("Brand Value (USD Million)")
plt.legend()
plt.grid(True, linestyle="--", alpha=0.5)
plt.tight_layout()
plt.savefig(f"{out_dir}/brand_forecast.png")
plt.close()

# ------------------------------------------------------
# ✅ Done
# ------------------------------------------------------
print("\n✅ All 4 colorful visuals created inside the folder 'new_visuals'!")
print("Files generated:")
print("  • avg_transparency.png")
print("  • brand_growth.png")
print("  • behavior_change.png")
print("  • brand_forecast.png")
