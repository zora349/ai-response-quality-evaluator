from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data" / "eval_dataset.csv"
BACKUP_DIR = BASE_DIR / "data" / "eval_dataset_backup.csv"

#读取原始CSV
df = pd.read_csv(DATA_DIR)

#先备份，防止误操作
df.to_csv(BACKUP_DIR, index=False,encoding="utf-8")

#需要新增的评分字段
new_columns = {
    "instruction_score": "",
    "completeness_score": "",
    "format_score": "",
    "naturalness_score": "",
    "accuracy_score": "",
}

#如果字段不存在，就新增；如果已经存在，就不重复加
for col,default_value in new_columns.items():
    if col not in df.columns:
        df[col] = default_value


#保存回放文件
df.to_csv(DATA_DIR, index=False,encoding="utf-8")

print("CSV字段升级完成")
print(f"原文件已备份到: {BACKUP_DIR}")
print("当前字段：")
print(list(df.columns))

