from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import font_manager

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_PATH = BASE_DIR / "data" / "eval_dataset.csv"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
CHART_DIR = OUTPUT_DIR / "charts"
CHART_DIR.mkdir(exist_ok=True)


font_path = "/mnt/c/Windows/Fonts/msyh.ttc"
font_manager.fontManager.addfont(font_path)

font_name = font_manager.FontProperties(fname=font_path).get_name()

plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = [font_name]
plt.rcParams["axes.unicode_minus"] = False

df = pd.read_csv(DATA_PATH)

required_columns = [
    "id",
    "user_prompt",
    "model_answer",
    "expected_rule",
    "manual_result",
    "error_type",
    "reason",
    "instruction_score",
    "completeness_score",
    "format_score",
    "naturalness_score",
    "accuracy_score",
]

missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    raise ValueError(f"缺少字段: {missing_columns}")

total = len(df)
pass_count = (df["manual_result"] == "合格").sum()
fail_count = (df["manual_result"] == "不合格").sum()
pass_rate = round(pass_count / total * 100, 2) if total else 0

error_stats = (
    df["error_type"]
    .value_counts()
    .reset_index()
)
error_stats.columns = ["错误类型", "数量"]

# 评分字段
score_columns = [
    "instruction_score",
    "completeness_score",
    "format_score",
    "naturalness_score",
    "accuracy_score",
]

# 确保评分字段是数字
for col in score_columns:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# 综合得分：五个维度平均分
df["overall_score"] = df[score_columns].mean(axis=1).round(2)

# 评分维度统计
score_summary = pd.DataFrame({
    "评分维度": [
        "指令遵循",
        "内容完整",
        "格式规范",
        "表达自然",
        "事实准确",
        "综合平均",
    ],
    "平均分": [
        round(df["instruction_score"].mean(), 2),
        round(df["completeness_score"].mean(), 2),
        round(df["format_score"].mean(), 2),
        round(df["naturalness_score"].mean(), 2),
        round(df["accuracy_score"].mean(), 2),
        round(df["overall_score"].mean(), 2),
    ]
})

# ========== 图表生成 ==========

# 1. 合格/不合格分布图
pass_fail_stats = df["manual_result"].value_counts()

plt.figure(figsize=(6, 4))
plt.bar(pass_fail_stats.index, pass_fail_stats.values)
plt.title("合格与不合格样本分布")
plt.xlabel("评估结果")
plt.ylabel("样本数量")
plt.tight_layout()
pass_fail_chart_path = CHART_DIR / "pass_fail_chart.png"
plt.savefig(pass_fail_chart_path, dpi=150)
plt.close()


# 2. 错误类型统计图
plt.figure(figsize=(8, 4))
plt.bar(error_stats["错误类型"], error_stats["数量"])
plt.title("错误类型统计")
plt.xlabel("错误类型")
plt.ylabel("数量")
plt.xticks(rotation=30)
plt.tight_layout()
error_type_chart_path = CHART_DIR / "error_type_chart.png"
plt.savefig(error_type_chart_path, dpi=150)
plt.close()


# 3. 评分维度统计图
plt.figure(figsize=(8, 4))
plt.bar(score_summary["评分维度"], score_summary["平均分"])
plt.title("评分维度平均分")
plt.xlabel("评分维度")
plt.ylabel("平均分")
plt.ylim(0, 5)
plt.xticks(rotation=30)
plt.tight_layout()
score_dimension_chart_path = CHART_DIR / "score_dimension_chart.png"
plt.savefig(score_dimension_chart_path, dpi=150)
plt.close()

# 汇总表
summary = pd.DataFrame([
    ["总样本数", total],
    ["合格数量", pass_count],
    ["不合格数量", fail_count],
    ["合格率", f"{pass_rate}%"],
    ["综合平均分", round(df["overall_score"].mean(), 2)],
])

# 低分样本：综合得分低于4分的样本
low_score_samples = df[df["overall_score"] < 4].copy()

excel_path = OUTPUT_DIR / "eval_result.xlsx"
report_path = OUTPUT_DIR / "error_report.md"


with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="评估明细")
    summary.to_excel(writer, index=False, header=["指标", "结果"], sheet_name="汇总")
    error_stats.to_excel(writer, index=False, sheet_name="错误类型统计")
    score_summary.to_excel(writer, index=False, sheet_name="评分维度统计")
    low_score_samples.to_excel(writer, index=False, sheet_name="低分样本")

report = f"""# AI回答质量评估项目报告

## 一、项目概述

本项目对 AI 模型回答进行人工质量评估，从指令遵循、内容完整性、格式规范性、表达自然度等维度判断回答是否合格，并对不合格样本进行错误归因分析。

## 二、数据概况

- 总样本数：{total}
- 合格数量：{pass_count}
- 不合格数量：{fail_count}
- 合格率：{pass_rate}%
- 综合平均分：{round(df["overall_score"].mean(), 2)}

## 三、错误类型分布

{error_stats.to_markdown(index=False)}

## 四、评分维度表现

{score_summary.to_markdown(index=False)}

## 五、初步结论

从当前样本看，模型存在一定比例的不合格回答，主要问题集中在指令违背、格式错误、回答不完整和内容不准确等方面。

从评分维度看，可以进一步观察模型在哪些方面失分更明显。如果指令遵循和格式规范得分较低，说明模型对细粒度限制条件的执行不稳定；如果内容完整性得分较低，说明模型容易遗漏用户要求；如果事实准确性得分较低，则需要重点关注概念解释和事实性回答质量。

## 六、下一步计划

1. 扩展样本量到 50 条以上。
2. 对错误类型进行更细分的归因分析。
3. 接入 DeepEval 自动评估指标。
4. 对人工评估和自动评估结果进行一致性分析。
5. 输出更完整的项目报告，用于简历和面试展示。

## 七、可视化图表

### 1. 合格与不合格样本分布

![合格与不合格样本分布](charts/pass_fail_chart.png)

### 2. 错误类型统计

![错误类型统计](charts/error_type_chart.png)

### 3. 评分维度平均分

![评分维度平均分](charts/score_dimension_chart.png)
"""

report_path.write_text(report, encoding="utf-8")

print("评估完成")
print(f"总样本数: {total}")
print(f"合格数量: {pass_count}")
print(f"不合格数量: {fail_count}")
print(f"合格率: {pass_rate}%")
print(f"综合平均分: {round(df['overall_score'].mean(), 2)}")
print(f"Excel结果已保存: {excel_path}")
print(f"项目报告已保存: {report_path}")
