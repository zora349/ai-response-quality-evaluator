from pathlib import Path
import pandas as pd

from deepeval.metrics import GEval
from deepeval.test_case import LLMTestCase, SingleTurnParams

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_PATH = BASE_DIR / "data" / "eval_dataset.csv"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_PATH = OUTPUT_DIR / "deepeval_result.xlsx"

df = pd.read_csv(DATA_PATH)

# 先只跑前5条，避免API调用太多
sample_df = df.head(5).copy()

instruction_metric = GEval(
    name="Instruction Compliance",
    criteria=(
        "Evaluate whether the actual output follows the user's instruction. "
        "Focus on constraints such as number of items, word limit, forbidden words, required format, and whether extra explanations are added."
    ),
    evaluation_steps=[
        "Read the user instruction carefully.",
        "Check whether the actual output satisfies all explicit constraints.",
        "Penalize outputs that violate word limits, item counts, forbidden words, required format, or add unwanted explanations.",
        "Compare the actual output with the expected rule.",
        "Give a score from 0 to 1, where 1 means fully compliant and 0 means completely non-compliant."
    ],
    evaluation_params=[
        SingleTurnParams.INPUT,
        SingleTurnParams.ACTUAL_OUTPUT,
        SingleTurnParams.EXPECTED_OUTPUT,
    ],
    threshold=0.7,
)

results = []

for _, row in sample_df.iterrows():
    test_case = LLMTestCase(
        input=str(row["user_prompt"]),
        actual_output=str(row["model_answer"]),
        expected_output=str(row["expected_rule"]),
    )

    instruction_metric.measure(test_case)

    deepeval_score = instruction_metric.score
    deepeval_reason = instruction_metric.reason
    deepeval_result = "合格" if deepeval_score >= 0.7 else "不合格"

    results.append({
        "id": row["id"],
        "user_prompt": row["user_prompt"],
        "model_answer": row["model_answer"],
        "expected_rule": row["expected_rule"],
        "manual_result": row["manual_result"],
        "manual_error_type": row["error_type"],
        "manual_reason": row["reason"],
        "deepeval_score": deepeval_score,
        "deepeval_result": deepeval_result,
        "deepeval_reason": deepeval_reason,
        "is_consistent": row["manual_result"] == deepeval_result,
    })

result_df = pd.DataFrame(results)

total = len(result_df)
consistent_count = result_df["is_consistent"].sum()
consistency_rate = round(consistent_count / total * 100, 2) if total else 0

summary = pd.DataFrame([
    ["自动评估样本数", total],
    ["人工与自动一致数量", consistent_count],
    ["一致率", f"{consistency_rate}%"],
])

with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    result_df.to_excel(writer, index=False, sheet_name="DeepEval明细")
    summary.to_excel(writer, index=False, header=["指标", "结果"], sheet_name="一致性汇总")

print("DeepEval自动评估完成")
print(f"样本数: {total}")
print(f"一致数量: {consistent_count}")
print(f"一致率: {consistency_rate}%")
print(f"结果已保存: {OUTPUT_PATH}")