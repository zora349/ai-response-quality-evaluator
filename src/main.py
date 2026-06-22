import re

import pandas as pd

from config import DATA_PATH,OUTPUT_PATH,REPORT_PATH
from data_loader import load_dataset
from evaluator import evaluate_row
from output_writer import generate_excel_analysis
from summary_report import generate_report

def main():
    df = load_dataset(DATA_PATH)

    eval_result = df.apply(evaluate_row, axis=1)
    result_df = pd.concat([df, eval_result], axis=1)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    #result_df.to_excel(OUTPUT_PATH, index=False)
    generate_excel_analysis(result_df, OUTPUT_PATH)
    generate_report(result_df)

    print("评估完成")
    print(f"总样本数: {len(result_df)}")
    print(f"合格数量: {len(result_df[result_df['auto_result'] == '合格'])}")
    print(f"不合格数量: {len(result_df[result_df['auto_result'] == '不合格'])}")
    print(f"综合平均分: {result_df['auto_total_score'].mean():.2f}")
    print(f"Excel结果已保存: {OUTPUT_PATH}")
    print(f"项目报告已保存: {REPORT_PATH}")


if __name__ == "__main__":
    main()