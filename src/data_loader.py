import pandas as pd
from config import DATA_PATH


#REQUIRED_COLUMNS = ["instruction", "model_answer"]
def load_dataset(file_path=DATA_PATH):
    """
    读取待评估数据集，并检查必要字段是否存在
    """
    df = pd.read_csv(file_path)

    #missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]

    #if missing_columns:
    #    raise ValueError(f"数据集缺少必要字段: {missing_columns}")

    return df