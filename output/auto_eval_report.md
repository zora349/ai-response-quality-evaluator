# AI回复质量自动评估报告

## 一、项目概况

本项目基于 Python 实现 AI 回复质量自动评估。系统从用户指令中动态解析条数、字数、必须包含关键词、禁用词、表格要求、禁止解释等规则，再对模型回复进行多维度评分和错误归因。

## 二、整体结果

- 总样本数：30

- 合格数量：21

- 不合格数量：9

- 合格率：70.00%

- 综合平均分：4.66

## 三、五项指标平均分

- 指令遵循平均分：4.50

- 完整性平均分：4.60

- 格式规范平均分：4.50

- 自然度平均分：4.80

- 准确性平均分：4.90

## 四、错误类型统计

- 无明显错误：19条

- 指令违背：5条

- 表达不符合要求：2条

- 回答不完整：2条

- 格式/长度错误：1条

- 判断错误：1条

## 五、典型错误样本

### 样本 4

- 用户要求：写3条防诈骗提醒，不要解释

- 模型回答：不要转账；核实身份；不点陌生链接。这些方法可以保护你。

- 解析规则：{'item_count': 3, 'max_len_per_item': None, 'max_len_total': None, 'required_keywords': [], 'forbidden_words': [], 'no_explanation': True, 'need_table': False, 'need_simple': False, 'no_exaggeration': False}

- 综合得分：3.2

- 错误类型：指令违背

- 判断理由：条数不符合要求：要求3条，实际4条；疑似存在多余解释：['这些方法']

### 样本 6

- 用户要求：写一个标题，必须包含卧室

- 模型回答：柔光小夜灯，起夜不摸黑

- 解析规则：{'item_count': None, 'max_len_per_item': None, 'max_len_total': None, 'required_keywords': ['卧室'], 'forbidden_words': [], 'no_explanation': False, 'need_table': False, 'need_simple': False, 'no_exaggeration': False}

- 综合得分：4.4

- 错误类型：指令违背

- 判断理由：缺少必须包含关键词：['卧室']

### 样本 10

- 用户要求：写3条建议，不要出现新手

- 模型回答：新手每天练习；复盘错误；整理标准

- 解析规则：{'item_count': 3, 'max_len_per_item': None, 'max_len_total': None, 'required_keywords': [], 'forbidden_words': ['新手'], 'no_explanation': False, 'need_table': False, 'need_simple': False, 'no_exaggeration': False}

- 综合得分：4.4

- 错误类型：指令违背

- 判断理由：出现禁用词：['新手']

### 样本 12

- 用户要求：用表格列出SFT和RLHF的定义和作用

- 模型回答：SFT是监督微调，RLHF是人类反馈强化学习。

- 解析规则：{'item_count': None, 'max_len_per_item': None, 'max_len_total': None, 'required_keywords': [], 'forbidden_words': [], 'no_explanation': False, 'need_table': True, 'need_simple': False, 'no_exaggeration': False}

- 综合得分：4.4

- 错误类型：格式/长度错误

- 判断理由：用户要求表格，但回答未体现表格结构

### 样本 18

- 用户要求：写3条小夜灯卖点，每条12字以内，不要出现刺眼

- 模型回答：柔光不刺眼；起夜方便；磁吸安装

- 解析规则：{'item_count': 3, 'max_len_per_item': 12, 'max_len_total': None, 'required_keywords': [], 'forbidden_words': ['刺眼'], 'no_explanation': False, 'need_table': False, 'need_simple': False, 'no_exaggeration': False}

- 综合得分：4.4

- 错误类型：指令违背

- 判断理由：出现禁用词：['刺眼']

### 样本 19

- 用户要求：判断下面回答是否合格：用户要求写3条，模型只写2条

- 模型回答：不合格，因为用户要求3条，但模型只写了2条。

- 解析规则：{'item_count': 3, 'max_len_per_item': None, 'max_len_total': None, 'required_keywords': [], 'forbidden_words': [], 'no_explanation': False, 'need_table': False, 'need_simple': False, 'no_exaggeration': False}

- 综合得分：3.8

- 错误类型：回答不完整

- 判断理由：条数不符合要求：要求3条，实际1条

### 样本 20

- 用户要求：判断下面回答是否合格：用户要求写3条，模型只写2条

- 模型回答：合格，模型回答比较简洁。

- 解析规则：{'item_count': 3, 'max_len_per_item': None, 'max_len_total': None, 'required_keywords': [], 'forbidden_words': [], 'no_explanation': False, 'need_table': False, 'need_simple': False, 'no_exaggeration': False}

- 综合得分：3.2

- 错误类型：判断错误

- 判断理由：条数不符合要求：要求3条，实际1条；判断错误：要求3条，实际2条，应判为不合格

### 样本 28

- 用户要求：写一个商品标题，必须包含感应灯和磁吸

- 模型回答：免布线小夜灯，起夜自动亮

- 解析规则：{'item_count': None, 'max_len_per_item': None, 'max_len_total': None, 'required_keywords': ['感应灯', '磁吸'], 'forbidden_words': [], 'no_explanation': False, 'need_table': False, 'need_simple': False, 'no_exaggeration': False}

- 综合得分：4.4

- 错误类型：指令违背

- 判断理由：缺少必须包含关键词：['感应灯', '磁吸']

### 样本 30

- 用户要求：总结用户投诉原因，要求列出3点

- 模型回答：用户主要是不满意购物体验，建议商家优化服务。

- 解析规则：{'item_count': 3, 'max_len_per_item': None, 'max_len_total': None, 'required_keywords': [], 'forbidden_words': [], 'no_explanation': False, 'need_table': False, 'need_simple': False, 'no_exaggeration': False}

- 综合得分：3.8

- 错误类型：回答不完整

- 判断理由：条数不符合要求：要求3条，实际1条

## 六、优化建议

1. 对指令违背类问题，应加强模型对必须包含、禁用词、禁止解释等限制条件的遵循能力。

2. 对格式/长度错误，应增加条数、字数、表格格式等约束类训练样本。

3. 对回答不完整问题，应优化模型对多条件任务的拆解能力。

4. 对自然度、准确性等语义类指标，建议结合人工复核或大模型辅助评估。
