---
name: llm-annotation
description: "Use when the user provides an Excel file with LLM Q&A data (callid, user question, reference answer, machine answer, optional TOP1-5 citation snippets) and wants to automatically annotate the quality of machine answers and citation accuracy using Claude API. Trigger when user mentions '标注', 'annotation', '标注工具', or provides an input Excel file for batch LLM evaluation. NOT for manual single-item annotation -- only when batch processing is needed."
---

# LLM 大模型预标注工具

批量标注 LLM 问答质量的自动化工具。读取 Excel 文件，逐行调用 Claude API，评估机器答案正确性与溯源片段准确性，并将结果实时写回输出文件（支持断点续标）。

## 适用场景

**使用本工具的情况：**
- 用户提供了含 LLM 问答数据的 Excel 文件，需要批量评估
- 用户提到"标注"、"大模型标注"、"annotation tool"、"批量评估"
- 用户需要评估机器答案质量或 TOP1-5 溯源片段准确性

**不适用的情况：**
- 单条手动标注
- 实时对话评估
- 非 Excel 格式输入（不做格式转换）

## 输入格式

| 列名 | 是否必填 | 说明 |
|------|----------|------|
| callid（必填） | 必填 | 对话唯一标识 |
| 用户问题（必填） | 必填 | 用户提出的问题 |
| 标准答案（必填） | 必填 | 人工标注的参考答案 |
| 机器答案（必填） | 必填 | 大模型生成的回答 |
| TOP1溯源片段（选填） | 选填 | 第1条溯源召回片段 |
| TOP2溯源片段（选填） | 选填 | 第2条溯源召回片段 |
| TOP3溯源片段（选填） | 选填 | 第3条溯源召回片段 |
| TOP4溯源片段（选填） | 选填 | 第4条溯源召回片段 |
| TOP5溯源片段（选填） | 选填 | 第5条溯源召回片段 |

> 列名中的"（必填）"/"（选填）"标记会被自动忽略，无需手动去除。

## 输出格式

在输入列基础上新增标注结果列：

| 新增列 | 取值范围 | 说明 |
|--------|----------|------|
| 机器答案是否正确 | 正确 / 错误 / 标注失败 | 机器答案与标准答案的一致性评估 |
| 溯源结果是否正确 | 完全正确 / 部分正确 / 错误 / 标注失败 | 整体溯源质量（无溯源片段时为空） |
| TOPn溯源是否正确 | 正确 / 错误 / 标注失败 | 各条溯源片段逐一评估（有内容的才评估） |

> `标注失败`：API 调用多次重试仍无法解析结果，标记为失败而非默认"错误"，避免污染数据。

## 使用方法

### 基础命令

```bash
python annotate.py <输入文件.xlsx> [输出文件.xlsx] [--api-key KEY] [--model MODEL]
```

### 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `input.xlsx` | 必填 | 输入 Excel 文件路径 |
| `output.xlsx` | 选填 | 输出文件路径，不填则自动生成：`{输入文件名}_输出_{YYYY年MM月DD日}.xlsx` |
| `--api-key` | 选填 | Anthropic API key，优先级低于 cc-switch 自动检测 |
| `--model` | 选填 | 使用的 Claude 模型，不填则自动从 cc-switch 读取 |
| `--base-url` | 选填 | API Base URL，不填则自动从 cc-switch 读取 |

### 断点续标

程序每标注完一行就立即保存到输出文件。若中途中断，再次运行**相同命令**时会自动检测输出文件中已完成的行数，从断点继续，无需重新标注全部数据。

### 环境配置与自动检测

配置优先级（从高到低）：

**API Key：**
1. `--api-key` 命令行参数
2. `ANTHROPIC_API_KEY` / `ANTHROPIC_AUTH_TOKEN` 环境变量
3. **cc-switch 数据库**（自动读取当前激活的服务商配置）

**模型：**
1. `--model` 命令行参数
2. `ANTHROPIC_MODEL` 环境变量
3. **cc-switch 数据库**（自动读取）

```bash
# 方式一：设置环境变量
export ANTHROPIC_API_KEY="your-api-key"

# 方式二：安装 cc-switch 并配置服务商（推荐，无需额外操作）
```

## 标注逻辑

### 机器答案是否正确

**正确**：机器答案所要表达的意思和标准答案一致，关键内容不少说且不多说错误信息。

**错误**：符合以下任一情况：
1. 机器答案所要表达的意思和标准答案不一致；
2. 标准答案需要多个方面回答，而机器答案回复不全；
3. 机器答案较标准答案多说完全不相关或错误的内容。

### 溯源结果是否正确（整体评估，只在有溯源片段时评估）

**完全正确**：符合以下任一情况：
1. 在TOP1-5的溯源片段结果中出现完整相关内容描述；
2. 标准答案内容需要多个方面的知识总结，而每个溯源结果中只有其中一部分，但整体均可溯源到。

**部分正确**：所有TOP1-5溯源片段包含部分标准答案内容，无法整体均可溯源到。

**错误**：所有TOP1-5溯源片段都没有标准答案完整相关的内容，缺失不全或完全没有。

### TOPn溯源是否正确（逐条评估，只能填"正确"或"错误"，不能填"部分正确"）

**正确**：符合以下任一情况：
1. 问题的标准答案在TOPn溯源中可完整找到；
2. TOPn溯源中只能找到部分标准答案。

**错误**：符合以下任一情况：
1. TOPn溯源中完全找不到与标准答案相关的信息；
2. TOPn溯源中可以找到与正确标准答案相关信息，但是溯源截断。

### 动态标注说明

- TOP1-TOP5 为选填字段，根据每行实际填写数量动态评估
- 某行 TOP4/TOP5 为空时，对应标注列也为空（NaN）
- "溯源结果是否正确"只评估实际填写了内容的溯源片段

## 运行结束后输出示例

```
=== 标注完成 ===
总行数: 200
机器答案正确率: 142/200 (71.0%)
溯源完全正确: 89/180 (49.4%)  部分正确: 60/180 (33.3%)  错误: 31/180 (17.2%)

输出文件：D:\data\测试集_输出_2026年04月08日.xlsx
```

## 快速命令参考

| 场景 | 命令 |
|------|------|
| 自动生成输出文件名 | `python annotate.py input.xlsx` |
| 指定输出文件名 | `python annotate.py input.xlsx output.xlsx` |
| 指定 API key | `python annotate.py input.xlsx --api-key KEY` |
| 指定模型 | `python annotate.py input.xlsx --model claude-sonnet-4-6` |
| 断点续标（重跑同一命令即可） | `python annotate.py input.xlsx output.xlsx` |

## 常见问题

1. **未找到 API key**：设置 `ANTHROPIC_API_KEY` 环境变量，或使用 `--api-key` 参数，或安装并配置 cc-switch
2. **缺少必填列**：输入文件必须包含 callid、用户问题、标准答案、机器答案
3. **中途失败如何恢复**：直接重跑相同命令，程序会自动从断点继续
4. **标注失败行**：运行结束后统计摘要会显示失败行数，可重跑命令补标
5. **大文件耗时估算**：约 5-10 秒/行，100 行约需 10-15 分钟

## 环境依赖

```bash
pip install pandas openpyxl anthropic
```

- Python 3.8+
- pandas
- openpyxl
- anthropic SDK
