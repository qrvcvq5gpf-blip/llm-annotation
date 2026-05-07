#!/usr/bin/env python3
"""
LLM 标注工具 - 批量标注 LLM 问答质量。

Usage:
    python annotate.py <input.xlsx> [output.xlsx] [--api-key KEY] [--model MODEL] [--base-url URL]
"""

import argparse
import json
import os
import sqlite3
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from anthropic import Anthropic

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")


def safe_print(text: str):
    """跨平台安全打印，避免 Windows 控制台中文乱码。"""
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode("utf-8", errors="replace").decode("ascii", errors="replace"))


def get_cc_switch_config() -> Optional[dict]:
    """从 cc-switch 数据库自动获取 API key 和 base_url。

    返回 dict，包含：
        api_key: API 认证 token
        base_url: 接口地址（直连 Anthropic 时为 None）
        default_model: 该服务商的默认模型（可能为 None）
    """
    cc_switch_db = Path.home() / ".cc-switch" / "cc-switch.db"
    if not cc_switch_db.exists():
        return None

    try:
        conn = sqlite3.connect(str(cc_switch_db))
        cursor = conn.cursor()
        cursor.execute(
            "SELECT settings_config FROM providers WHERE is_current = 1 LIMIT 1"
        )
        row = cursor.fetchone()
        conn.close()
        if not row:
            return None
        settings = json.loads(row[0])
        env = settings.get("env", {})

        api_key = env.get("ANTHROPIC_AUTH_TOKEN")
        base_url = env.get("ANTHROPIC_BASE_URL")
        if base_url and not base_url.endswith("/"):
            base_url = base_url + "/"

        default_model = env.get("ANTHROPIC_MODEL")
        return {
            "api_key": api_key,
            "base_url": base_url,
            "default_model": default_model,
        }
    except Exception:
        return None


def build_cite_snippets(row) -> tuple:
    """从行数据中构建溯源片段文本。

    Returns:
        tuple: (snippets_text, filled_top_numbers)，filled_top_numbers 为有内容的 TOP 编号列表
    """
    snippets = []
    filled_tops = []
    for i in range(1, 6):
        col = f"TOP{i}溯源片段"
        if col in row and pd.notna(row[col]) and str(row[col]).strip():
            snippets.append(f"- TOP{i}溯源片段：{row[col]}")
            filled_tops.append(i)
    if not snippets:
        return "-（无溯源片段）", []
    return "\n".join(snippets), filled_tops


def build_prompt(row, filled_tops: list) -> str:
    """为单行数据构建标注 prompt。"""
    snippets = []
    for n in filled_tops:
        col = f"TOP{n}溯源片段"
        try:
            snippets.append(f"- {col}：{row[col]}")
        except KeyError:
            pass
    cite_snippets_text = "\n".join(snippets) if snippets else "-（无溯源片段）"

    top_eval_lines = []
    for n in filled_tops:
        top_eval_lines.append(f'    "TOP{n}溯源是否正确": "正确"/"错误"')

    prompt = f"""你是一个大模型预标注工具，用于评估机器答案的质量和溯源片段的准确性。

## 标注标准

### 一、机器答案是否正确

**正确**：机器答案所要表达的意思和标准答案一致，关键内容不少说且不多说错误信息。

**错误**：符合以下任一情况：
1. 机器答案所要表达的意思和标准答案不一致；
2. 标准答案需要多个方面回答，而机器答案回复不全；
3. 机器答案较标准答案多说完全不相关或错误的内容。

### 二、溯源结果是否正确（整体评估，只在有溯源片段时评估）

**完全正确**：符合以下任一情况：
1. 在TOP1-5的溯源片段结果中出现完整相关内容描述；
2. 标准答案内容需要多个方面的知识总结，而每个溯源结果中只有其中一部分，但整体均可溯源到。

**部分正确**：所有TOP1-5溯源片段包含部分标准答案内容，无法整体均可溯源到。

**错误**：所有TOP1-5溯源片段都没有标准答案完整相关的内容，缺失不全或完全没有。

### 三、TOPn溯源是否正确（逐条评估，只能填"正确"或"错误"，不能填"部分正确"）

**正确**：符合以下任一情况：
1. 问题的标准答案在TOPn溯源中可完整找到；
2. TOPn溯源中只能找到部分标准答案。

**错误**：符合以下任一情况：
1. TOPn溯源中完全找不到与标准答案相关的信息；
2. TOPn溯源中可以找到与正确标准答案相关信息，但是溯源截断。

## 输入数据

- 用户问题：{row.get("用户问题", "")}
- 标准答案：{row.get("标准答案", "")}
- 机器答案：{row.get("机器答案", "")}
{cite_snippets_text}

## 输出要求

请严格按照以下JSON格式输出，不要包含任何其他内容：

{{
    "机器答案是否正确": "正确"/"错误",
    "溯源结果是否正确": "完全正确"/"部分正确"/"错误",
{chr(10).join(top_eval_lines)}
}}
"""
    return prompt


def parse_annotation_response(response_text: str) -> Optional[dict]:
    """解析 Claude 返回的 JSON 响应。"""
    try:
        text = response_text.strip()
        if text.startswith("```"):
            lines = text.split("\n")
            text = "\n".join(lines[1:-1])
        text = text.strip()
        if text.startswith("{"):
            end = text.rfind("}")
            if end != -1:
                text = text[:end+1]
        return json.loads(text)
    except json.JSONDecodeError:
        return None


def annotate_row(client: Anthropic, row: pd.Series, model: str, filled_tops: list) -> dict:
    """使用 Claude API 标注单行数据。"""
    prompt = build_prompt(row, filled_tops)

    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.messages.create(
                model=model,
                max_tokens=1024,
                messages=[{"role": "user", "content": prompt}]
            )
            # 跳过 thinking blocks，只取 text block
            text = None
            for block in response.content:
                if block.type == "text":
                    text = block.text
                    break
            if text is None:
                continue
            result = parse_annotation_response(text)
            if result:
                return result
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
            else:
                raise

    # 标注失败：标记为"标注失败"而非默认"错误"，避免引入系统性偏差污染数据
    result = {
        "机器答案是否正确": "标注失败",
        "溯源结果是否正确": "标注失败",
    }
    for n in filled_tops:
        result[f"TOP{n}溯源是否正确"] = "标注失败"
    return result


def print_summary(output_df: pd.DataFrame):
    """打印标注统计摘要。"""
    total = len(output_df)
    answer_col = "机器答案是否正确"
    cite_col = "溯源结果是否正确"

    # 机器答案统计
    answer_series = output_df[answer_col].fillna("")
    correct_count = (answer_series == "正确").sum()
    wrong_count = (answer_series == "错误").sum()
    failed_count = (answer_series == "标注失败").sum()
    answer_valid_total = correct_count + wrong_count  # 排除标注失败

    # 溯源统计（只统计有溯源数据且非标注失败的行）
    cite_series = output_df[cite_col].fillna("")
    cite_valid = cite_series[(cite_series != "") & (cite_series != "标注失败")]
    cite_total = len(cite_valid)
    cite_full = (cite_valid == "完全正确").sum()
    cite_partial = (cite_valid == "部分正确").sum()
    cite_wrong = (cite_valid == "错误").sum()

    safe_print("\n=== 标注完成 ===")
    safe_print(f"总行数: {total}")

    if answer_valid_total > 0:
        safe_print(f"机器答案正确率: {correct_count}/{answer_valid_total} ({correct_count/answer_valid_total*100:.1f}%)")
    if failed_count > 0:
        safe_print(f"标注失败行数: {failed_count}（请检查这些行后重新运行，程序支持断点续标）")

    if cite_total > 0:
        safe_print(
            f"溯源完全正确: {cite_full}/{cite_total} ({cite_full/cite_total*100:.1f}%)  "
            f"部分正确: {cite_partial}/{cite_total} ({cite_partial/cite_total*100:.1f}%)  "
            f"错误: {cite_wrong}/{cite_total} ({cite_wrong/cite_total*100:.1f}%)"
        )
    else:
        safe_print("溯源片段: 无（所有行均无溯源数据）")


def main():
    # 修复 Windows 控制台中文打印乱码
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except AttributeError:
            pass

    parser = argparse.ArgumentParser(description="LLM 标注工具")
    parser.add_argument("input_file", help="输入 Excel 文件路径")
    parser.add_argument("output_file", nargs="?", default=None, help="输出 Excel 文件路径（默认：<输入文件名>_输出_YYYY年MM月DD日.xlsx）")
    parser.add_argument("--api-key", default=None, help="API key（未提供时自动从 cc-switch 读取）")
    parser.add_argument("--base-url", default=None, help="API base URL（未提供时自动从 cc-switch 读取）")
    parser.add_argument("--model", default=None, help="使用的模型（未提供时自动从 cc-switch 读取）")
    args = parser.parse_args()

    input_path = Path(args.input_file)
    if args.output_file:
        output_file = args.output_file
    else:
        date_str = datetime.now().strftime("%Y年%m月%d日")
        output_file = str(input_path.parent / f"{input_path.stem}_输出_{date_str}.xlsx")

    # 配置解析：命令行参数 > 环境变量 > cc-switch 自动检测
    cc_config = get_cc_switch_config()

    api_key = args.api_key or ANTHROPIC_API_KEY or (cc_config.get("api_key") if cc_config else None)
    base_url = args.base_url or os.environ.get("ANTHROPIC_BASE_URL")
    if base_url is None and cc_config:
        base_url = cc_config.get("base_url")
    model = args.model or os.environ.get("ANTHROPIC_MODEL")
    if not model and cc_config and cc_config.get("default_model"):
        model = cc_config["default_model"]
    if not model:
        model = "claude-opus-4-6"

    if not api_key:
        safe_print("错误：未找到 API key。请设置 ANTHROPIC_API_KEY/ANTHROPIC_AUTH_TOKEN 环境变量、使用 --api-key 参数，或安装并配置 cc-switch。")
        sys.exit(1)

    if not os.path.exists(args.input_file):
        safe_print(f"错误：输入文件不存在：{args.input_file}")
        sys.exit(1)

    df = pd.read_excel(args.input_file)

    # 归一化列名：去除"（必填）"/"（选填）"后缀
    col_mapping = {}
    for col in df.columns:
        normalized = col.replace("（必填）", "").replace("（选填）", "")
        if normalized != col:
            col_mapping[col] = normalized
    if col_mapping:
        df.rename(columns=col_mapping, inplace=True)
        safe_print(f"列名已归一化：{col_mapping}")

    required_cols = ["callid", "用户问题", "标准答案", "机器答案"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        safe_print(f"错误：缺少必填列：{missing}")
        sys.exit(1)

    client_kwargs = {"api_key": api_key}
    if base_url:
        client_kwargs["base_url"] = base_url
    client = Anthropic(**client_kwargs)

    safe_print(f"使用服务商：cc-switch")
    safe_print(f"Base URL：{base_url or '默认 (auth.anthropic.com)'}")
    safe_print(f"模型：{model}")

    # 固定输出列顺序
    output_cols = [
        "callid（必填）", "用户问题（必填）", "标准答案（必填）", "机器答案（必填）",
        "机器答案是否正确", "溯源结果是否正确",
        "TOP1溯源片段（选填）", "TOP1溯源是否正确",
        "TOP2溯源片段（选填）", "TOP2溯源是否正确",
        "TOP3溯源片段（选填）", "TOP3溯源是否正确",
        "TOP4溯源片段（选填）", "TOP4溯源是否正确",
        "TOP5溯源片段（选填）", "TOP5溯源是否正确",
    ]
    output_df = pd.DataFrame(columns=output_cols)

    col_map = {
        "callid": "callid（必填）",
        "用户问题": "用户问题（必填）",
        "标准答案": "标准答案（必填）",
        "机器答案": "机器答案（必填）",
    }
    for i in range(1, 6):
        src = f"TOP{i}溯源片段"
        if src in df.columns:
            col_map[src] = f"TOP{i}溯源片段（选填）"

    for src_col, dst_col in col_map.items():
        output_df[dst_col] = df[src_col]

    # ── 断点续标：检查已有输出文件，恢复已完成的标注 ──
    start_idx = 0
    if os.path.exists(output_file):
        try:
            existing_df = pd.read_excel(output_file)
            if "机器答案是否正确" in existing_df.columns:
                annotated_mask = (
                    existing_df["机器答案是否正确"].notna()
                    & (existing_df["机器答案是否正确"].astype(str).str.strip() != "")
                )
                start_idx = int(annotated_mask.sum())
                if start_idx > 0:
                    safe_print(f"发现已有输出文件，跳过前 {start_idx} 行，从第 {start_idx + 1} 行继续...")
                    annotation_cols = (
                        ["机器答案是否正确", "溯源结果是否正确"]
                        + [f"TOP{n}溯源是否正确" for n in range(1, 6)]
                    )
                    for ann_col in annotation_cols:
                        if ann_col in existing_df.columns:
                            for i in range(min(start_idx, len(existing_df))):
                                output_df.at[i, ann_col] = existing_df.at[i, ann_col]
        except Exception as e:
            safe_print(f"警告：读取已有输出文件失败，将重新标注全部数据。({e})")
            start_idx = 0

    total = len(df)
    safe_print(f"开始标注，共 {total} 行，从第 {start_idx + 1} 行开始...")

    for idx, (_, row) in enumerate(df.iterrows()):
        if idx < start_idx:
            continue

        if (idx + 1) % 10 == 0:
            safe_print(f"正在处理第 {idx + 1}/{total} 行...")

        row_filled_tops = []
        for n in range(1, 6):
            col = f"TOP{n}溯源片段"
            if col in row and pd.notna(row[col]) and str(row[col]).strip():
                row_filled_tops.append(n)

        result = annotate_row(client, row, model, row_filled_tops)

        output_df.at[idx, "机器答案是否正确"] = result.get("机器答案是否正确", "标注失败")
        if row_filled_tops:
            output_df.at[idx, "溯源结果是否正确"] = result.get("溯源结果是否正确", "标注失败")
        for n in row_filled_tops:
            output_df.at[idx, f"TOP{n}溯源是否正确"] = result.get(f"TOP{n}溯源是否正确", "标注失败")

        # 每行完成后实时保存，确保中途失败可从断点继续
        try:
            output_df.to_excel(output_file, index=False)
        except Exception as e:
            safe_print(f"警告：保存进度时出错（{e}），将在下一行尝试重新保存")

    print_summary(output_df)
    safe_print(f"\n输出文件：{output_file}")


if __name__ == "__main__":
    main()
