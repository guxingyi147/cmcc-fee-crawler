# -*- coding: utf-8 -*-
"""
文本解析工具函数
"""

import re

from ..constants import (
    COL_COUNT, COL_NAME, COL_PLAN_ID, COL_PRICE,
    FIELD_MAP, EXCLUDE_KEYWORDS,
)


def extract_field_value(text, label):
    """从文本中提取 'label：value' 格式的字段值"""
    pattern = rf'{re.escape(label)}[：:\s]\s*(.+?)(?:\n|$)'
    match = re.search(pattern, text)
    if match:
        value = match.group(1).strip()
        value = re.sub(r'^[●\-\*•·]\s*', '', value)
        if value:
            return value
    return ""


def extract_resource_value(text, label, units):
    """提取资源类字段（通话分钟、流量、短信条数）"""
    # 先尝试键值对
    value = extract_field_value(text, label)
    if value:
        for unit in units:
            if unit in value:
                return value
        # 纯数字补单位
        if re.match(r'^[\d.]+$', value):
            if "通话" in label:
                return f"{value}分钟"
            elif "流量" in label:
                num = float(value)
                return f"{num/1024:.1f}GB" if num >= 1024 else f"{value}MB"
            elif "短信" in label:
                return f"{value}条"

    # 正则提取 label + 数字 + 单位
    unit_pat = "|".join(re.escape(u) for u in units)
    m = re.search(
        rf'{re.escape(label)}[：:\s]*([\d.]+)\s*({unit_pat})',
        text, re.IGNORECASE)
    if m:
        return f"{m.group(1)}{m.group(2)}"

    return ""


def extract_plan_name(card, text):
    """提取套餐名称"""
    from selenium.webdriver.common.by import By

    # 方法1：标题元素
    for sel in ["h3", "h4", "h5",
                "[class*='title']", "[class*='name']",
                "[class*='plan-name']", "[class*='card-title']"]:
        try:
            for t in card.find_elements(By.CSS_SELECTOR, sel):
                t_text = t.text.strip()
                if (t_text and 2 < len(t_text) < 100
                        and not any(kw in t_text for kw in EXCLUDE_KEYWORDS)):
                    return t_text
        except Exception:
            continue

    # 方法2：第一行非字段文本
    for line in text.strip().split("\n"):
        line = line.strip()
        if (line and 2 < len(line) < 80
                and not line.startswith("●")
                and not any(kw in line for kw in EXCLUDE_KEYWORDS)):
            return line
    return ""


def parse_text(text, plan_name=""):
    """从纯文本解析套餐数据 → 返回长度为 COL_COUNT 的列表，失败返回 None"""
    if not text or len(text) < 10:
        return None

    data = [""] * COL_COUNT

    # 套餐名称：优先使用 JS 提取的精确名称
    if plan_name and len(plan_name) > 1:
        data[COL_NAME] = plan_name
    else:
        # 回退：取第一行非字段文本
        for line in text.strip().split("\n"):
            line = line.strip()
            if (line and 2 < len(line) < 80
                    and not line.startswith("●")
                    and not any(kw in line for kw in EXCLUDE_KEYWORDS)
                    and not re.match(r'^[\d.]+元', line)):
                data[COL_NAME] = line
                break

    # 键值对字段
    for label, col_idx in FIELD_MAP.items():
        val = extract_field_value(text, label)
        if val:
            data[col_idx] = val

    # 4 个特殊字段
    data[11] = extract_resource_value(text, "国内通话", ["分钟", "分"])
    data[12] = extract_resource_value(
        text, "国内通用流量", ["GB", "MB", "gb", "mb"])
    data[13] = extract_resource_value(
        text, "定向流量", ["GB", "MB", "gb", "mb"])
    data[14] = extract_resource_value(text, "短信", ["条"])

    return data


def parse_card(card):
    """解析单个套餐卡片 → 返回长度为 COL_COUNT 的列表，失败返回 None"""
    try:
        text = card.text
    except Exception:
        return None

    if not text or len(text) < 10:
        return None

    data = [""] * COL_COUNT

    # 套餐名称
    data[COL_NAME] = extract_plan_name(card, text)

    # 键值对字段
    for label, col_idx in FIELD_MAP.items():
        val = extract_field_value(text, label)
        if val:
            data[col_idx] = val

    # 4 个特殊字段
    data[11] = extract_resource_value(text, "国内通话", ["分钟", "分"])
    data[12] = extract_resource_value(
        text, "国内通用流量", ["GB", "MB", "gb", "mb"])
    data[13] = extract_resource_value(
        text, "定向流量", ["GB", "MB", "gb", "mb"])
    data[14] = extract_resource_value(text, "短信", ["条"])

    return data


def find_cards(driver):
    """在页面中查找所有套餐卡片元素"""
    from selenium.webdriver.common.by import By
    from ..constants import CARD_SELECTORS

    # CSS 选择器
    for sel in CARD_SELECTORS:
        try:
            cards = driver.find_elements(By.CSS_SELECTOR, sel)
            if cards:
                return cards
        except Exception:
            continue

    # 回退：通过文本定位父容器
    try:
        elements = driver.find_elements(
            By.XPATH,
            "//*[contains(text(),'方案编号') or contains(text(),'资费标准')]")
        cards, seen = [], set()
        for el in elements:
            try:
                parent = el.find_element(By.XPATH,
                    "./ancestor::*[contains(@class,'card') or "
                    "contains(@class,'item') or contains(@class,'plan') or "
                    "contains(@class,'tariff') or contains(@class,'fee') or "
                    "contains(@class,'list') or contains(@class,'content') or "
                    "contains(@class,'detail')][1]")
                pid = parent.id
                if pid not in seen:
                    seen.add(pid)
                    cards.append(parent)
            except Exception:
                continue
        if cards:
            return cards
    except Exception:
        pass

    return []