# -*- coding: utf-8 -*-
"""
常量定义
"""

TARGET_URL = "https://www.10086.cn/fee/"

COLUMNS = [
    "方案编号", "套餐名称", "资费标准", "资费类型", "适用地区",
    "适用范围", "销售渠道", "上线日期", "下线日期", "有效期限",
    "在网要求", "国内通话", "国内通用流量", "定向流量", "短信",
    "退订方式", "违约责任", "其他服务内容"
]
COL_COUNT = len(COLUMNS)

# 列索引
COL_PLAN_ID = 0
COL_NAME = 1
COL_PRICE = 2

# CSS 选择器候选列表（按优先级排序）
CARD_SELECTORS = [
    ".tariff-item-container",     # 中国移动资费公示页面精确选择器
    ".fee-card-item",
    ".tariff-card",
    ".plan-card",
    ".card-item",
]

# 键值对字段映射
FIELD_MAP = {
    "方案编号": COL_PLAN_ID,
    "资费标准": COL_PRICE,
    "资费类型": 3,
    "适用地区": 4,
    "适用范围": 5,
    "销售渠道": 6,
    "上线日期": 7,
    "下线日期": 8,
    "有效期限": 9,
    "在网要求": 10,
    "退订方式": 15,
    "违约责任": 16,
    "其他服务内容": 17,
}

# 排除关键字（用于识别套餐名称）
EXCLUDE_KEYWORDS = [
    "资费标准", "适用地区", "方案编号", "适用范围", "销售渠道",
    "上线日期", "下线日期", "有效期限", "在网要求", "退订方式",
    "违约责任", "其他服务", "国内通话", "国内通用流量", "定向流量",
    "短信", "资费类型",
]

# 资费类型下拉框的 DOM 选择器
TYPE_DROPDOWN_CONTAINER = ".line-3 .select-container.the-select"
TYPE_DROPDOWN_BOX = ".select-box"
TYPE_DROPDOWN_LIST = ".select-list-box"
TYPE_DROPDOWN_ITEM = ".select-item"