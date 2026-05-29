# -*- coding: utf-8 -*-
"""
爬虫线程
"""

import time
import threading

from selenium.webdriver.common.by import By
from selenium.common.exceptions import (
    WebDriverException,
)
from PySide6.QtCore import Signal, QObject

from ..constants import (
    COL_PLAN_ID,
    CARD_SELECTORS,
    TYPE_DROPDOWN_CONTAINER,
    TYPE_DROPDOWN_BOX,
    TYPE_DROPDOWN_ITEM,
)
from ..parser.text_parser import parse_text


# ============================================================
# 信号
# ============================================================
class WorkerSignals(QObject):
    log = Signal(str)
    status = Signal(str)
    new_data = Signal(list)
    type_switched = Signal(str)
    error = Signal(str)
    finished = Signal()


# ============================================================
# 爬虫线程
# ============================================================
class CrawlerThread(threading.Thread):

    def __init__(self, driver, signals, stop_event,
                 traverse_types=True):
        super().__init__(daemon=True)
        self.driver = driver
        self.signals = signals
        self.stop_event = stop_event
        self.traverse_types = traverse_types
        self.seen_ids = set()

    def run(self):
        try:
            if self.traverse_types:
                self._traverse_types_loop()
            else:
                self._scroll_loop()
        except Exception as e:
            self.signals.error.emit(f"抓取异常: {str(e)}")
            self.signals.log.emit(f"[错误] {str(e)}")
        finally:
            self.signals.finished.emit()

    # ---------- 自动遍历资费类型 ----------

    def _traverse_types_loop(self):
        self.signals.status.emit("遍历资费类型中...")
        self.signals.log.emit(
            "[信息] ========== 开始自动遍历所有资费类型 ==========")

        type_options = self._get_type_options()
        if not type_options:
            self.signals.log.emit(
                "[警告] 未能获取资费类型选项列表，回退为仅抓取当前页面")
            self._scroll_loop()
            return

        self.signals.log.emit(
            f"[信息] 检测到 {len(type_options)} 个资费类型: "
            + "、".join(t["label"] for t in type_options))

        for idx, type_opt in enumerate(type_options):
            if self.stop_event.is_set():
                break

            type_label = type_opt["label"]
            type_value = type_opt["value"]
            self.signals.log.emit(
                f"\n[信息] ----- 切换资费类型 "
                f"[{idx+1}/{len(type_options)}]: {type_label} -----")
            self.signals.type_switched.emit(
                f"遍历中 [{idx+1}/{len(type_options)}] {type_label}")
            self.signals.status.emit(
                f"正在抓取: {type_label} ({idx+1}/{len(type_options)})")

            if not self._switch_type(type_value, type_label):
                self.signals.log.emit(
                    f"[警告] 切换资费类型「{type_label}」失败，跳过")
                continue

            self.signals.log.emit("[信息] 等待页面加载...")
            time.sleep(2.0)
            self._scroll_for_current_type()

        self.signals.log.emit(
            "\n[信息] ========== 所有资费类型遍历完毕 ==========")
        self.signals.status.emit("遍历完成")

    def _get_type_options(self):
        try:
            containers = self.driver.find_elements(
                By.CSS_SELECTOR, TYPE_DROPDOWN_CONTAINER)
            if not containers:
                self.signals.log.emit("[警告] 未找到资费类型下拉框容器")
                return []

            type_container = containers[0]
            box = type_container.find_element(
                By.CSS_SELECTOR, TYPE_DROPDOWN_BOX)
            box.click()
            time.sleep(0.8)

            items = type_container.find_elements(
                By.CSS_SELECTOR, TYPE_DROPDOWN_ITEM)
            options = []
            for item in items:
                label_el = item.find_element(
                    By.CSS_SELECTOR, ".item-label")
                value = item.get_attribute("data-value") or ""
                label = label_el.text.strip()
                if label:
                    options.append({"value": value, "label": label})

            self.driver.execute_script(
                "document.querySelector('.select-box')?.click();"
                "document.body.click();")
            time.sleep(0.3)
            return options

        except Exception as e:
            self.signals.log.emit(f"[警告] 获取资费类型选项失败: {e}")
            return []

    def _switch_type(self, value, label):
        try:
            containers = self.driver.find_elements(
                By.CSS_SELECTOR, TYPE_DROPDOWN_CONTAINER)
            if not containers:
                return False

            type_container = containers[0]
            box = type_container.find_element(
                By.CSS_SELECTOR, TYPE_DROPDOWN_BOX)
            box.click()
            time.sleep(0.8)

            items = type_container.find_elements(
                By.CSS_SELECTOR, TYPE_DROPDOWN_ITEM)
            for item in items:
                item_value = item.get_attribute("data-value") or ""
                item_label_el = item.find_element(
                    By.CSS_SELECTOR, ".item-label")
                item_label = item_label_el.text.strip()
                if item_value == value or item_label == label:
                    item.click()
                    self.signals.log.emit(
                        f"[信息] 已选择资费类型: {label}")
                    time.sleep(0.5)
                    return True

            self.signals.log.emit(
                f"[警告] 未找到选项 value={value} label={label}")
            return False

        except Exception as e:
            self.signals.log.emit(f"[警告] 切换资费类型出错: {e}")
            return False

    def _scroll_for_current_type(self):
        try:
            self.driver.execute_script("window.scrollTo(0, 0)")
            time.sleep(0.8)
        except Exception:
            pass

        last_card_count = self._get_card_count_js()
        self._parse_all_and_emit()

        no_new_count = 0
        max_no_new = 3
        scroll_round = 0

        while not self.stop_event.is_set():
            scroll_round += 1
            try:
                self.driver.execute_script(
                    "window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(1.0)

                current_count = self._get_card_count_js()

                if current_count == last_card_count:
                    no_new_count += 1
                    self.signals.log.emit(
                        f"[信息] 卡片数未变化 ({current_count}条, "
                        f"{no_new_count}/{max_no_new})")
                    if no_new_count >= max_no_new:
                        self.signals.log.emit(
                            f"[信息] 资费类型「"
                            f"{self._current_type_label()}」"
                            f"加载完毕（{scroll_round}轮，"
                            f"{current_count}条）")
                        break
                else:
                    no_new_count = 0
                    self.signals.log.emit(
                        f"[信息] 第{scroll_round}轮: "
                        f"{last_card_count} → {current_count} 条")
                    self._parse_all_and_emit()
                    last_card_count = current_count

            except WebDriverException as e:
                self.signals.log.emit(f"[警告] 滚动异常: {e}")
                try:
                    _ = self.driver.current_url
                except Exception:
                    self.signals.error.emit("浏览器已断开连接")
                    return
                time.sleep(3)

    def _get_card_count_js(self):
        try:
            selector_list = ", ".join(
                f"'{s}'" for s in CARD_SELECTORS)
            count = self.driver.execute_script(f"""
                var selectors = [{selector_list}];
                for (var i = 0; i < selectors.length; i++) {{
                    var els = document.querySelectorAll(selectors[i]);
                    if (els.length > 0) return els.length;
                }}
                return 0;
            """)
            return count or 0
        except Exception:
            return 0

    def _parse_all_and_emit(self):
        try:
            items = self.driver.execute_script("""
                var cards = document.querySelectorAll(
                    '.tariff-item-container');
                var result = [];
                for (var i = 0; i < cards.length; i++) {
                    var nameEl = cards[i].querySelector('.item-name');
                    result.push({
                        name: nameEl
                            ? nameEl.textContent.trim() : '',
                        text: cards[i].innerText
                    });
                }
                return result;
            """)
            if not items:
                return

            new_batch = []
            for item in items:
                try:
                    data = parse_text(item["text"], item["name"])
                    if not data:
                        continue
                    plan_id = data[COL_PLAN_ID]
                    if not plan_id:
                        continue
                    if plan_id in self.seen_ids:
                        continue
                    self.seen_ids.add(plan_id)
                    new_batch.append(data)
                except Exception as e:
                    self.signals.log.emit(
                        f"[调试] 解析条目失败: {e}")
                    continue

            if new_batch:
                self.signals.new_data.emit(new_batch)

        except Exception as e:
            self.signals.log.emit(f"[警告] 批量提取失败: {e}")

    def _current_type_label(self):
        try:
            containers = self.driver.find_elements(
                By.CSS_SELECTOR, TYPE_DROPDOWN_CONTAINER)
            if containers:
                tips = containers[0].find_element(
                    By.CSS_SELECTOR, ".tipsText")
                return tips.text.strip()
        except Exception:
            pass
        return ""

    # ---------- 单类型滚动 ----------

    def _scroll_loop(self):
        self.signals.status.emit("抓取中...")
        self.signals.log.emit("[信息] 开始自动滚动页面...")

        last_count = self._get_card_count_js()
        self._parse_all_and_emit()
        stable_count = 0
        max_stable = 5
        scroll_round = 0

        while not self.stop_event.is_set():
            scroll_round += 1
            try:
                self.driver.execute_script(
                    "window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(1.0)

                current_count = self._get_card_count_js()

                if current_count == last_count:
                    stable_count += 1
                    if stable_count >= max_stable:
                        self.signals.log.emit(
                            "[信息] 已滚动到底部，数据加载完毕")
                        break
                else:
                    stable_count = 0
                    self._parse_all_and_emit()
                    last_count = current_count

                self.signals.log.emit(
                    f"[信息] 第 {scroll_round} 轮，"
                    f"累计 {current_count} 条")

            except WebDriverException as e:
                self.signals.log.emit(f"[警告] 浏览器异常: {e}")
                try:
                    _ = self.driver.current_url
                except Exception:
                    self.signals.error.emit("浏览器已断开连接")
                    return
                time.sleep(3)

        self.signals.status.emit("抓取完成")