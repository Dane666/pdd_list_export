# -*- coding: utf-8 -*-
"""
拼多多订单抓取脚本（复用本机 Chromium 登录态）
- 直接读取页面内嵌 window.rawData.ordersStore.initOdersList（全量订单，无需翻页）
- 抓取 2026-06-01 至今的订单，输出 Excel
"""
import os
import json
import shutil
from datetime import datetime

from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

# 配置
START_DATE = datetime(2026, 6, 1)          # 起始日期（含），改成你需要的最早日期
ORDER_URL = "https://mobile.yangkeduo.com/orders.html"
PROFILE_DIR = "./pdd_profile"
OUTPUT_XLSX = "拼多多订单_2026-06至今.xlsx"


def detect_chrome_user_data():
    """自动探测本机 Chrome 用户数据目录（macOS / Linux / Windows）"""
    import platform
    home = os.path.expanduser("~")
    sysname = platform.system()
    candidates = []
    if sysname == "Darwin":
        candidates = [
            os.path.join(home, "Library/Application Support/Google/Chrome"),
            os.path.join(home, "Library/Application Support/Chromium"),
            os.path.join(home, "Library/Application Support/Microsoft Edge"),
        ]
    elif sysname == "Windows":
        candidates = [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google/Chrome/User Data"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft/Edge/User Data"),
        ]
    else:  # Linux
        candidates = [
            os.path.join(home, ".config/google-chrome"),
            os.path.join(home, ".config/chromium"),
            os.path.join(home, ".config/microsoft-edge"),
        ]
    for c in candidates:
        if c and os.path.isdir(c):
            return c
    return ""


CHROME_USER_DATA = detect_chrome_user_data()

HEADERS = ["下单时间", "订单号", "商品名称", "规格", "单价(元)", "数量",
           "订单金额(元)", "实付(元)", "订单状态", "店铺名", "快递单号"]


def ts_to_str(ts):
    if not ts:
        return ""
    try:
        return datetime.fromtimestamp(int(ts)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""


def yuan(cent):
    """分 → 元（保留 2 位）"""
    if cent in (None, ""):
        return ""
    try:
        return round(float(cent) / 100.0, 2)
    except Exception:
        return ""


def extract_orders(page):
    """读取 initOdersList 全量订单（蛇形命名字段）"""
    js = """
    () => {
        const rd = window.rawData;
        if (!rd) return [];
        const store = rd.ordersStore || {};
        const init = store.initOdersList || {};
        // initOdersList 形如 { '0': [订单...] }，合并所有 tab
        let all = [];
        for (const k in init) {
            if (Array.isArray(init[k])) all = all.concat(init[k]);
        }
        // 去重
        const uniq = new Map();
        all.forEach(o => {
            const sn = o.order_sn || o.orderSn;
            if (sn) uniq.set(sn, o);
        });
        return Array.from(uniq.values()).map(o => ({
            orderSn: o.order_sn || o.orderSn,
            orderTime: o.order_time || o.orderTime,
            orderAmount: o.order_amount ?? o.orderAmount,
            displayAmount: o.display_amount ?? o.displayAmount,
            statusPrompt: o.order_status_prompt || o.orderStatusPrompt,
            mallName: (o.mall && (o.mall.mall_name || o.mall.mallName)) || "",
            trackingNumber: o.tracking_number || o.trackingNumber || "",
            goods: (o.order_goods || o.orderGoods || []).map(g => ({
                goodsName: g.goods_name || g.goodsName,
                goodsPrice: g.goods_price ?? g.goodsPrice,
                goodsNumber: g.goods_number ?? g.goodsNumber,
                spec: g.spec,
            })),
        }));
    }
    """
    return page.evaluate(js)


def main():
    if not CHROME_USER_DATA:
        print("[!] 未找到本机 Chrome/Edge 用户数据目录，请在脚本顶部手动设置 CHROME_USER_DATA。")
        return

    if not os.path.exists(PROFILE_DIR):
        print("[*] 首次运行：正在复制本机 Chrome 登录态（约需几十秒）...")
        shutil.copytree(CHROME_USER_DATA, PROFILE_DIR,
                        ignore=shutil.ignore_patterns(
                            "Singleton*", "lockfile", "*.pma",
                            "Crashpad", "ShaderCache", "GPUCache",
                            "Code Cache", "Cache", "CachedData",
                            "Service Worker", "GCM Store", "Network",
                        ))

    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False, channel="chrome",
            viewport={"width": 1280, "height": 900},
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
            args=["--disable-blink-features=AutomationControlled"],
        )
        browser.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        )
        page = browser.pages[0] if browser.pages else browser.new_page()
        page.goto(ORDER_URL, timeout=60000)
        print("[*] 已打开订单页。若未登录，请在弹出的浏览器里扫码/登录。")
        input("登录完成后按回车，开始读取全量订单...")

        # 等待 rawData 加载完成
        page.wait_for_function(
            "() => window.rawData && window.rawData.ordersStore "
            "&& window.rawData.ordersStore.initOdersList",
            timeout=30000,
        )
        page.wait_for_timeout(3000)

        # 滚动窗口触发「上拉加载」，让 initOdersList 累积增长
        seen = {}
        reached_start = False
        idle = 0
        max_rounds = 2000
        for rnd in range(max_rounds):
            orders = extract_orders(page)
            before = len(seen)
            earliest_ts = None
            for o in orders:
                sn = o.get("orderSn")
                if sn and sn not in seen:
                    seen[sn] = o
                ts = o.get("orderTime")
                if ts:
                    earliest_ts = ts if earliest_ts is None else min(earliest_ts, ts)
            new_cnt = len(seen) - before

            if earliest_ts and datetime.fromtimestamp(int(earliest_ts)) < START_DATE:
                reached_start = True

            print(f"[*] 第{rnd}轮：累计 {len(seen)} 条（本轮 +{new_cnt}）"
                  f"{'  [已触及 2026-06-01]' if reached_start else ''}")

            if reached_start:
                break

            # 滚动窗口到底，触发「上拉加载」
            # 先用 wheel 模拟真实滚动（触发拼多多的滚动监听），再直接定位到底
            page.mouse.move(640, 800)
            page.mouse.wheel(0, 8000)
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(1500)

            # 检测是否真的滚到了页面最底部
            at_bottom = page.evaluate(
                "() => window.scrollY + window.innerHeight >= "
                "document.body.scrollHeight - 20"
            )

            if new_cnt == 0 and at_bottom:
                idle += 1
                if idle >= 15:
                    print("[*] 已滚到底且连续多轮无新订单，停止。")
                    break
            elif new_cnt > 0:
                idle = 0

        orders = list(seen.values())
        print(f"[*] 最终累计订单 {len(orders)} 条")
        browser.close()

    # 过滤 2026-06-01 至今
    rows = []
    for o in orders:
        ts = o.get("orderTime")
        dt = datetime.fromtimestamp(int(ts)) if ts else None
        if not dt or dt < START_DATE:
            continue
        goods = o.get("goods", []) or [{}]
        g0 = goods[0] if goods else {}
        rows.append([
            ts_to_str(ts),
            o.get("orderSn", ""),
            g0.get("goodsName", ""),
            g0.get("spec", ""),
            yuan(g0.get("goodsPrice")),
            g0.get("goodsNumber", ""),
            yuan(o.get("orderAmount")),
            yuan(o.get("displayAmount")),
            o.get("statusPrompt", ""),
            o.get("mallName", ""),
            o.get("trackingNumber", ""),
        ])

    # 按下单时间倒序
    rows.sort(key=lambda r: r[0], reverse=True)

    # 写 Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "拼多多订单"
    ws.append(HEADERS)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="E02E24")
    for c in ws[1]:
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    for row in rows:
        ws.append(row)
    widths = [19, 24, 45, 22, 10, 7, 12, 10, 16, 18, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A2"
    wb.save(OUTPUT_XLSX)

    total_pay = sum(float(r[7]) for r in rows if isinstance(r[7], (int, float)))
    print(f"\n✅ 完成！共 {len(rows)} 条订单（2026-06 至今），实付合计约 ¥{total_pay:.2f}")
    print(f"   已保存到：{os.path.abspath(OUTPUT_XLSX)}")


if __name__ == "__main__":
    main()
