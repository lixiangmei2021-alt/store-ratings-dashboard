from __future__ import annotations

import html
import json
import logging
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
LOG_DIR = ROOT / "logs"
LATEST_JSON = DATA_DIR / "latest.json"
HISTORY_JSON = DATA_DIR / "history.json"
STATUS_JSON = DATA_DIR / "status.json"
DASHBOARD = ROOT / os.environ.get("DASHBOARD_NAME", "评分看板.html")
LATEST_XLSX = ROOT / "商店评分汇总_latest.xlsx"
VERSION_JSON = ROOT / "version.json"
PAGE_VERSION = "2026-11-24.4"
TZ = ZoneInfo("Asia/Shanghai")
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)

GAMES = [
    {
        "key": "frontier",
        "name": "Frontier City",
        "android": "com.camelgames.xcity",
        "ios_bundle": "com.camelgames.xcity",
        "ios_id": "6751943304",
        "ios_slug": "frontier-city",
        "ru_android": "com.camelgames.xcityru",
        "countries": ["US", "UK", "CA", "AU", "DE", "FR", "ES", "MX", "BR", "MY", "ID", "PH", "UA", "BY"],
    },
    {
        "key": "vice",
        "name": "Vice Throne",
        "android": "com.camelgames.mnnew",
        "ios_bundle": "com.camelgames.mnnew",
        "ios_id": "6798176762",
        "ios_slug": "vice-throne",
        "ru_android": "com.camelgames.mnnew.ru",
        "countries": ["US", "CA", "AU", "SG", "MY", "UK", "DE", "FR", "ES", "MX", "BR", "JP", "KR", "TW", "HK", "MO", "UA", "BY"],
    },
]

COUNTRY_NAMES = {
    "US": "美国", "UK": "英国", "CA": "加拿大", "AU": "澳大利亚", "SG": "新加坡",
    "MY": "马来西亚", "ID": "印度尼西亚", "PH": "菲律宾", "DE": "德国", "FR": "法国",
    "ES": "西班牙", "MX": "墨西哥", "BR": "巴西", "JP": "日本", "KR": "韩国",
    "TW": "中国台湾", "HK": "中国香港", "MO": "中国澳门", "UA": "乌克兰",
    "BY": "白俄罗斯", "RU": "俄罗斯",
}
GL = {"UK": "GB"}
HL = {
    "US": "en", "UK": "en_GB", "CA": "en", "AU": "en", "SG": "en", "MY": "en",
    "ID": "id", "PH": "en", "DE": "de", "FR": "fr", "ES": "es", "MX": "es",
    "BR": "pt-BR", "JP": "ja", "KR": "ko", "TW": "zh-TW", "HK": "zh-HK",
    "MO": "zh-TW", "UA": "uk", "BY": "ru", "RU": "ru",
}
APPLE_CC = {"UK": "gb"}


def now_local() -> datetime:
    return datetime.now(TZ)


def log_setup() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    day = now_local().strftime("%Y%m%d")
    logger = logging.getLogger("ratings")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    for path in (LOG_DIR / f"update_{day}.log", LOG_DIR / "latest.log"):
        handler = logging.FileHandler(path, encoding="utf-8")
        handler.setFormatter(fmt)
        logger.addHandler(handler)
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(fmt)
    logger.addHandler(console)
    return logger


def atomic_json(path: Path, obj: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def fetch(url: str, retries: int = 4, timeout: int = 35) -> tuple[bytes | None, dict]:
    last_error = None
    for attempt in range(retries):
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept": "text/html,application/json;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.9",
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return response.read(), {"ok": True, "http_status": response.status, "error": None}
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            if attempt + 1 < retries:
                time.sleep(1.2 * (attempt + 1))
    return None, {"ok": False, "http_status": None, "error": last_error}


def parse_count(text: str | None) -> int | None:
    if not text:
        return None
    match = re.search(r"[0-9][0-9,.]*", text)
    if not match:
        return None
    digits = re.sub(r"[^0-9]", "", match.group(0))
    return int(digits) if digits else None


def parse_gplay(body: bytes) -> dict:
    source = body.decode("utf-8", errors="ignore")
    title = None
    title_patterns = [
        r'<h1><span class="[^"]*" itemprop="name">(.*?)</span>',
        r'<h1[^>]*>.*?<span[^>]*itemprop="name"[^>]*>(.*?)</span>',
        r'<meta property="og:title" content="([^"]+)"',
    ]
    for pattern in title_patterns:
        match = re.search(pattern, source, re.S)
        if match:
            title = html.unescape(re.sub(r"<[^>]+>", "", match.group(1))).strip()
            break

    rating = None
    count = None
    count_text = None
    match = re.search(r'itemprop="starRating"(.{0,1000}?)<div class="g1rdde">([^<]+)', source, re.S)
    if match:
        count_text = html.unescape(match.group(2)).replace("\u00a0", " ").strip()
        rating_match = re.search(r'<span aria-hidden="true">([0-9]+(?:[.,][0-9]+)?)</span>', match.group(1))
        if rating_match:
            rating = float(rating_match.group(1).replace(",", "."))
        count = parse_count(count_text)
    if rating is None:
        meta_rating = re.search(r'itemprop="ratingValue"\s+content="([0-9.,]+)"', source)
        if meta_rating:
            rating = float(meta_rating.group(1).replace(",", "."))
    if count is None:
        meta_count = re.search(r'itemprop="ratingCount"\s+content="([0-9,.]+)"', source)
        if meta_count:
            count = parse_count(meta_count.group(1))

    if not title and "<html" not in source.lower():
        return {"ok": False, "rating": None, "rating_count": None, "rating_count_text": None, "title": None, "error": "Google Play 页面内容无法识别"}
    return {"ok": True, "rating": rating, "rating_count": count, "rating_count_text": count_text, "title": title, "error": None}


def parse_apple(body: bytes) -> dict:
    try:
        payload = json.loads(body.decode("utf-8", errors="ignore"))
    except Exception as exc:
        return {"ok": False, "available": False, "rating": None, "rating_count": None, "error": f"Apple JSON 解析失败: {exc}"}
    results = payload.get("results") or []
    if not results:
        return {"ok": True, "available": False, "rating": None, "rating_count": None, "error": None}
    item = results[0]
    count = int(item.get("userRatingCount") or 0)
    rating = item.get("averageUserRating")
    return {
        "ok": True,
        "available": True,
        "rating": float(rating) if rating is not None else None,
        "rating_count": count,
        "rating_source": "itunes_lookup",
        "track_id": item.get("trackId"),
        "track_name": item.get("trackName"),
        "bundle_id": item.get("bundleId"),
        "url": item.get("trackViewUrl"),
        "error": None,
    }


def parse_apple_page(body: bytes) -> dict:
    """Read the public App Store page when iTunes Lookup transiently reports zero ratings."""
    source = body.decode("utf-8", errors="ignore")
    scripts = re.findall(
        r'<script[^>]+type="application/ld\+json"[^>]*>(.*?)</script>',
        source,
        re.S | re.I,
    )
    for raw in scripts:
        try:
            data = json.loads(html.unescape(raw).strip())
        except Exception:
            continue
        stack = [data]
        while stack:
            node = stack.pop()
            if isinstance(node, dict):
                if node.get("@type") == "AggregateRating":
                    rating = node.get("ratingValue")
                    count = node.get("reviewCount", node.get("ratingCount"))
                    rating_count = parse_count(str(count)) if count is not None else None
                    try:
                        rating_value = float(str(rating).replace(",", "."))
                    except (TypeError, ValueError):
                        rating_value = None
                    if rating_value is not None and rating_value > 0 and rating_count:
                        return {
                            "ok": True,
                            "available": True,
                            "rating": rating_value,
                            "rating_count": rating_count,
                            "rating_source": "app_store_page",
                            "error": None,
                        }
                stack.extend(node.values())
            elif isinstance(node, list):
                stack.extend(node)
    return {
        "ok": False,
        "available": True,
        "rating": None,
        "rating_count": None,
        "error": "App Store 页面未找到评分",
    }


def google_url(android_id: str, cc: str) -> str:
    gl = GL.get(cc, cc)
    hl = HL.get(cc, "en")
    return f"https://play.google.com/store/apps/details?id={urllib.parse.quote(android_id)}&hl={urllib.parse.quote(hl)}&gl={urllib.parse.quote(gl)}"


def apple_url(game: dict, cc: str, result: dict | None = None) -> str:
    if result and result.get("url"):
        return result["url"]
    store_cc = APPLE_CC.get(cc, cc.lower())
    return f"https://apps.apple.com/{store_cc}/app/{game['ios_slug']}/id{game['ios_id']}?uo=4"


def scrape(logger: logging.Logger) -> dict:
    started = now_local()
    result = {
        "version": 1,
        "fetched_at": started.isoformat(timespec="seconds"),
        "games": {},
        "metrics": {"google_ok": 0, "google_total": 0, "apple_ok": 0, "apple_total": 0},
        "errors": [],
    }
    for game in GAMES:
        gdata = {
            "name": game["name"],
            "android_id": game["android"],
            "ios_bundle": game["ios_bundle"],
            "ios_id": game["ios_id"],
            "countries": {},
            "ru": {},
        }
        logger.info("开始抓取 %s", game["name"])
        for cc in game["countries"]:
            gurl = google_url(game["android"], cc)
            result["metrics"]["google_total"] += 1
            body, meta = fetch(gurl)
            if body is None:
                g_rec = {"ok": False, "rating": None, "rating_count": None, "rating_count_text": None, "error": meta["error"]}
                result["errors"].append(f"{game['name']} {cc} Google Play: {meta['error']}")
            else:
                g_rec = parse_gplay(body)
                if g_rec.get("ok"):
                    result["metrics"]["google_ok"] += 1
                else:
                    result["errors"].append(f"{game['name']} {cc} Google Play: {g_rec.get('error')}")
            logger.info("%s %s Google Play: %s / %s", game["name"], cc, g_rec.get("rating"), g_rec.get("rating_count"))
            time.sleep(0.55)

            apple_cc = APPLE_CC.get(cc, cc.lower())
            iurl = f"https://itunes.apple.com/lookup?bundleId={urllib.parse.quote(game['ios_bundle'])}&country={urllib.parse.quote(apple_cc)}&entity=software"
            result["metrics"]["apple_total"] += 1
            ibody, imeta = fetch(iurl, timeout=25)
            if ibody is None:
                a_rec = {"ok": False, "available": False, "rating": None, "rating_count": None, "error": imeta["error"]}
                result["errors"].append(f"{game['name']} {cc} App Store: {imeta['error']}")
            else:
                a_rec = parse_apple(ibody)
                if a_rec.get("ok"):
                    result["metrics"]["apple_ok"] += 1
                    if a_rec.get("available") and int(a_rec.get("rating_count") or 0) == 0:
                        page_url = apple_url(game, cc, a_rec)
                        page_body, _ = fetch(page_url, timeout=25)
                        if page_body is not None:
                            page_rec = parse_apple_page(page_body)
                            if page_rec.get("rating") is not None and int(page_rec.get("rating_count") or 0) > 0:
                                a_rec.update({
                                    "rating": page_rec["rating"],
                                    "rating_count": page_rec["rating_count"],
                                    "rating_source": "app_store_page",
                                })
                                a_rec.pop("stale", None)
                                logger.info(
                                    "%s %s App Store page fallback: %s / %s",
                                    game["name"], cc, a_rec["rating"], a_rec["rating_count"],
                                )
                else:
                    result["errors"].append(f"{game['name']} {cc} App Store: {a_rec.get('error')}")
            logger.info("%s %s App Store: %s / %s", game["name"], cc, a_rec.get("rating"), a_rec.get("rating_count"))
            gdata["countries"][cc] = {
                "country": cc,
                "country_name": COUNTRY_NAMES.get(cc, cc),
                "google_url": gurl,
                "apple_url": apple_url(game, cc, a_rec),
                "google": g_rec,
                "apple": a_rec,
            }
            time.sleep(0.2)


        ru_gurl = google_url(game["ru_android"], "RU")
        result["metrics"]["google_total"] += 1
        body, meta = fetch(ru_gurl)
        if body is None:
            ru_g = {"ok": False, "rating": None, "rating_count": None, "rating_count_text": None, "error": meta["error"]}
            result["errors"].append(f"{game['name']} RU Google Play: {meta['error']}")
        else:
            ru_g = parse_gplay(body)
            if ru_g.get("ok"):
                result["metrics"]["google_ok"] += 1
            else:
                result["errors"].append(f"{game['name']} RU Google Play: {ru_g.get('error')}")
        logger.info("%s RU Google Play: %s / %s", game["name"], ru_g.get("rating"), ru_g.get("rating_count"))
        gdata["ru"] = {
            "country": "RU",
            "country_name": COUNTRY_NAMES["RU"],
            "google_url": ru_gurl,
            "apple_url": None,
            "google": ru_g,
            "apple": {"ok": True, "available": False, "rating": None, "rating_count": None, "error": None, "unavailable_reason": "App Store 在 RU 区不可用"},
        }
        result["games"][game["key"]] = gdata
        time.sleep(0.3)
    result["finished_at"] = now_local().isoformat(timespec="seconds")
    result["duration_seconds"] = round((now_local() - started).total_seconds(), 1)
    return result


def workbook_previous() -> dict:
    from openpyxl import load_workbook

    files = sorted(ROOT.glob("商店评分汇总_????-??-??.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        return {}
    wb = load_workbook(files[0], data_only=True)
    previous = {"games": {}}
    mapping = {"Frontier City": "frontier", "Vice Throne": "vice"}
    for ws in wb.worksheets:
        key = mapping.get(ws.title)
        if not key:
            continue
        game = {"countries": {}, "ru": {}}
        for row in range(2, ws.max_row + 1):
            cc = ws.cell(row, 1).value
            if not cc or cc not in COUNTRY_NAMES:
                continue
            g_cell = ws.cell(row, 2).value
            a_cell = ws.cell(row, 3).value
            google = float(g_cell) if isinstance(g_cell, (int, float)) else None
            apple = float(a_cell) if isinstance(a_cell, (int, float)) else None
            item = {
                "google": {"ok": True, "rating": google, "rating_count": None},
                "apple": {"ok": True, "rating": apple, "rating_count": None},
            }
            if cc == "RU":
                game["ru"] = item
            else:
                game["countries"][cc] = item
        previous["games"][key] = game
    return previous


def load_previous() -> dict:
    if LATEST_JSON.exists():
        try:
            return json.loads(LATEST_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass
    return workbook_previous()


def previous_item(previous: dict, key: str, cc: str) -> dict | None:
    game = (previous.get("games") or {}).get(key) or {}
    if cc == "RU":
        return game.get("ru")
    return (game.get("countries") or {}).get(cc)


def apply_fallbacks(current: dict, previous: dict, logger: logging.Logger) -> None:
    for game in GAMES:
        key = game["key"]
        gdata = current["games"][key]
        for cc in [*game["countries"], "RU"]:
            current_item = gdata["ru"] if cc == "RU" else gdata["countries"][cc]
            old = previous_item(previous, key, cc) or {}
            for store in ("google", "apple"):
                rec = current_item.get(store) or {}
                old_rec = old.get(store) or {}
                if (
                    store == "apple"
                    and rec.get("ok")
                    and int(rec.get("rating_count") or 0) == 0
                    and old_rec.get("rating") is not None
                    and int(old_rec.get("rating_count") or 0) > 0
                ):
                    rec["rating"] = old_rec.get("rating")
                    rec["rating_count"] = old_rec.get("rating_count")
                    rec["stale"] = True
                    logger.warning("%s %s App Store 返回 0 条评分，沿用上次数据 %s", game["name"], cc, rec["rating"])
                    continue
                if rec.get("ok"):
                    continue
                if old_rec.get("rating") is not None or old_rec.get("rating_count") is not None:
                    rec["rating"] = old_rec.get("rating")
                    rec["rating_count"] = old_rec.get("rating_count")
                    rec["stale"] = True
                    logger.warning("%s %s %s 抓取失败，沿用上次数据 %s", game["name"], cc, store, rec["rating"])


def r1(value) -> float | None:
    if value is None:
        return None
    try:
        value = float(value)
    except Exception:
        return None
    if value <= 0:
        return None
    return float(Decimal(str(value)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))


def show_rating(value) -> str:
    value = r1(value)
    return "暂无评分" if value is None else f"{value:.1f}"


def get_rating(item: dict | None, store: str):
    return r1(((item or {}).get(store) or {}).get("rating"))


def get_count(item: dict | None, store: str):
    return ((item or {}).get(store) or {}).get("rating_count")



def collect_change_items(previous: dict, current: dict) -> dict[str, list[dict]]:
    all_changes: dict[str, list[dict]] = {}
    for game in GAMES:
        key = game["key"]
        cgame = current["games"][key]
        items: list[dict] = []
        for cc in [*game["countries"], "RU"]:
            old = previous_item(previous, key, cc)
            new = cgame["ru"] if cc == "RU" else cgame["countries"][cc]
            for store in ("google", "apple"):
                old_value = get_rating(old, store)
                new_value = get_rating(new, store)
                if old_value == new_value:
                    continue
                if old_value is None and new_value is None:
                    continue
                items.append({
                    "country": cc,
                    "country_name": COUNTRY_NAMES.get(cc, cc),
                    "store": store,
                    "old": old_value,
                    "new": new_value,
                    "count": get_count(new, store),
                })
        all_changes[key] = items
    return all_changes


def _compact_changes(items: list[dict]) -> str:
    if not items:
        return "各国家/地区评分与上次一致"
    parts = []
    for item in items[:8]:
        cc, old, new, count = item["country"], item["old"], item["new"], item.get("count")
        if old is None and new is not None:
            text = f"{cc} 新增 {new:.1f}"
        elif old is not None and new is None:
            text = f"{cc} 转为暂无评分"
        else:
            text = f"{cc} {old:.1f}→{new:.1f}"
        if item["store"] == "apple" and new is not None:
            text += f"（{count or 0} 条）"
        parts.append(text)
    visible = "；".join(parts)
    if len(items) > 8:
        visible += f"；另有 {len(items) - 8} 个地区变化"
    return visible + "。"


def build_changes(previous: dict, current: dict, key: str, items: list[dict] | None = None) -> list[str]:
    game = next(g for g in GAMES if g["key"] == key)
    cgame = current["games"][key]
    if items is None:
        items = collect_change_items(previous, current)[key]
    google_items = [item for item in items if item["store"] == "google" and item["country"] != "RU"]
    apple_items = [item for item in items if item["store"] == "apple"]

    old_ru = previous_item(previous, key, "RU")
    new_ru = cgame["ru"]
    old_ru_g, new_ru_g = get_rating(old_ru, "google"), get_rating(new_ru, "google")
    ru_count = get_count(new_ru, "google")
    if old_ru_g is None and new_ru_g is not None:
        ru_text = f"Google Play 新增 {new_ru_g:.1f}（{ru_count or 0} 条）"
    elif old_ru_g is not None and new_ru_g is None:
        ru_text = "Google Play 转为暂无评分"
    elif old_ru_g is not None and new_ru_g is not None and abs(old_ru_g - new_ru_g) >= 0.05:
        ru_text = f"Google Play {old_ru_g:.1f}→{new_ru_g:.1f}（{ru_count or 0} 条）"
    else:
        ru_text = f"Google Play {show_rating(new_ru_g)}（{ru_count or 0} 条）与上次一致"
    ru_text += "；App Store 在 RU 区不可用。"

    metrics = current["metrics"]
    status_text = f"Google Play {metrics['google_ok']}/{metrics['google_total']}、App Store {metrics['apple_ok']}/{metrics['apple_total']} 成功"
    stale = 0
    for cc in [*game["countries"], "RU"]:
        item = cgame["ru"] if cc == "RU" else cgame["countries"][cc]
        for store in ("google", "apple"):
            if (item.get(store) or {}).get("stale"):
                stale += 1
    status_text += f"；{stale} 项沿用上次数据" if stale else "；无异常"
    status_text += "。"
    return [
        f"\u2022 Google Play：{_compact_changes(google_items)}",
        f"\u2022 iOS：{_compact_changes(apple_items)}",
        f"\u2022 RU：{ru_text}",
        f"\u2022 更新状态：{status_text}",
    ]


def style_workbook(wb: Workbook, changes: dict[str, list[str]], current: dict) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="等线", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="等线", size=11)
    muted_font = Font(name="等线", size=11, color="808080", italic=True)
    link_font = Font(name="等线", size=11, color="0563C1", underline="single")
    note_title_font = Font(name="等线", size=11, bold=True, color="1F4E78")
    note_font = Font(name="等线", size=10.5, color="404040")
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    alt_fill = PatternFill("solid", fgColor="F5F8FC")
    none_fill = PatternFill("solid", fgColor="F2F2F2")

    for game in GAMES:
        key = game["key"]
        ws = wb[game["name"]]
        headers = ["国家/地区", "谷歌评分", "iOS评分", "Google Play链接", "App Store链接"]
        for col, text in enumerate(headers, 1):
            cell = ws.cell(1, col, text)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = center
        cgame = current["games"][key]
        rows = [(cc, cgame["countries"][cc]) for cc in game["countries"]] + [("RU", cgame["ru"])]
        for ridx, (cc, item) in enumerate(rows, start=2):
            for col in range(1, 6):
                cell = ws.cell(ridx, col)
                cell.border = border
                cell.font = body_font
                cell.alignment = center
                if ridx % 2 == 1:
                    cell.fill = alt_fill
            ws.cell(ridx, 1, cc)
            g_value = get_rating(item, "google")
            a_value = get_rating(item, "apple")
            if g_value is None:
                cell = ws.cell(ridx, 2, "暂无评分")
                cell.font = muted_font
                cell.fill = none_fill
            else:
                cell = ws.cell(ridx, 2, g_value)
                cell.number_format = "0.0"
            if a_value is None:
                cell = ws.cell(ridx, 3, "暂无评分")
                cell.font = muted_font
                cell.fill = none_fill
            else:
                cell = ws.cell(ridx, 3, a_value)
                cell.number_format = "0.0"
            dcell = ws.cell(ridx, 4, f"Google Play ({cc})")
            dcell.hyperlink = item["google_url"]
            dcell.font = link_font
            if cc == "RU":
                ecell = ws.cell(ridx, 5, "App Store（RU）不可用")
                ecell.font = muted_font
            else:
                ecell = ws.cell(ridx, 5, f"App Store ({cc})")
                ecell.hyperlink = item.get("apple_url")
                ecell.font = link_font


        last_row = len(rows) + 1
        widths = {"A": 12, "B": 14, "C": 14, "D": 21, "E": 20}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        ws.row_dimensions[1].height = 30
        for row in range(2, last_row + 1):
            ws.row_dimensions[row].height = 24
        ws.row_dimensions[last_row + 1].height = 8
        note_start = last_row + 2
        title = f"{game['name']} 数据变化说明（{current['fetched_at'][:16].replace('T', ' ')}）"
        ws.merge_cells(start_row=note_start, start_column=1, end_row=note_start, end_column=5)
        tcell = ws.cell(note_start, 1, title)
        tcell.font = note_title_font
        tcell.alignment = left_wrap
        ws.row_dimensions[note_start].height = 26
        for offset, text in enumerate(changes[key], start=1):
            row = note_start + offset
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            cell = ws.cell(row, 1, text)
            cell.font = note_font
            cell.alignment = left_wrap
            cell.fill = PatternFill("solid", fgColor="F7FAFD")
            for col in range(1, 6):
                ws.cell(row, col).border = border
            ws.row_dimensions[row].height = 42 if len(text) > 65 else 30
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:E{last_row}"
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 95
        ws.print_options.horizontalCentered = True
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True


    notes = wb.create_sheet("数据说明")
    notes.sheet_view.showGridLines = False
    notes.column_dimensions["A"].width = 24
    notes.column_dimensions["B"].width = 108
    notes["A1"] = "商店评分自动更新说明"
    notes["A1"].font = Font(name="等线", size=16, bold=True, color="1F4E78")
    notes.merge_cells("A1:B1")
    notes.row_dimensions[1].height = 28
    metrics = current["metrics"]
    info = [
        ("更新时间", current["fetched_at"][:16].replace("T", " ") + "（Asia/Shanghai）"),
        ("更新频率", "每 4 小时自动抓取并刷新 Excel 与本地评分看板。"),
        ("Google Play 评分", "取对应国家/地区 Google Play 商店页面公开显示的星级；没有公开地区评分时填“暂无评分”，不使用全球评分或默认评分回退。"),
        ("App Store 评分", "取 Apple iTunes Lookup 对应国家/地区的平均用户评分；0 条评分时填“暂无评分”。"),
        ("链接核对", "Google Play 和 App Store 链接逐行对应当前国家/地区：Google Play 使用对应 gl 参数（UK 使用 GB），App Store 使用对应国家商店路径。"),
        ("抓取状态", f"Google Play {metrics['google_ok']}/{metrics['google_total']}、App Store {metrics['apple_ok']}/{metrics['apple_total']} 成功；失败项会沿用上次数据并在看板中标记。"),
        ("Frontier City", "Android 包名：com.camelgames.xcity（RU：com.camelgames.xcityru）；App Store ID：6751943304。"),
        ("Vice Throne", "Android 包名：com.camelgames.mnnew（RU：com.camelgames.mnnew.ru）；App Store ID：6798176762。"),
        ("注意事项", "评分会随商店数据更新；没有公开地区评分的地区不能用其他地区或总量推算。本地看板文件：评分看板.html。"),
    ]
    for row, (label, text) in enumerate(info, start=3):
        a = notes.cell(row, 1, label)
        b = notes.cell(row, 2, text)
        a.font = Font(name="等线", size=11, bold=True, color="1F4E78")
        b.font = Font(name="等线", size=11)
        a.alignment = Alignment(horizontal="left", vertical="top")
        b.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        a.fill = PatternFill("solid", fgColor="EAF1F8")
        for cell in (a, b):
            cell.border = border
        notes.row_dimensions[row].height = 58 if label in ("Google Play 评分", "链接核对", "抓取状态") else 34


def save_workbook(current: dict, changes: dict[str, list[str]]) -> tuple[Path, Path]:
    wb = Workbook()
    wb.remove(wb.active)
    for game in GAMES:
        wb.create_sheet(game["name"])
    style_workbook(wb, changes, current)
    wb.active = 0
    day_path = ROOT / f"商店评分汇总_{now_local().strftime('%Y-%m-%d')}.xlsx"
    temp_path = ROOT / f".ratings_{os.getpid()}.xlsx"
    wb.save(temp_path)
    os.replace(temp_path, LATEST_XLSX)
    temp_day = ROOT / f".ratings_day_{os.getpid()}.xlsx"
    wb.save(temp_day)
    os.replace(temp_day, day_path)
    return day_path, LATEST_XLSX



def _snapshot_ratings(current: dict) -> dict:
    ratings: dict[str, dict] = {}
    for game in GAMES:
        key = game["key"]
        game_data = current["games"][key]
        rows = {}
        for cc in [*game["countries"], "RU"]:
            item = game_data["ru"] if cc == "RU" else game_data["countries"][cc]
            rows[cc] = [
                get_rating(item, "google"),
                get_count(item, "google"),
                get_rating(item, "apple"),
                get_count(item, "apple"),
            ]
        ratings[key] = rows
    return ratings


def compact_snapshot(current: dict, change_items: dict[str, list[dict]]) -> dict:
    flat_items = []
    for key in ("frontier", "vice"):
        product_name = next(game["name"] for game in GAMES if game["key"] == key)
        for item in change_items.get(key, []):
            flat_items.append({**item, "product": key, "product_name": product_name})
    return {
        "at": current["fetched_at"],
        "changes": len(flat_items),
        "change_items": flat_items,
        "metrics": current["metrics"],
        "errors": len(current.get("errors", [])),
        "ratings": _snapshot_ratings(current),
    }


def _normalise_legacy_snapshot(item: dict) -> dict:
    return {
        "at": item.get("at"),
        "changes": item.get("changes", 0),
        "change_items": [],
        "metrics": {
            "google_ok": item.get("google_ok"),
            "google_total": item.get("google_total"),
            "apple_ok": item.get("apple_ok"),
            "apple_total": item.get("apple_total"),
        },
        "errors": item.get("errors", 0),
        "legacy": True,
    }


def load_history() -> list[dict]:
    if not HISTORY_JSON.exists():
        return []
    try:
        value = json.loads(HISTORY_JSON.read_text(encoding="utf-8"))
        if isinstance(value, dict):
            values = value.get("snapshots") or []
        elif isinstance(value, list):
            values = value
        else:
            values = []
        snapshots = []
        for item in values:
            if not isinstance(item, dict) or not item.get("at"):
                continue
            snapshots.append(item if isinstance(item.get("ratings"), dict) else _normalise_legacy_snapshot(item))
        snapshots.sort(key=lambda item: item["at"])
        return snapshots
    except Exception:
        return []


def history_document(snapshots: list[dict]) -> dict:
    return {
        "version": 2,
        "updated_at": snapshots[-1]["at"] if snapshots else None,
        "snapshots": snapshots,
    }


def change_summary(change_items: dict[str, list[dict]]) -> int:
    return sum(len(items) for items in change_items.values())


def html_data(current: dict, changes: dict[str, list[str]], history: list[dict], day_path: Path, status: dict, change_items: dict[str, list[dict]]) -> dict:
    payload = {
        "updated_at": current["fetched_at"],
        "metrics": current["metrics"],
        "errors": current.get("errors", []),
        "changes": changes,
        "latest_changes": change_items,
        "history": history[-20:],
        "history_total": len(history),
        "xlsx_file": LATEST_XLSX.name,
        "products": [],
    }
    for game in GAMES:
        key = game["key"]
        gdata = current["games"][key]
        change_lookup = {(item["country"], item["store"]): item for item in change_items.get(key, [])}
        rows = []
        for cc in [*game["countries"], "RU"]:
            item = gdata["ru"] if cc == "RU" else gdata["countries"][cc]
            rows.append({
                "country": cc,
                "country_name": COUNTRY_NAMES.get(cc, cc),
                "google_rating": get_rating(item, "google"),
                "google_count": get_count(item, "google"),
                "ios_rating": get_rating(item, "apple"),
                "ios_count": get_count(item, "apple"),
                "google_url": item.get("google_url"),
                "apple_url": item.get("apple_url"),
                "google_stale": bool((item.get("google") or {}).get("stale")),
                "ios_stale": bool((item.get("apple") or {}).get("stale")),
                "ios_unavailable": cc == "RU",
                "google_change": change_lookup.get((cc, "google")),
                "ios_change": change_lookup.get((cc, "apple")),
            })
        rated_google = sum(1 for row in rows if row["google_rating"] is not None)
        rated_ios = sum(1 for row in rows if row["ios_rating"] is not None)
        payload["products"].append({
            "key": key,
            "name": game["name"],
            "rows": rows,
            "rated_google": rated_google,
            "rated_ios": rated_ios,
            "total": len(rows),
            "changes": changes[key],
        })
    return payload


def write_dashboard(payload: dict) -> None:
    template_path = ROOT / "dashboard_template.html"
    if not template_path.exists():
        raise FileNotFoundError(f"缺少看板模板：{template_path}")
    template = template_path.read_text(encoding="utf-8")
    content = template.replace("__DATA__", json.dumps(payload, ensure_ascii=False).replace("</", "<\\/"))
    content = content.replace("__PAGE_VERSION__", PAGE_VERSION)
    temp = ROOT / ".评分看板.html.tmp"
    temp.write_text(content, encoding="utf-8")
    os.replace(temp, DASHBOARD)
    VERSION_JSON.write_text(
        json.dumps({"version": PAGE_VERSION, "built_at": payload.get("updated_at")}, ensure_ascii=False),
        encoding="utf-8",
    )


def main() -> int:
    logger = log_setup()
    started = now_local()
    logger.info("========== 自动更新开始 ==========")
    previous = load_previous()
    try:
        current = scrape(logger)
        metrics = current["metrics"]
        if metrics["google_ok"] == 0 and metrics["apple_ok"] == 0:
            raise RuntimeError("Google Play 与 App Store 全部抓取失败，保留上次结果")
        apply_fallbacks(current, previous, logger)
        change_items = collect_change_items(previous, current)
        changes = {game["key"]: build_changes(previous, current, game["key"], change_items[game["key"]]) for game in GAMES}
        day_path, latest_path = save_workbook(current, changes)
        history = load_history()
        snapshot = compact_snapshot(current, change_items)
        if history and history[-1]["at"] == snapshot["at"]:
            history[-1] = snapshot
        else:
            history.append(snapshot)
        history.sort(key=lambda item: item["at"])
        atomic_json(HISTORY_JSON, history_document(history))
        atomic_json(LATEST_JSON, current)
        status = {
            "ok": True,
            "updated_at": current["fetched_at"],
            "duration_seconds": current.get("duration_seconds"),
            "google_ok": metrics["google_ok"],
            "google_total": metrics["google_total"],
            "apple_ok": metrics["apple_ok"],
            "apple_total": metrics["apple_total"],
            "errors": current.get("errors", []),
            "xlsx": str(day_path),
            "dashboard": str(DASHBOARD),
        }
        atomic_json(STATUS_JSON, status)
        write_dashboard(html_data(current, changes, history, day_path, status, change_items))
        logger.info("完成：%s；变化 %s 项；耗时 %.1fs", day_path, change_summary(change_items), current.get("duration_seconds") or 0)
        print(json.dumps({"ok": True, "xlsx": str(day_path), "dashboard": str(DASHBOARD), "changes": change_summary(change_items), "errors": len(current.get("errors", []))}, ensure_ascii=False))
        return 0
    except Exception as exc:
        logger.exception("自动更新失败：%s", exc)
        status = {
            "ok": False,
            "updated_at": now_local().isoformat(timespec="seconds"),
            "message": str(exc),
            "duration_seconds": round((now_local() - started).total_seconds(), 1),
        }
        atomic_json(STATUS_JSON, status)
        print(json.dumps(status, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main())
