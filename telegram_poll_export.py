"""
Lấy poll từ Telegram, fetch điểm từ ESPN, tính bảng xếp hạng.

Cài đặt:
    pip install telethon openpyxl pandas requests

Cách dùng:
    python telegram_poll_export.py                     # chạy đủ 3 bước
    python telegram_poll_export.py --skip-poll         # bỏ bước lấy poll Telegram
    python telegram_poll_export.py --fetch-scores      # tự fetch điểm từ ESPN
    python telegram_poll_export.py --skip-ranking      # bỏ bước tính BXH
    python telegram_poll_export.py --skip-poll --fetch-scores   # chỉ fetch điểm + BXH
"""

import argparse
import asyncio
import json
import os
import re
from datetime import datetime, timedelta
from difflib import get_close_matches

import openpyxl
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

# ==================== CẤU HÌNH ====================
API_ID   = 34813147
API_HASH = "ec672e74a5ffa437c7eea4a144c0b423"
PHONE    = "+84 869 322 628"

CHAT_ID  = -1003904575691
TOPIC_ID = 3

OUTPUT_FILE  = "poll_results_1.xlsx"
SCORES_FILE  = "scores.json"
RANKING_FILE = "ranking.xlsx"

# World Cup 2026 bắt đầu
WC_START_DATE = datetime(2026, 6, 11)
WC_DAYS       = 60
# ===================================================

HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
ROW_FILLS   = [
    PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"),
    PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid"),
]
EMPTY_FILL  = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")

ESPN_URL = "https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/scoreboard"


# ─────────────────────────────────────────────────────────────
# BƯỚC 1: Lấy poll từ Telegram
# ─────────────────────────────────────────────────────────────

async def _get_voters(client, chat, message, option_bytes, option_text):
    from telethon.tl.functions.messages import GetPollVotesRequest
    from telethon.errors import PollVoteRequiredError
    voters, offset = {}, None
    while True:
        try:
            res = await client(GetPollVotesRequest(
                peer=chat, id=message.id,
                option=option_bytes, offset=offset, limit=100,
            ))
        except PollVoteRequiredError:
            # Tài khoản chưa vote poll này → không xem được voter list
            break
        if not res.users:
            break
        for u in res.users:
            voters[u.id] = {
                "name":     " ".join(filter(None, [u.first_name, u.last_name])),
                "username": f"@{u.username}" if u.username else "",
                "option":   option_text,
            }
        offset = res.next_offset
        if not offset:
            break
    return voters


async def fetch_polls():
    from telethon import TelegramClient
    from telethon.tl.types import MessageMediaPoll

    client = TelegramClient("session_poll", API_ID, API_HASH)
    await client.start(phone=PHONE)
    print("Đã đăng nhập Telegram.")

    chat = await client.get_entity(CHAT_ID)
    print(f"Group: {getattr(chat, 'title', CHAT_ID)}")

    polls = []
    async for msg in client.iter_messages(chat, reply_to=TOPIC_ID, limit=None):
        if msg.media and isinstance(msg.media, MessageMediaPoll):
            polls.append(msg)

    try:
        root = await client.get_messages(chat, ids=TOPIC_ID)
        if root and isinstance(root.media, MessageMediaPoll):
            polls.insert(0, root)
    except Exception:
        pass

    polls.sort(key=lambda m: m.id)

    if not polls:
        print("Không tìm thấy poll nào trong topic.")
        await client.disconnect()
        return

    print(f"Tìm thấy {len(polls)} poll(s).\n")

    poll_data = []
    members   = {}

    for idx, msg in enumerate(polls, 1):
        poll     = msg.media.poll
        results  = msg.media.results
        question = poll.question.text if hasattr(poll.question, "text") else str(poll.question)
        print(f"[Poll {idx}] {question}")

        all_voters = {}

        for answer in poll.answers:
            text  = answer.text.text if hasattr(answer.text, "text") else str(answer.text)
            count = 0
            if results and results.results:
                for r in results.results:
                    if r.option == answer.option:
                        count = r.voters
                        break

            if poll.public_voters:
                v = await _get_voters(client, chat, msg, answer.option, text)
                all_voters.update(v)
                if v:
                    print(f"  ✓ {text}: {count} phiếu ({len(v)} tên)")
                else:
                    print(f"  ⚠ {text}: {count} phiếu (chưa vote → không lấy được tên)")
            else:
                print(f"  ✗ {text}: {count} phiếu (ẩn danh)")

        for uid, info in all_voters.items():
            if uid not in members:
                members[uid] = {"name": info["name"], "username": info["username"]}

        poll_data.append({"question": question, "voters": all_voters})

    # Ghi Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kết quả vote"
    num_polls = len(poll_data)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2 + num_polls)
    c = ws.cell(row=1, column=1,
                value=f"Tổng hợp kết quả vote — Xuất lúc: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    c.font      = Font(bold=True, size=13, color="FFFFFF")
    c.fill      = HEADER_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.cell(row=2, column=1, value="STT").font     = HEADER_FONT
    ws.cell(row=2, column=1).fill                  = HEADER_FILL
    ws.cell(row=2, column=1).alignment             = Alignment(horizontal="center")
    ws.cell(row=2, column=2, value="Họ và tên").font = HEADER_FONT
    ws.cell(row=2, column=2).fill                    = HEADER_FILL
    ws.cell(row=2, column=2).alignment               = Alignment(horizontal="center")

    for col_idx, pd_item in enumerate(poll_data, 3):
        c = ws.cell(row=2, column=col_idx, value=pd_item["question"])
        c.font = HEADER_FONT; c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[2].height = 40

    member_list = list(members.items())
    for row_idx, (uid, info) in enumerate(member_list, 1):
        row  = row_idx + 2
        fill = ROW_FILLS[row_idx % 2]

        ws.cell(row=row, column=1, value=row_idx).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=1).fill = fill

        name_display = info["name"]
        if info["username"]:
            name_display += f"\n{info['username']}"
        c = ws.cell(row=row, column=2, value=name_display)
        c.fill = fill; c.alignment = Alignment(wrap_text=True)

        for col_idx, pd_item in enumerate(poll_data, 3):
            voter_info = pd_item["voters"].get(uid)
            if voter_info:
                c = ws.cell(row=row, column=col_idx, value=voter_info["option"])
                c.fill = fill; c.alignment = Alignment(horizontal="center", wrap_text=True)
            else:
                c = ws.cell(row=row, column=col_idx, value="")
                c.fill = EMPTY_FILL

    total_row = len(member_list) + 3
    ws.cell(row=total_row, column=1, value="Tổng").font          = Font(bold=True)
    ws.cell(row=total_row, column=2, value=f"{len(member_list)} người").font = Font(bold=True)
    for col_idx, pd_item in enumerate(poll_data, 3):
        c = ws.cell(row=total_row, column=col_idx, value=f"{len(pd_item['voters'])} phiếu")
        c.font = Font(bold=True); c.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 28
    for col_idx in range(3, 3 + num_polls):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 30

    wb.save(OUTPUT_FILE)
    print(f"\nDone! File: {OUTPUT_FILE}")
    await client.disconnect()


# ─────────────────────────────────────────────────────────────
# BƯỚC 2: Tự động fetch điểm từ ESPN
# ─────────────────────────────────────────────────────────────

def _parse_espn_events(data: dict) -> dict:
    """Trích xuất kết quả hoàn thành từ response ESPN."""
    results = {}
    for event in data.get("events", []):
        if not event.get("status", {}).get("type", {}).get("completed"):
            continue
        competitors = event["competitions"][0]["competitors"]
        home = next((c for c in competitors if c["homeAway"] == "home"), None)
        away = next((c for c in competitors if c["homeAway"] == "away"), None)
        if not home or not away:
            continue
        key = f"{home['team']['displayName']} vs {away['team']['displayName']}"
        results[key] = [int(home["score"]), int(away["score"])]
    return results


def _espn_fetch_all(start_date: datetime, days: int) -> dict:
    """
    Fetch kết quả từ ESPN.
    Thử 2 cách:
      1. Date-range (1 request): dates=YYYYMMDD-YYYYMMDD
      2. Per-day fallback nếu cách 1 không trả đủ dữ liệu
    Chỉ fetch đến hôm nay (bỏ qua ngày tương lai).
    """
    today     = datetime.now()
    end_date  = min(start_date + timedelta(days=days - 1), today)
    start_str = start_date.strftime("%Y%m%d")
    end_str   = end_date.strftime("%Y%m%d")

    # Cách 1: date range (1 request duy nhất)
    try:
        resp = requests.get(
            ESPN_URL,
            params={"dates": f"{start_str}-{end_str}", "limit": 200},
            timeout=15,
        )
        resp.raise_for_status()
        api_results = _parse_espn_events(resp.json())
        if api_results:
            print(f"  ESPN (range {start_str}–{end_str}): {len(api_results)} trận")
            return api_results
    except Exception as e:
        print(f"  ESPN range request lỗi: {e}")

    # Cách 2: per-day fallback
    print("  Thử per-day fallback...")
    api_results = {}
    current = start_date
    while current <= end_date:
        date_str = current.strftime("%Y%m%d")
        try:
            resp = requests.get(ESPN_URL, params={"dates": date_str, "limit": 50}, timeout=10)
            resp.raise_for_status()
            day_results = _parse_espn_events(resp.json())
            if day_results:
                print(f"  {date_str}: {len(day_results)} trận")
                api_results.update(day_results)
        except Exception as e:
            print(f"  Lỗi ngày {date_str}: {e}")
        current += timedelta(days=1)

    return api_results


def _fuzzy_match(target: str, candidates: list) -> str | None:
    if target in candidates:
        return target
    close = get_close_matches(target, candidates, n=1, cutoff=0.55)
    if close:
        return close[0]
    # Match từng đội riêng
    if " vs " in target:
        h, a = target.split(" vs ", 1)
        for cand in candidates:
            if " vs " not in cand:
                continue
            ch, ca = cand.split(" vs ", 1)
            if get_close_matches(h, [ch], n=1, cutoff=0.6) and \
               get_close_matches(a, [ca], n=1, cutoff=0.6):
                return cand
    return None


def _rebuild_scores_from_poll() -> dict:
    """Tạo lại scores dict từ poll_results_1.xlsx khi scores.json bị rỗng/corrupt."""
    if not os.path.exists(OUTPUT_FILE):
        return {}
    from openpyxl import load_workbook as _lw
    wb   = _lw(OUTPUT_FILE, data_only=True)
    ws   = wb["Kết quả vote"]
    hdrs = list(ws.iter_rows(values_only=True))[1]
    return {h: None for h in hdrs[2:] if h}


def fetch_scores():
    # Đọc scores.json — tạo lại nếu file không tồn tại hoặc rỗng/corrupt
    scores = {}
    if os.path.exists(SCORES_FILE):
        try:
            with open(SCORES_FILE, encoding="utf-8") as f:
                content = f.read().strip()
            if content:
                scores = json.loads(content)
        except (json.JSONDecodeError, ValueError):
            print(f"⚠ {SCORES_FILE} bị lỗi/rỗng — đang tạo lại từ {OUTPUT_FILE}...")

    if not scores:
        scores = _rebuild_scores_from_poll()
        if not scores:
            print(f"Chưa có dữ liệu. Chạy bước poll trước.")
            return
        with open(SCORES_FILE, "w", encoding="utf-8") as f:
            json.dump(scores, f, ensure_ascii=False, indent=2)
        print(f"Đã tạo lại {SCORES_FILE} với {len(scores)} trận.")

    # Sync các trận mới từ poll vào scores (phòng trường hợp kéo poll mới hơn)
    poll_matches = _rebuild_scores_from_poll()
    new_matches = [m for m in poll_matches if m not in scores]
    if new_matches:
        for m in new_matches:
            scores[m] = None
        with open(SCORES_FILE, "w", encoding="utf-8") as f:
            json.dump(scores, f, ensure_ascii=False, indent=2)
        print(f"Đã thêm {len(new_matches)} trận mới từ poll vào {SCORES_FILE}:")
        for m in new_matches:
            print(f"  + {m}")

    null_matches = [k for k, v in scores.items() if v is None]
    if not null_matches:
        print("Tất cả trận đã có điểm rồi.")
        return

    print(f"Cần tìm kết quả cho {len(null_matches)} trận...")
    print("Đang fetch từ ESPN...\n")

    api_results = _espn_fetch_all(WC_START_DATE, WC_DAYS)
    if not api_results:
        print("Không lấy được dữ liệu từ ESPN. Kiểm tra kết nối mạng.")
        return

    print(f"ESPN trả về {len(api_results)} trận đã hoàn thành.\n")

    updated, not_found = 0, []
    for match in null_matches:
        best = _fuzzy_match(match, list(api_results.keys()))
        if best:
            scores[match] = api_results[best]
            updated += 1
            label = f" (khớp với ESPN: '{best}')" if best != match else ""
            print(f"  ✓ {match}{label}: {api_results[best]}")
        else:
            not_found.append(match)
            print(f"  ✗ Không tìm thấy: {match}")

    with open(SCORES_FILE, "w", encoding="utf-8") as f:
        json.dump(scores, f, ensure_ascii=False, indent=2)
    print(f"\nĐã cập nhật {updated}/{len(null_matches)} trận → {SCORES_FILE}")

    # Luôn in toàn bộ tên trận ESPN đã fetch để dễ đối chiếu
    print(f"\n── Toàn bộ trận ESPN đã fetch ({len(api_results)}) ──")
    for k, v in sorted(api_results.items()):
        print(f"  {k}: {v}")

    if not_found:
        print(f"\n── Còn {len(not_found)} trận chưa khớp (điền tay vào {SCORES_FILE}) ──")
        for m in not_found:
            print(f"  - {m}")


# ─────────────────────────────────────────────────────────────
# BƯỚC 3: Tính bảng xếp hạng
# ─────────────────────────────────────────────────────────────

def _load_scores(match_names: list) -> dict:
    """Đọc scores.json, tạo template nếu chưa có. Trả về các trận đã có điểm."""
    raw = {}
    if os.path.exists(SCORES_FILE):
        try:
            with open(SCORES_FILE, encoding="utf-8") as f:
                content = f.read().strip()
            if content:
                raw = json.loads(content)
        except (json.JSONDecodeError, ValueError):
            print(f"⚠ {SCORES_FILE} bị lỗi/rỗng — tạo lại template.")

    if not raw:
        raw = {name: None for name in match_names}
        with open(SCORES_FILE, "w", encoding="utf-8") as f:
            json.dump(raw, f, ensure_ascii=False, indent=2)
        print(f"Đã tạo template: {SCORES_FILE}")
        print("Điền điểm số theo dạng [home, away] rồi chạy lại.")
        return {}

    # Thêm trận mới nếu có
    new_matches = [n for n in match_names if n not in raw]
    if new_matches:
        for name in new_matches:
            raw[name] = None
        with open(SCORES_FILE, "w", encoding="utf-8") as f:
            json.dump(raw, f, ensure_ascii=False, indent=2)
        print(f"Đã thêm {len(new_matches)} trận mới vào {SCORES_FILE} — điền điểm rồi chạy lại.")

    return {name: tuple(score) for name, score in raw.items() if score is not None}


def _resolve_team(team: str, home: str, away: str) -> str | None:
    """Khớp tên đội — exact trước, rồi prefix (xử lý 'Netherland', 'Bosnia', ...)."""
    if team == home: return home
    if team == away: return away
    for candidate in (home, away):
        if candidate.startswith(team) or team.startswith(candidate):
            return candidate
    return None


def _grade(pred, match, match_results):
    if pred is None or str(pred).strip() == "":
        return 0
    pred = str(pred).strip()
    m = re.match(r"(.+?)\s*([+-]\d+(?:\.\d+)?)$", pred)
    if not m:
        return 0
    team     = m.group(1).strip()
    handicap = float(m.group(2))
    home, away = match.split(" vs ")
    home_score, away_score = match_results[match]
    resolved = _resolve_team(team, home, away)
    if resolved == home:
        diff = home_score + handicap - away_score
    elif resolved == away:
        diff = away_score + handicap - home_score
    else:
        return 0
    return 1 if diff > 0 else 0


def generate_ranking():
    if not os.path.exists(OUTPUT_FILE):
        print(f"Chưa có {OUTPUT_FILE}. Chạy bước poll trước.")
        return

    wb = load_workbook(OUTPUT_FILE, data_only=True)
    ws = wb["Kết quả vote"]
    rows    = list(ws.iter_rows(values_only=True))
    headers = rows[1]

    match_names   = [h for h in headers[2:] if h]
    match_results = _load_scores(match_names)

    if not match_results:
        print("Chưa có điểm số. Dùng --fetch-scores hoặc điền thủ công vào scores.json.")
        return

    TOTAL  = len(match_results)
    scores = []

    for row in rows[2:]:
        if not row[1]:
            continue
        if row[0] == "Tổng":  # hàng tổng cộng cuối sheet
            continue
        name    = str(row[1]).split("\n")[0]
        correct = sum(
            _grade(row[col_idx], match, match_results)
            for col_idx, match in enumerate(headers[2:], start=2)
            if match in match_results
        )
        scores.append({
            "Tên":       name,
            "Đúng":      correct,
            "Sai":       TOTAL - correct,
            "Tỷ lệ (%)": round(correct / TOTAL * 100, 1),
        })

    df = pd.DataFrame(scores)
    df = df.sort_values(["Đúng", "Sai", "Tên"], ascending=[False, True, True]).reset_index(drop=True)
    df.index     += 1
    df.index.name = "Hạng"

    print(df.head(20).to_string())
    df.to_excel(RANKING_FILE)
    print(f"\nĐã lưu: {RANKING_FILE}")


# ─────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Telegram poll → Excel → BXH",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ví dụ:
  python telegram_poll_export.py                          # chạy đủ 3 bước
  python telegram_poll_export.py --skip-poll              # bỏ bước Telegram
  python telegram_poll_export.py --fetch-scores           # thêm bước tự fetch điểm ESPN
  python telegram_poll_export.py --skip-poll --fetch-scores --skip-ranking
        """,
    )
    parser.add_argument("--poll-file",  type=str, default=OUTPUT_FILE, help="File Excel kết quả poll (mặc định: poll_results_1.xlsx)")
    parser.add_argument("--skip-poll",         action="store_true", help="Bỏ qua bước lấy poll từ Telegram")
    parser.add_argument("--fetch-scores",      action="store_true", help="Tự động fetch điểm từ ESPN")
    parser.add_argument("--skip-fetch-scores", action="store_true", help="Bỏ qua bước fetch ESPN (alias)")
    parser.add_argument("--skip-ranking",      action="store_true", help="Bỏ qua bước tính bảng xếp hạng")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    if not args.skip_poll:
        print("=" * 50)
        print("BƯỚC 1: Lấy poll từ Telegram")
        print("=" * 50)
        asyncio.run(fetch_polls())

    if args.fetch_scores:
        print("\n" + "=" * 50)
        print("BƯỚC 2: Tự động fetch điểm từ ESPN")
        print("=" * 50)
        fetch_scores()

    if not args.skip_ranking:
        print("\n" + "=" * 50)
        print("BƯỚC 3: Tính bảng xếp hạng")
        print("=" * 50)
        generate_ranking()
