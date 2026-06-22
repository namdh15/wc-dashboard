"""
Tự động lấy kết quả các trận từ ESPN API và điền vào scores.json.
Cài đặt: pip install requests
"""

import json
import requests
from difflib import get_close_matches
from datetime import datetime, timedelta

SCORES_FILE = "scores.json"

# ESPN unofficial API — không cần API key
ESPN_URL = "https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/scoreboard"


def fetch_espn_results(start_date: datetime, days: int = 60) -> dict:
    """Lấy tất cả kết quả đã hoàn thành trong khoảng thời gian."""
    api_results = {}

    for i in range(days):
        date_str = (start_date + timedelta(days=i)).strftime("%Y%m%d")
        try:
            resp = requests.get(ESPN_URL, params={"dates": date_str, "limit": 50}, timeout=10)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            print(f"  Lỗi ngày {date_str}: {e}")
            continue

        for event in data.get("events", []):
            status = event.get("status", {}).get("type", {})
            if not status.get("completed"):
                continue

            competitors = event["competitions"][0]["competitors"]
            home = next((c for c in competitors if c["homeAway"] == "home"), None)
            away = next((c for c in competitors if c["homeAway"] == "away"), None)
            if not home or not away:
                continue

            key = f"{home['team']['displayName']} vs {away['team']['displayName']}"
            api_results[key] = [int(home["score"]), int(away["score"])]

    return api_results


def match_team_name(target: str, candidates: list[str], cutoff: float = 0.55) -> str | None:
    """Tìm tên trận gần nhất trong danh sách API (fuzzy match)."""
    # Thử exact trước
    if target in candidates:
        return target

    # Fuzzy match toàn bộ chuỗi "A vs B"
    close = get_close_matches(target, candidates, n=1, cutoff=cutoff)
    if close:
        return close[0]

    # Thử match từng đội riêng
    home_target, away_target = target.split(" vs ")
    for cand in candidates:
        if " vs " not in cand:
            continue
        home_cand, away_cand = cand.split(" vs ")
        home_ok = get_close_matches(home_target, [home_cand], n=1, cutoff=0.6)
        away_ok = get_close_matches(away_target, [away_cand], n=1, cutoff=0.6)
        if home_ok and away_ok:
            return cand

    return None


def fetch_and_update():
    # Đọc scores.json
    with open(SCORES_FILE, encoding="utf-8") as f:
        scores = json.load(f)

    null_matches = [k for k, v in scores.items() if v is None]
    if not null_matches:
        print("Tất cả trận đã có điểm rồi.")
        return

    print(f"Cần tìm kết quả cho {len(null_matches)} trận...")
    print("Đang fetch từ ESPN...\n")

    # FIFA World Cup 2026: bắt đầu ~11/6/2026
    start = datetime(2026, 6, 11)
    api_results = fetch_espn_results(start, days=60)

    if not api_results:
        print("Không lấy được dữ liệu từ ESPN. Kiểm tra kết nối mạng.")
        return

    print(f"ESPN trả về {len(api_results)} trận đã hoàn thành.\n")

    updated = 0
    not_found = []

    for match in null_matches:
        best = match_team_name(match, list(api_results.keys()))
        if best:
            scores[match] = api_results[best]
            updated += 1
            if best == match:
                print(f"  ✓ {match}: {api_results[best]}")
            else:
                print(f"  ✓ {match}")
                print(f"      → khớp với: {best}: {api_results[best]}")
        else:
            not_found.append(match)
            print(f"  ✗ Không tìm thấy: {match}")

    # Lưu lại
    with open(SCORES_FILE, "w", encoding="utf-8") as f:
        json.dump(scores, f, ensure_ascii=False, indent=2)

    print(f"\nĐã cập nhật {updated}/{len(null_matches)} trận → {SCORES_FILE}")

    if not_found:
        print(f"\nCòn {len(not_found)} trận chưa tìm thấy (điền tay):")
        for m in not_found:
            print(f"  - {m}")

        print("\nCác trận ESPN trả về (để đối chiếu tên):")
        for k, v in api_results.items():
            print(f"  {k}: {v}")


if __name__ == "__main__":
    fetch_and_update()
