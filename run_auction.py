# -*- coding: utf-8 -*-
"""
madangs API → 감정가/최저가/직전가를 '한글 표기' 그대로 추출
필드: court, case_num, apt_name, addr, area_m2, eval_price(감정가), low_price(최저가), last_price(직전가), bid_date
출력: madangs_min_kor.csv
"""

import re, requests, pandas as pd
from datetime import datetime
from collections import deque

URL = r"https://madangs.com/api/search/list?query%5Baddr%5D=11680%2B11215%2B11440%2B11650%2B11590%2B11200%2B11710%2B11470%2B11560%2B11170%2B11140&query%5Bcourt%5D=&query%5Basset_classification%5D=&query%5Bbuild_area_max%5D=0&query%5Bbuild_area_min%5D=0&query%5Beval_p_max%5D=1800000000&query%5Beval_p_min%5D=900000000&query%5Bg_use_type%5D=2000%2C2001%2C2007&query%5Bland_area_max%5D=0&query%5Bland_area_min%5D=0&query%5Blist_type%5D=1&query%5Blow_p_max%5D=0&query%5Blow_p_min%5D=0&query%5Bstate%5D=10&query%5Bshare%5D=2&query%5Bg_share%5D=2&query%5Bspecial%5D=003%2B011%2B007%2B006%2B008%2B002%2B015%2B009%2B004%2B010%2B001%2B005%2B012%2B016%2B017%2B018%2B022%2B023&query%5Bcontain_special%5D=0&query%5Buchal%5D=&query%5Buse_type%5D=2000&query%5Bsort%5D=view_desc&query%5Bstart_date%5D=&query%5Bend_date%5D=&page=1"

OUT_CSV = "madangs_min_kor.csv"

# --- K-apt 엑셀 연동 ---
MIN_HOUSEHOLD = 200 # 필터링할 최소 세대수 기준.
df_danzi = None

def safe_text(value):
    if isinstance(value, str):
        return value.strip()
    if pd.isna(value):
        return ""
    return str(value).strip()

def compact_text(value):
    return re.sub(r"\s+", "", safe_text(value))

def clean_apt_name(name):
    name = safe_text(name)
    if not name:
        return ""
    # 1. Handle (Dong, Name) format if present
    if "(" in name and ")" in name:
        m = re.search(r"\(([^)]+)\)", name)
        if m:
            parts = m.group(1).split(",")
            name = parts[-1].strip() if len(parts) > 1 else parts[0].strip()

    # 2. Sequential cleaning of typical building/floor/unit patterns
    patterns = [
        r"제?\d+동\s*제?\d+층\s*제?\d+호",
        r"제?\d+동\s*\d+층\s*\d+호",
        r"\d+동\s*\d+층\d+호",
        r"\d+층\d+동\d+호",
        r"제?\d+동",
        r"제?\d+층",
        r"제?\d+호",
        r"지\d+층", 
        r"비\d+호",
        r"외\s*\d+\s*개\s*호.*"
    ]
    
    cleaned = name
    for p in patterns:
        cleaned = re.sub(p, "", cleaned)
    
    cleaned = re.sub(r"\s+\d+호?$", "", cleaned)
    return cleaned.strip().strip(",")

def extract_gu_dong_beonji(addr):
    addr = safe_text(addr)
    if not addr:
        return ""
    dong = ""
    if "(" in addr:
        m = re.search(r"\(([^)]+)\)", addr)
        if m:
            for p in m.group(1).split(","):
                p_s = p.strip()
                if "동" in p_s:
                    dm = re.search(r"([가-힣\d]+동([가-힣\d]+가)?)", p_s)
                    if dm:
                        dong = dm.group(1)
                        break
    gu_m = re.search(r"([가-힣]+구)", addr)
    gu = gu_m.group(1) if gu_m else ""
    if not dong:
        dm = re.search(r"([가-힣\d]+동([가-힣\d]+가)?)", addr)
        dong = dm.group(1) if dm else ""
    beonji = ""
    if dong:
        bm = re.search(re.escape(dong) + r"\s+(\d+[-]?\d*)", addr)
        if bm: beonji = bm.group(1)
        else:
            parts = addr.split(dong)
            if len(parts) > 1:
                nm = re.search(r"(\d+[-]?\d*)", parts[1])
                if nm: beonji = nm.group(1)
    return f"{gu} {dong} {beonji}".strip()

def score_danzi_match(row, clean_search_name, norm_search_addr):
    score = 0
    danzi_name = compact_text(row.get('단지명', ''))
    legal_addr = compact_text(row.get('법정동주소', ''))

    if clean_search_name and danzi_name:
        if clean_search_name == danzi_name:
            score += 6
        elif clean_search_name in danzi_name or danzi_name in clean_search_name:
            score += 3

    if norm_search_addr and legal_addr:
        compact_addr = compact_text(norm_search_addr)
        if compact_addr == legal_addr:
            score += 6
        elif compact_addr in legal_addr or legal_addr in compact_addr:
            score += 4

    return score

def get_household_count_excel(apt_name, addr):
    """
    로컬에 저장된 K-apt 엑셀 파일(20260417_danzi_baseinfo.xlsx)에서 아파트 세대수 추출
    """
    global df_danzi
    
    try:
        # 최초 1회만 엑셀 로딩 (성능 향상)
        if df_danzi is None:
            # 첫번째 행(0번 인덱스)은 공지사항이므로 건너뛰고 1번 인덱스를 헤더로 파싱
            df_danzi = pd.read_excel("20260417_danzi_baseinfo.xlsx", header=1)
            # 결측치 빈 문자열로 처리
            df_danzi['단지명'] = df_danzi['단지명'].fillna('')
            df_danzi['법정동주소'] = df_danzi['법정동주소'].fillna('')
            # 세대수 숫자로 강제 변환
            df_danzi['세대수'] = pd.to_numeric(df_danzi['세대수'], errors='coerce').fillna(0).astype(int)
    except Exception as e:
        print("엑셀 데이터 로딩 실패:", e)
        df_danzi = pd.DataFrame()
        return None

    if df_danzi.empty:
        return None

    # 주소 정규화 (구 동 번지)
    norm_search_addr = extract_gu_dong_beonji(addr)
    clean_search_name = compact_text(clean_apt_name(apt_name))

    matched = pd.DataFrame()
    
    # 1. 주소 기반 검색 시도 (가장 정확)
    if norm_search_addr:
        simple_search_addr = compact_text(norm_search_addr)
        matched = df_danzi[
            df_danzi['법정동주소'].apply(
                lambda x: simple_search_addr in compact_text(x) if x else False
            )
        ]
        
    # 2. 주소 매칭 실패 시 이름으로 시도
    if matched.empty and clean_search_name:
        matched = df_danzi[
            df_danzi['단지명'].apply(
                lambda x: (
                    compact_text(x) in clean_search_name or clean_search_name in compact_text(x)
                ) if x else False
            )
        ]

    if len(matched) > 1:
        matched = matched.copy()
        matched['match_score'] = matched.apply(
            lambda row: score_danzi_match(row, clean_search_name, norm_search_addr),
            axis=1,
        )
        matched = matched.sort_values(['match_score', '세대수'], ascending=[False, False])

    if not matched.empty:
        return int(matched.iloc[0]['세대수'])
        
    return None

def generate_html_report(df, filename="madangs_report.html"):
    today = datetime.now().strftime("%Y년 %m월 %d일")
    # HTML/CSS 기반의 모던 리포트 템플릿 (Glassmorphism 적용)
    html = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>오늘의 우량 경매 리포트</title>
        <link href="https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;600;800&display=swap" rel="stylesheet">
        <style>
            :root {{
                --primary: #4F46E5;
                --primary-light: #818CF8;
                --bg: #F3F4F6;
                --card: rgba(255, 255, 255, 0.7);
                --text: #1F2937;
                --text-light: #6B7280;
            }}
            body {{
                font-family: 'Pretendard', -apple-system, sans-serif;
                background: linear-gradient(135deg, #E0E7FF 0%, #F5F3FF 100%);
                color: var(--text);
                margin: 0;
                padding: 40px 20px;
                min-height: 100vh;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
            }}
            header {{
                text-align: center;
                margin-bottom: 40px;
            }}
            h1 {{
                font-size: 2.5rem;
                font-weight: 800;
                background: linear-gradient(to right, var(--primary), #9333EA);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                margin: 0 0 10px 0;
            }}
            .subtitle {{
                font-size: 1.1rem;
                color: var(--text-light);
            }}
            .grid {{
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
                gap: 24px;
            }}
            .card {{
                background: var(--card);
                backdrop-filter: blur(10px);
                border: 1px solid rgba(255,255,255,0.4);
                border-radius: 20px;
                padding: 24px;
                box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);
                transition: transform 0.2s, box-shadow 0.2s;
            }}
            .card:hover {{
                transform: translateY(-5px);
                box-shadow: 0 20px 30px -10px rgba(0, 0, 0, 0.1);
            }}
            .card-header {{
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                margin-bottom: 16px;
            }}
            .badge {{
                background: var(--primary-light);
                color: white;
                padding: 4px 12px;
                border-radius: 20px;
                font-size: 0.8rem;
                font-weight: 600;
            }}
            .badge-dday {{
                background: #EF4444;
            }}
            .case-num {{
                font-weight: 800;
                font-size: 1.2rem;
                color: var(--primary);
                margin: 0 0 4px 0;
            }}
            .court {{
                font-size: 0.9rem;
                color: var(--text-light);
                margin: 0;
            }}
            h2 {{
                font-size: 1.2rem;
                margin: 0 0 8px 0;
                line-height: 1.4;
                color: #111827;
            }}
            .info-row {{
                display: flex;
                justify-content: space-between;
                padding: 10px 0;
                border-bottom: 1px dashed rgba(0,0,0,0.1);
                font-size: 0.95rem;
            }}
            .info-row:last-child {{
                border-bottom: none;
            }}
            .label {{
                color: var(--text-light);
                font-weight: 600;
            }}
            .value {{
                font-weight: 600;
                text-align: right;
            }}
            .value.price {{
                color: #B91C1C;
                font-size: 1.1rem;
            }}
            .footer {{
                text-align: center;
                margin-top: 50px;
                color: var(--text-light);
                font-size: 0.9rem;
            }}
            
            /* 필터 스타일 */
            .filter-container {{
                margin: 30px 0;
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 12px;
            }}
            .filter-label {{
                font-size: 0.9rem;
                font-weight: 600;
                color: var(--text-light);
            }}
            .filter-bar {{
                background: rgba(255, 255, 255, 0.5);
                backdrop-filter: blur(5px);
                padding: 6px;
                border-radius: 50px;
                display: flex;
                gap: 5px;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            }}
            .filter-btn {{
                border: none;
                background: transparent;
                padding: 8px 18px;
                border-radius: 40px;
                font-size: 0.85rem;
                font-weight: 600;
                color: var(--text-light);
                cursor: pointer;
                transition: all 0.2s;
            }}
            .filter-btn:hover {{
                background: rgba(255, 255, 255, 0.8);
            }}
            .filter-btn.active {{
                background: var(--primary);
                color: white;
                box-shadow: 0 4px 12px rgba(79, 70, 229, 0.3);
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <h1>🔔 우량 경매 추천 리포트</h1>
                <div class="subtitle">{today} 기준 (200세대 이상 추천)</div>
                
                <!-- 필터 영역 -->
                <div class="filter-container">
                    <div class="filter-label">세대수 필터링</div>
                    <div class="filter-bar" id="householdFilter">
                        <button class="filter-btn active" data-min="0">전체</button>
                        <button class="filter-btn" data-min="200">200세대+</button>
                        <button class="filter-btn" data-min="500">500세대+</button>
                        <button class="filter-btn" data-min="1000">1000세대+</button>
                        <button class="filter-btn" data-min="2000">2000세대+</button>
                    </div>
                </div>
            </header>
            <div class="grid" id="auctionGrid">
    """
    
    for _, row in df.iterrows():
        apt_info = row['apt_name'] if pd.notna(row['apt_name']) and row['apt_name'] else '아파트명 미상'
        addr = row['addr'] if pd.notna(row['addr']) else '-'
        special = row['special_right'] if pd.notna(row['special_right']) and str(row['special_right']).strip() else ''
        special_html = f'<div class="badge" style="background:#F59E0B; margin-top:12px;">⚠️ {special}</div>' if special else ''
        
        household = str(int(row['household'])) + '세대' if pd.notna(row['household']) else '정보없음'
        area = row['area'] if pd.notna(row['area']) else '-'
        
        eval_p = row['eval_price']
        low_p = row['low_price']
        eval_per = row['eval_per'] if pd.notna(row['eval_per']) else ''
        dday = row['dday'] if pd.notna(row['dday']) else '-'
        vcount = row['view_count'] if pd.notna(row['view_count']) else '-'
        
        html += f"""
                <div class="card" data-household="{row['household'] if pd.notna(row['household']) else 0}">
                    <div class="card-header">
                        <div>
                            <p class="court">{row.get('court', '')}</p>
                            <h3 class="case-num">{row.get('case_num', '')}</h3>
                        </div>
                        <div class="badge badge-dday">{dday}</div>
                    </div>
                    <h2>{apt_info}</h2>
                    <p style="font-size:0.9rem; color:#6B7280; margin-top:0; margin-bottom:16px;">{addr}</p>
                    
                    <div class="info-row">
                        <span class="label">단지 규모</span>
                        <span class="value">{household}</span>
                    </div>
                    <div class="info-row">
                        <span class="label">면적</span>
                        <span class="value">{area}</span>
                    </div>
                    <div class="info-row">
                        <span class="label">조회수</span>
                        <span class="value">🔥 {vcount}</span>
                    </div>
                    <div class="info-row">
                        <span class="label">감정가</span>
                        <span class="value" style="color:#6B7280; text-decoration:line-through;">{eval_p}</span>
                    </div>
                    <div class="info-row">
                        <span class="label">최저 입찰가</span>
                        <span class="value price">{low_p} <span style="font-size:0.85rem; color:#EF4444;">{eval_per}</span></span>
                    </div>
                    {special_html}
                </div>
        """
        
    html += """
            </div>
            <div class="footer">
                Automated by Antigravity Python Scraper &middot; Data from Madangs &amp; K-Apt
            </div>
        </div>

        <script>
            document.addEventListener('DOMContentLoaded', () => {
                const buttons = document.querySelectorAll('.filter-btn');
                const cards = document.querySelectorAll('.card');
                
                buttons.forEach(btn => {
                    btn.addEventListener('click', () => {
                        // 액티브 클래스 교체
                        buttons.forEach(b => b.classList.remove('active'));
                        btn.classList.add('active');
                        
                        const minHousehold = parseInt(btn.dataset.min);
                        
                        cards.forEach(card => {
                            const household = parseInt(card.dataset.household);
                            if (household >= minHousehold) {
                                card.style.display = 'block';
                            } else {
                                card.style.display = 'none';
                            }
                        });
                    });
                });
            });
        </script>
    </body>
    </html>
    """
    
    with open(filename, "w", encoding="utf-8-sig") as f:
        f.write(html)
    print(f"[OK] HTML 리포트 생성 완료 -> {filename}")

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json, text/plain, */*",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://madangs.com/search",
}

# -------- 유틸 --------
def flatten(obj, prefix="", out=None):
    if out is None: out = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            key = f"{prefix}.{k}" if prefix else str(k)
            flatten(v, key, out)
    else:
        out[(prefix or "").lower()] = obj
    return out

def pick(flat: dict, *cands):
    for key in cands:
        k = key.lower()
        if k in flat and flat[k] not in (None, ""):
            return flat[k]
        for fk in flat.keys():
            if fk.endswith("." + k) and flat[fk] not in (None, ""):
                return flat[fk]
    return None

def norm_date(s):
    if not isinstance(s, str): return s
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except: pass
    return s

CASE_NO_RE = re.compile(r"(\d{4}[-./]?\d{4,})")
def case_num_from_url(u):
    if not isinstance(u, str): return None
    m = CASE_NO_RE.search(u)
    return m.group(1) if m else None

APT_IN_PAREN_RE = re.compile(r"\(([^)]+)\)")
SPLIT_APT_RE = re.compile(r"(.+\d+[-]?\d*)\s+(.+)")

def apt_from_addr(addr):
    if not isinstance(addr, str): return None
    
    # 1. 괄호 안의 이름 추출 (예: 서울특별시 중구 중림동 355 (브라운스톤 서울))
    m = APT_IN_PAREN_RE.search(addr)
    if m:
        parts = [x.strip() for x in m.group(1).split(",") if x.strip()]
        if parts: return parts[-1]
    
    # 2. 콤마 뒤의 이름 추출 (예: 서울특별시 중구 중림동 355, 브라운스톤 서울)
    if "," in addr:
        return addr.split(",")[-1].strip()
    
    # 3. 주소 번지수 뒤에 공백으로 구분된 이름 추출 
    # (예: 서울특별시 중구 중림동 355 브라운스톤 서울)
    m = SPLIT_APT_RE.search(addr)
    if m:
        # Group 1: 주소부, Group 2: 아파트명부
        # 단, Group 2가 너무 짧거나('A동' 등) 하면 아파트명으로 보기 어려울 수 있으나 일단 반환
        apt_part = m.group(2).strip()
        if apt_part:
            return apt_part
            
    return None

def find_items_container(j):
    if isinstance(j, dict):
        for k in ("list", "rows", "data", "items", "result"):
            if isinstance(j.get(k), list) and j[k] and isinstance(j[k][0], dict):
                return j[k]
    # fallback
    cands, dq = [], deque([j])
    while dq:
        cur = dq.popleft()
        if isinstance(cur, dict):
            for v in cur.values(): dq.append(v)
        elif isinstance(cur, list):
            if cur and all(isinstance(x, dict) for x in cur): cands.append(cur)
            else:
                for v in cur: dq.append(v)
    return max(cands, key=len) if cands else []

# -------- 필요한 필드만 --------
def normalize_kor(item: dict) -> dict:
    f = flatten(item)

    court    = pick(f, "bubwon", "bubwon_short", "court", "court_name")
    case_num = pick(f, "case_num")
    case_url = pick(f, "case_url", "view_url", "detail_url")
    if not case_num:
        case_num = pick(f, "case_no", "case", "caseno") or case_num_from_url(case_url)

    addr     = pick(f, "addr", "address", "road_addr", "load_addr", "location")
    apt_name = clean_apt_name(pick(f, "apt_name", "name", "title") or apt_from_addr(addr))

    # 평형/면적(m2) 정보 추출
    area_m2 = pick(f, "areas.build.m")
    if not area_m2:
        area_m2 = pick(f, "build_area", "area_m2", "size")
        
    area_p = pick(f, "areas.build.p", "area_pyeong", "area_int")
    
    area_info = None
    if area_m2 and area_p:
        area_info = f"{area_m2}m² ({area_p}평)"
    elif area_m2:
        area_info = f"{area_m2}m²"
    elif area_p:
        area_info = f"{area_p}평"

    # ✅ 한글 표기 그대로
    eval_price = pick(f, "eval_price")       # 감정가
    low_price  = pick(f, "low_price_kor")    # 최저가(한글)
    last_price = pick(f, "last_price_kor")   # 직전가(한글)

    # 추가 인사이트 필드 반영
    use_type = pick(f, "use_type") # 물건 용도
    
    special_right = pick(f, "special_right") # 특수 권리
    if isinstance(special_right, list):
        special_right = ", ".join(map(str, special_right))
        
    view_count = pick(f, "view_count") # 조회수
    dday_num = pick(f, "dday_num") # 입찰일까지 남은 일수
    if dday_num is not None:
        dday_num = f"D-{dday_num}"
        
    eval_per = pick(f, "eval_per") or pick(f, "eval_per_value") # 하락률
    if str(eval_per).isdigit():
        eval_per = f"-{eval_per}%"

    bid_date   = norm_date(pick(f, "m_bid_date", "bid_date", "auction_date", "sale_date"))

    # 엑셀 파일(로컬)에서 세대수 연동
    household = get_household_count_excel(apt_name, addr)

    return {
        "court": court,
        "case_num": case_num,
        "use_type": use_type,
        "apt_name": apt_name,
        "household": household,
        "addr": addr,
        "area": area_info,
        "eval_price": eval_price,
        "low_price": low_price,
        "eval_per": eval_per,
        "last_price": last_price,
        "bid_date": bid_date,
        "dday": dday_num,
        "view_count": view_count,
        "special_right": special_right,
    }

def main():
    print("Fetching data...")
    r = requests.get(URL, headers=HEADERS, timeout=15)
    r.raise_for_status()
    j = r.json()

    items = find_items_container(j)
    rows = []
    skipped_count = 0
    for it in items:
        row = normalize_kor(it)
        rows.append(row)

    cols = ["court", "case_num", "use_type", "apt_name", "household", "addr", "area", "eval_price", "low_price", "eval_per", "last_price", "bid_date", "dday", "view_count", "special_right"]
    df = pd.DataFrame(rows, columns=cols)

    if "bid_date" in df.columns:
        df["bid_date_sort"] = df["bid_date"].fillna("9999-12-31")
        df = df.sort_values(["bid_date_sort"], ascending=[True]).drop(columns=["bid_date_sort"])

    df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    print(f"[OK] {len(df)}건 저장 -> {OUT_CSV}")
    
    # HTML 레포트 생성 
    generate_html_report(df)
    
    print(df.head(12).to_string(index=False))

if __name__ == "__main__":
    main()
