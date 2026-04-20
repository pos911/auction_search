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

def get_household_count_excel(apt_name, addr):
    """
    로컬에 저장된 K-apt 엑셀 파일(20260417_danzi_baseinfo.xlsx)에서 아파트 세대수 추출
    """
    global df_danzi
    if not apt_name:
        return None
        
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

    # 아파트 이름으로 필터링 (부분 일치 포함)
    matched = df_danzi[df_danzi['단지명'].str.contains(apt_name, regex=False, na=False)]
    
    # 매칭된 아파트가 여러 개일 경우 주소로 추가 교차 검증 (옵션)
    if len(matched) > 1 and addr:
        # 경매 주소 '서울특별시 은평구 증산동 15' 에서 동 정보(증산동 등) 추출해 비교
        # 엑셀의 법정동주소 '서울특별시 은평구 증산동 15' 와 매칭
        best_matches = []
        for idx, row in matched.iterrows():
            # 간단히 엑셀 법정동 주소의 핵심 키워드가 경매 주소에 포함되는지 확인
            # 예: 법정동주소 2번째 단어("은평구", "증산동" 등)
            addr_parts = str(row['법정동주소']).split()
            if len(addr_parts) >= 3 and addr_parts[2] in addr:
                best_matches.append(row)
        
        if best_matches:
            return int(best_matches[0]['세대수'])
            
    # 첫번째 매칭된 단지 세대수 반환
    if len(matched) > 0:
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
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <h1>🔔 우량 경매 추천 리포트</h1>
                <div class="subtitle">{today} 기준 (200세대 이상 아파트 필터링)</div>
            </header>
            <div class="grid">
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
                <div class="card">
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
def apt_from_addr(addr):
    if not isinstance(addr, str): return None
    m = APT_IN_PAREN_RE.search(addr)
    if m:
        parts = [x.strip() for x in m.group(1).split(",") if x.strip()]
        if parts: return parts[-1]
    if "," in addr: return addr.split(",")[-1].strip()
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
    apt_name = apt_from_addr(addr) or pick(f, "apt_name", "name", "title")

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
        
        # 세대수 기반 필터링 로직
        if row["household"] is not None:
            if row["household"] < MIN_HOUSEHOLD:
                skipped_count += 1
                continue # 세대수가 MIN_HOUSEHOLD 미만이면 제외
                
        rows.append(row)

    if skipped_count > 0:
        print(f"[{skipped_count}건] {MIN_HOUSEHOLD}세대 미만으로 필터링되어 제외되었습니다.")

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
