# -*- coding: utf-8 -*-
"""
잡알리오(공공기관 채용정보시스템) 고졸 채용공고 주 1회 수집기

하는 일
  1. 잡알리오에서 '고졸 + 신입' 공고 목록을 가져온다
  2. 지난번에 본 적 없는 공고만 골라 상세 페이지를 읽는다
  3. 기계/전기/전자/용접/사무 계열만 남긴다
  4. 구글시트 '공고2026' 탭과 똑같은 24열 형식으로 저장한다 → 확인 후 복사해 붙여넣기
  5. 사람이 읽는 요약 리포트(HTML)를 만든다

주의
  - 이 스크립트는 구글시트를 직접 고치지 않는다. 시트가 원천이고, 여기서는 '후보'만 제안한다.
  - 리눅스(깃허브 서버)에서는 파이썬 기본 방식으로 받고, 윈도우에서 SSL이
    막히면 PowerShell 로 우회한다. fetch() 가 알아서 고른다.
"""
import csv
import datetime as dt
import html
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request

from bs4 import BeautifulSoup

# ─────────────────────────── 설정 ───────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
STATE_PATH = os.path.join(BASE, "data", "state_seen.json")
LOG_PATH = os.path.join(BASE, "data", "log.txt")

# 직렬을 가르는 낱말 표.
#   공고가 '기계'라고 딱 적어 주는 일은 오히려 드물다. 시설·차량·설비·배관처럼
#   공업 계열 학생이 지원할 수 있는 표현을 폭넓게 넣어 둔다.
#   낱말을 넣고 빼는 것만으로 검색 범위가 바뀐다.
TAG_KEYWORDS = {
    "기계": ["기계", "설비", "정비", "차량", "배관", "공조", "냉동", "보일러",
             "금형", "가공", "선반", "밀링", "플랜트", "동력", "기전", "장비",
             "유지보수", "전동차", "궤도"],
    "전기": ["전기", "전력", "배전", "송전", "수배전", "계전", "발전설비"],
    "전자": ["전자", "통신", "신호", "계측", "제어", "자동화", "반도체", "정보통신"],
    "용접": ["용접", "접합", "판금"],
    "시설": ["시설", "토목", "건축", "건설", "소방", "안전", "환경", "화공",
             "조경", "산업안전", "위험물"],
    "사무": ["사무", "경영", "회계", "행정", "총무", "인사"],
}

# ── 계열 판별 ──────────────────────────────────────────────────
# 예전에는 공고 글자에서 '기계','설비' 같은 낱말을 찾아 계열을 추측했다.
# 그러다 '무기계약직' 안의 '기계'에 걸려 병동 업무직이 기계직으로 잡히곤 했다.
# 이제는 오픈API가 주는 NCS 대분류를 그대로 쓴다. 추측이 아니라 사실이다.
NCS_TO_TAGS = {
    "기계":              ["기계"],      # 항공기 제작·정비도 여기 들어간다
    "전기.전자":         ["전기", "전자"],   # 반도체도 여기 들어간다
    "정보통신":          ["전자"],
    "건설":              ["시설"],
    "재료":              ["용접"],      # 금속재료·용접
    "환경.에너지.안전":  ["시설"],
    "운전.운송":         ["기계"],
}
NCS_CODES = "R600015,R600019,R600020,R600014,R600016,R600023,R600009"

# NCS 대분류에는 '항공'도 '반도체'도 없다. 각각 기계·전기전자 아래 중분류라
# 대분류만으로는 학과별로 갈라 보여줄 수가 없어서, 낱말로 세부 표시를 덧붙인다.
# (계열 판별이 아니라 '이름표'용이라 조금 헐거워도 문제가 없다)
EXTRA_TAGS = {
    "항공드론": ["항공", "드론", "무인기", "무인항공", "UAM", "비행", "관제"],
    "반도체":   ["반도체", "웨이퍼", "포토", "식각", "증착", "패키징", "클린룸", "팹"],
}

CORE_TAGS = {"기계", "전기", "전자", "용접", "시설", "항공드론", "반도체"}
CORE_LABEL = "기계 · 전기 · 전자 · 용접 · 시설 · 항공드론 · 반도체"
INCLUDE_SAMU = False
TARGET_TAGS = CORE_TAGS | ({"사무"} if INCLUDE_SAMU else set())

# 잡알리오 검색 조건
#   education=R7030 고졸 / eduType=single 해당 학력만 / career=R2010 신입, R2030 신입+경력
#   eduType=single 은 학력요건이 '고졸'인 공고만, multi 는 고졸을 포함하는 공고까지
#   (대졸과 함께 뽑는 자리도 고졸 지원이 가능하므로 같이 훑고 비고에 표시한다)
QUERIES = [
    ("고졸-신입", "education=R7030&eduType=single&career=R2010"),
    ("고졸-신입경력", "education=R7030&eduType=single&career=R2030"),
    ("고졸포함-신입", "education=R7030&eduType=multi&career=R2010"),
    ("고졸포함-신입경력", "education=R7030&eduType=multi&career=R2030"),
]
LIST_URL = "https://job.alio.go.kr/recruit.do?pageNo={page}&pageSet=50&{q}"
VIEW_URL = "https://job.alio.go.kr/recruitview.do?idx={idx}"
MAX_PAGES = 5

COLS = ["순번", "구분", "기관명", "채용유형", "직렬태그", "모집직렬_원문", "인원",
        "대상", "학교장추천", "추천마감", "접수시작", "접수마감", "상태",
        "필기일", "면접일", "최종발표", "전형절차", "필기과목", "가점·자격",
        "성적요건", "공고링크", "관련서류", "출처", "비고"]


# 결과물은 깃허브 Pages 가 그대로 서비스하는 docs/ 아래에 쌓는다.
#   docs/index.html            ← 늘 최신 리포트
#   docs/2026-08-27.html       ← 주차별 리포트(지우지 않는다)
#   docs/archive.html          ← 지난 주차 목록
#   docs/files/2026-08-27/...  ← 엑셀·CSV·붙여넣기용 텍스트
#   docs/files/수집이력_전체.xlsx
DOCS_DIR = os.path.join(BASE, "docs")            # 깃허브 Pages 가 보는 폴더
ARCHIVE_DIR = DOCS_DIR                           # 주차별 리포트도 같은 깊이에 둔다
                                                 # (index.html 과 상대경로를 맞추려고)
FILES_DIR = os.path.join(DOCS_DIR, "files")      # 엑셀·CSV 내려받기
DATA_DIR = os.path.join(BASE, "data")            # 실행 상태(사람이 볼 일 없음)
OUT_DIR = FILES_DIR

# 지금까지 수집한 것을 모두 모아두는 누적 파일
HIST_COLS = COLS + ["잡알리오idx", "제목"]
HIST_CSV = os.path.join(BASE, "data", "history.csv")
HIST_XLSX = os.path.join(FILES_DIR, "수집이력_전체.xlsx")
DATE_COLS = ("추천마감", "접수시작", "접수마감", "필기일", "면접일", "최종발표")


for _d in (DOCS_DIR, ARCHIVE_DIR, FILES_DIR, DATA_DIR):
    os.makedirs(_d, exist_ok=True)


def log(msg):
    line = "[%s] %s" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ─────────────────────────── 내려받기 ───────────────────────────
# 리눅스(깃허브 서버)에서는 파이썬 기본 방식이 잘 된다.
# 다만 학교 노트북은 인증서 체인 문제로 파이썬 SSL이 깨져서, 윈도우에서는
# PowerShell 로 우회한다. 둘 다 준비해 두고 되는 쪽을 앞으로 당겨 쓴다.
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
_FETCHERS = None


def _fetch_urllib(url, timeout):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", "ignore")


def _fetch_powershell(url, timeout):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
    tmp.close()
    ps = (
        "$ErrorActionPreference='Stop';"
        "$r=Invoke-WebRequest -Uri '%s' -UseBasicParsing -TimeoutSec %d -UserAgent '%s';"
        "[IO.File]::WriteAllText('%s',$r.Content,[Text.Encoding]::UTF8)"
        % (url.replace("'", "''"), timeout, UA,
           tmp.name.replace("\\", "\\\\").replace("'", "''"))
    )
    try:
        r = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps],
            capture_output=True, text=True, timeout=timeout + 30,
        )
        if r.returncode != 0:
            raise RuntimeError(re.sub(r"\s+", " ", r.stderr or "").strip()[:140])
        with open(tmp.name, encoding="utf-8", errors="ignore") as f:
            return f.read()
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


def fetch(url, timeout=25, tries=2):
    """되는 방식으로 받아 UTF-8 텍스트를 돌려준다.

    잡알리오는 짧은 간격으로 연달아 부르면 406을 돌려주므로 몇 초 쉬고
    다시 시도한다.
    """
    global _FETCHERS
    if _FETCHERS is None:
        _FETCHERS = [_fetch_urllib]
        if os.name == "nt":
            _FETCHERS.append(_fetch_powershell)

    last = ""
    for attempt in range(1, tries + 1):
        for fn in list(_FETCHERS):
            try:
                text = fn(url, timeout)
                if text:
                    if _FETCHERS[0] is not fn:      # 되는 방식을 앞으로
                        _FETCHERS.remove(fn)
                        _FETCHERS.insert(0, fn)
                    return text
            except Exception as e:                  # noqa: BLE001 - 어떤 실패든 다음 방식으로
                last = "%s: %s" % (type(e).__name__, str(e)[:120])
        if attempt < tries:
            time.sleep(3 * attempt)
    log("  내려받기 실패(%d회 시도) %s :: %s" % (tries, url, last))
    return None


# ─────────────────────────── 파싱 ───────────────────────────
def parse_list(html_text):
    """목록 페이지 → [{idx, 제목, 기관명, 근무지, 고용형태, 등록일, 마감일, 상태}]

    제목 링크의 <a>는 비어 있고 글자는 그 <td> 안에 있다. 앞에 빈 칸이
    하나 더 있어 열이 밀리므로, 링크가 들어 있는 td를 기준으로 센다.
    """
    soup = BeautifulSoup(html_text, "lxml")
    out = []
    for a in soup.select("a[href*=recruitview]"):
        m = re.search(r"idx=(\d+)", a["href"])
        if not m:
            continue
        td = a.find_parent("td")
        tr = a.find_parent("tr")
        if not td or not tr:
            continue
        cells = tr.find_all("td")
        try:
            base = cells.index(td)          # 제목 칸의 위치
        except ValueError:
            continue
        get = lambda off: (cells[base + off].get_text(" ", strip=True)
                           if 0 <= base + off < len(cells) else "")
        out.append({
            "idx": m.group(1),
            "제목": get(0),
            "기관명": get(1),
            "근무지": get(2),
            "고용형태": get(3),
            "등록일": get(4),
            "마감일": get(5).split()[0] if get(5) else "",   # '26.09.09 D-13' → '26.09.09'
            "상태": get(6),
        })
    return out


def _section(soup, heading):
    """h4 제목 다음부터 다음 h4 전까지의 텍스트."""
    for h in soup.find_all(["h4", "h3"]):
        if heading in h.get_text(strip=True):
            parts = []
            for sib in h.next_siblings:
                if getattr(sib, "name", None) in ("h4", "h3"):
                    break
                if hasattr(sib, "get_text"):
                    t = sib.get_text(" ", strip=True)
                    if t:
                        parts.append(t)
            return re.sub(r"\s{2,}", " ", " ".join(parts)).strip()
    return ""


def parse_detail(html_text):
    soup = BeautifulSoup(html_text, "lxml")
    info = {}

    # 첫 표의 th/td 짝
    tables = soup.find_all("table")
    if tables:
        for tr in tables[0].find_all("tr"):
            cells = tr.find_all(["th", "td"])
            i = 0
            while i + 1 < len(cells):
                if cells[i].name == "th":
                    info[cells[i].get_text(" ", strip=True)] = cells[i + 1].get_text(" ", strip=True)
                    i += 2
                else:
                    i += 1

    # 모집분야: 2번째 표부터 각 표의 첫 th
    fields = []
    for t in tables[1:]:
        th = t.find("th")
        if th:
            txt = th.get_text(" ", strip=True)
            if txt and txt not in ("구분", "공고문") and len(txt) < 40:
                fields.append(txt)
    info["_모집분야"] = ", ".join(dict.fromkeys(fields))

    h2 = soup.find_all("h2")
    info["_기관명"] = h2[1].get_text(strip=True) if len(h2) > 1 else ""

    link = ""
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.startswith("http") and "alio.go.kr" not in href and "publicjob.kr" not in href:
            if a.get_text(strip=True) == href:      # 표 안의 홈페이지 주소만
                link = href
                break
    info["_공고링크"] = link

    # 첨부파일(공고문·입사지원서)은 alio 세션이 있어야 내려받아지는 주소라
    # 밖에서 열면 실패한다. 그래서 넣지 않고 지원 링크 하나만 남긴다.

    info["_응시자격"] = _section(soup, "응시자격")
    info["_전형절차"] = _section(soup, "전형절차")
    info["_우대내용"] = _section(soup, "우대내용")
    return info


# ────────────────────────── 분류·변환 ──────────────────────────
def classify(detail):
    """직렬태그를 뽑는다.

    '기계'라고 적히지 않아도 시설·차량·설비·배관처럼 공업 계열 학생이
    지원할 수 있는 표현이 많아, 낱말 표로 훑는다.

    표준직무(NCS)는 '경영.회계.사무,기계,전기.전자'처럼 뭉뚱그려 적히는 일이
    많아 그대로 쓰면 죄다 사무로 잡힌다. 실제 모집분야(공고에 분야별 표로
    들어가는 값)가 있으면 그쪽을 먼저 쓰고, 없을 때만 NCS로 되돌아간다.
    """
    blob = detail.get("_모집분야", "").strip()
    if not blob:
        blob = " ".join([detail.get("근무분야", ""), detail.get("표준직무(NCS)", "")])

    # '무기계약직' 안에 '기계'가 들어 있어 기계직으로 잘못 잡힌다. 고용형태
    # 표현은 직렬과 무관하므로 먼저 지운다.
    for noise in ("무기계약직", "무기계약", "비정규직", "정규직", "계약직"):
        blob = blob.replace(noise, " ")

    tags = [tag for tag, words in TAG_KEYWORDS.items()
            if any(w in blob for w in words)]
    if not tags:
        tags.append("기타")
    return ";".join(dict.fromkeys(tags))


def guess_target(detail):
    """응시자격 문구로 재학생·졸업생 대상을 가린다."""
    t = " ".join([detail.get("_응시자격", ""), detail.get("_우대내용", "")])
    예정 = bool(re.search(r"졸업\s*예정", t))
    졸업 = bool(re.search(r"졸업자|기졸업|졸업한|최종\s*학력.{0,8}고졸", t))
    if 예정 and 졸업:
        return "공통"
    if 예정:
        return "재학생"
    if 졸업:
        return "졸업생"
    return "확인필요"


RECOMMEND_RE = re.compile(
    # '학교장 추천'뿐 아니라 '학교장이 추천 기준에 맞게 추천한 자'처럼
    # 조사가 끼는 표현이 흔해서 사이를 넉넉히 열어 둔다.
    r"[^.\n]{0,80}("
    r"학교장.{0,8}추천"
    r"|학교.{0,4}추천서"
    r"|고교.{0,4}추천"
    r"|추천\s*기준에\s*맞게"
    r"|학교장의?\s*추천"
    r")[^.\n]{0,120}"
)


def detect_recommendation(detail):
    """학교장 추천이 필요한 공고인지 보고, 근거 문장을 함께 돌려준다."""
    t = " ".join([detail.get("_응시자격", ""), detail.get("_우대내용", ""),
                  detail.get("_전형절차", "")])
    m = RECOMMEND_RE.search(t)
    if m:
        return "필요", re.sub(r"\s+", " ", m.group(0)).strip()
    return "미표기", ""


def to_date(s):
    """'26.08.13' 또는 '2026.08.13' → 'YYYY-MM-DD'"""
    if not s:
        return ""
    s = s.strip()
    m = re.match(r"^(\d{2}|\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$", s)
    if not m:
        return ""
    y = int(m.group(1))
    if y < 100:
        y += 2000
    try:
        return dt.date(y, int(m.group(2)), int(m.group(3))).isoformat()
    except ValueError:
        return ""


def split_period(s):
    """'26.08.13 ~ 26.08.28' → ('2026-08-13','2026-08-28')"""
    if not s:
        return "", ""
    parts = re.split(r"~", s)
    a = to_date(parts[0]) if parts else ""
    b = to_date(parts[1]) if len(parts) > 1 else ""
    return a, b


def build_row(item, detail, today):
    tags = classify(detail)
    start, end = split_period(detail.get("채용기간", ""))
    if not end:
        end = to_date(item.get("마감일", ""))

    가점 = detail.get("우대조건", "") or detail.get("_우대내용", "")
    학력 = detail.get("학력정보", "").strip()
    대상 = guess_target(detail)
    추천, 추천근거 = detect_recommendation(detail)

    비고 = ["🔵 잡알리오 신규 — 시트 반영 검토 필요"]
    if 추천 == "필요":
        비고.append("⚠️ 학교장 추천: " + 추천근거[:200])
    if 학력:
        비고.append("학력요건: %s%s" % (학력, "" if 학력 == "고졸" else " (고졸 지원 가능, 타 학력과 병행 선발)"))
    if detail.get("근무지"):
        비고.append("근무지: " + detail["근무지"])
    if detail.get("_응시자격"):
        비고.append("응시자격: " + detail["_응시자격"][:180])

    상세URL = VIEW_URL.format(idx=item["idx"])
    # 기관 채용 사이트가 없으면 잡알리오 상세라도 열리게 한다
    지원링크 = detail.get("_공고링크", "") or 상세URL

    return {
        "순번": "",
        "구분": "확인필요",
        "기관명": detail.get("_기관명") or item.get("기관명", ""),
        "채용유형": detail.get("고용형태", "") or item.get("고용형태", ""),
        "직렬태그": tags,
        "모집직렬_원문": detail.get("_모집분야", "") or detail.get("근무분야", ""),
        "인원": detail.get("채용인원", ""),
        "대상": 대상,
        "학교장추천": 추천,
        "추천마감": "",
        "접수시작": start,
        "접수마감": end,
        "상태": "",
        "필기일": "", "면접일": "", "최종발표": "",
        "전형절차": (detail.get("_전형절차", "") or "")[:600],
        "필기과목": "",
        "가점·자격": (가점 or "")[:600],
        "성적요건": "",
        "공고링크": 지원링크,
        "관련서류": 상세URL,
        "출처": "잡알리오(%s 수집)" % today,
        "비고": " / ".join(비고),
        "_제목": item.get("제목", ""),
        "_idx": item["idx"],
        "_상세URL": 상세URL,
        "_학력": 학력,
    }


# ─────────────────────────── 출력 ───────────────────────────
def write_xlsx(rows, path, cols=None, title="신규공고"):
    cols = cols or COLS
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    ws.append(cols)
    thin = Border(*[Side(style="thin", color="D1D5DB")] * 4)
    for i in range(1, len(cols) + 1):
        c = ws.cell(row=1, column=i)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="374151")
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = thin
    widths = [6, 12, 20, 20, 16, 26, 10, 8, 10, 12, 12, 12, 9, 12, 12, 12, 46, 40, 40, 30, 28, 26, 18, 46]
    widths += [12] * max(0, len(cols) - len(widths))
    for i, w in enumerate(widths[:len(cols)], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    LINK = Font(color="1D4ED8", underline="single", size=10)
    for r in rows:
        ws.append([r.get(c, "") for c in cols])
    for i, data in enumerate(rows):
        r = i + 2
        ws.cell(row=r, column=13).value = (
            '=IF($L{0}="","미정",IF($L{0}<TODAY(),"마감",'
            'IF($L{0}-TODAY()<=7,"임박","진행중")))'.format(r)
        )
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cols[c - 1] in ("추천마감", "접수시작", "접수마감", "필기일", "면접일", "최종발표"):
                cell.number_format = "yyyy-mm-dd"

        # 눌러서 바로 열리는 건 지원 링크 하나만 둔다
        if data.get("공고링크"):
            cell = ws.cell(row=r, column=21)
            cell.hyperlink = data["공고링크"]
            cell.font = LINK
        # 학교장 추천이 필요한 건은 눈에 띄게
        if data.get("학교장추천") == "필요":
            c = ws.cell(row=r, column=9)
            c.fill = PatternFill("solid", fgColor="FCA5A5")
            c.font = Font(bold=True, size=10)
        if data.get("대상") == "재학생":
            ws.cell(row=r, column=8).fill = PatternFill("solid", fgColor="FFF176")
        ws.row_dimensions[r].height = 58

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # 확인 없이 그대로 쓰는 일이 없도록 안내 시트를 뒤에 붙인다
    note = wb.create_sheet("⚠️먼저읽기")
    note.column_dimensions["A"].width = 104
    lines = [
        ("⚠️ 반드시 원문 공고를 직접 확인하세요", True),
        ("", False),
        ("이 표는 잡알리오(공공기관 채용정보시스템) 요약을 자동으로 긁어 정리한 것입니다.", False),
        ("", False),
        ("· '대상(재학생/졸업생)'과 '학교장추천'은 공고 문구로 추정한 값입니다.", False),
        ("· 학교장 추천 마감일, 제출 서류, 세부 자격요건은 잡알리오에 실려 있지 않습니다.", False),
        ("· 추천 마감이 원서접수보다 한 달 넘게 앞서는 경우가 있습니다.", False),
        ("· '고졸 지원 가능, 타 학력과 병행 선발'로 표시된 건은 대졸과 함께 뽑는 공고입니다.", False),
        ("", False),
        ("학생에게 안내하거나 추천 절차를 진행하기 전에, '공고링크' 열의 주소로", False),
        ("기관 채용 사이트의 원문 공고문을 반드시 열어 확인해 주세요.", False),
        ("", False),
        ("이 파일은 구글시트 '공고2026' 탭과 열 순서가 같습니다.", False),
        ("필요한 행만 골라 시트 맨 아래에 붙여넣고, 순번·구분을 채운 뒤 비고의 🔵를 지우세요.", False),
    ]
    for i, (text, bold) in enumerate(lines, start=1):
        c = note.cell(row=i, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if bold:
            c.font = Font(bold=True, size=13, color="B91C1C")

    return save_workbook(wb, path)


def unique_path(path):
    """이미 있는 파일은 건드리지 않고 옆에 '(사본N)'을 붙여 새로 만든다.

    지난 실행 결과를 덮어쓰지 않기 위한 것. 기록이 계속 쌓인다.
    """
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    for n in range(2, 100):
        alt = "%s(사본%d)%s" % (stem, n, ext)
        if not os.path.exists(alt):
            return alt
    return "%s(사본%d)%s" % (stem, 99, ext)


def save_workbook(wb, path):
    """엑셀에서 파일을 열어 둔 채로 수집이 돌아도 죽지 않게 한다.

    쓰기가 막히면 옆에 '(사본N)'을 붙여 저장하고 어디에 넣었는지 알려 준다.
    """
    try:
        wb.save(path)
        return path
    except PermissionError:
        stem, ext = os.path.splitext(path)
        for n in range(2, 12):
            alt = "%s(사본%d)%s" % (stem, n, ext)
            try:
                wb.save(alt)
                log("  ⚠ %s 이(가) 열려 있어 %s 로 저장했습니다. "
                    "엑셀을 닫고 다시 실행하면 원래 이름으로 저장됩니다."
                    % (os.path.basename(path), os.path.basename(alt)))
                return alt
            except PermissionError:
                continue
        log("  ⚠ %s 을(를) 저장하지 못했습니다. 엑셀에서 닫고 다시 실행해 주세요."
            % os.path.basename(path))
        return None


def save_text(path, text, encoding="utf-8"):
    """텍스트 파일도 같은 이유로 막힐 수 있다."""
    try:
        with open(path, "w", encoding=encoding, newline="") as f:
            f.write(text)
        return path
    except PermissionError:
        log("  ⚠ %s 이(가) 열려 있어 건너뜁니다." % os.path.basename(path))
        return None


def tsv_of(rows):
    """구글시트에 그대로 붙여넣을 수 있는 탭 구분 텍스트.

    셀 안에 줄바꿈이나 탭이 있으면 붙여넣을 때 칸이 어긋나므로 공백으로 바꾼다.
    """
    lines = ["\t".join(COLS)]
    for r in rows:
        vals = []
        for c in COLS:
            v = r.get(c, "")
            if isinstance(v, dt.date):
                v = v.isoformat()
            vals.append(re.sub(r"[\t\r\n]+", " ", str(v or "")).strip())
        lines.append("\t".join(vals))
    return "\n".join(lines)


def write_report(rows, path, today, scanned, failed=None, downloads=None, mode="weekly"):
    esc = lambda s: html.escape(str(s or ""))
    css = """body{font-family:'Malgun Gothic','맑은 고딕',sans-serif;margin:24px;color:#111;background:#fff}
h1{font-size:20px;margin:0 0 4px}.sub{color:#666;font-size:13px;margin-bottom:14px}
h3{font-size:15px;margin:24px 0 10px;padding-bottom:5px;border-bottom:2px solid #111}
.card{border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;margin-bottom:12px}
.card h2{font-size:15px;margin:0 0 6px}.tag{display:inline-block;background:#dbeafe;color:#1e40af;
border-radius:99px;padding:1px 9px;font-size:12px;margin-right:5px}
.d{color:#b91c1c;font-weight:700}table{border-collapse:collapse;font-size:13px;margin-top:8px}
td{padding:2px 10px 2px 0;vertical-align:top}td.k{color:#666;white-space:nowrap}
.none{color:#666;padding:20px;border:1px dashed #d1d5db;border-radius:10px;text-align:center}
a{color:#1d4ed8}
.warn{border:2px solid #b91c1c;background:#fef2f2;color:#7f1d1d;border-radius:10px;
padding:12px 16px;margin:0 0 16px;font-size:13px;line-height:1.65}
.warn b{font-size:14px}
.bar{margin:0 0 18px;display:flex;gap:8px;flex-wrap:wrap}
.bar button{font-family:inherit;font-size:13px;padding:8px 16px;border:1px solid #111;
background:#111;color:#fff;border-radius:8px;cursor:pointer}
.bar button.ghost{background:#fff;color:#111}
.bar .btnlink{font-size:13px;padding:8px 16px;border:1px solid #111;background:#fff;
color:#111;border-radius:8px;text-decoration:none;display:inline-block}
.bar .btnlink:hover{background:#f3f4f6}
.bar button:hover{opacity:.85}
#tsv{width:100%;height:150px;font-family:Consolas,monospace;font-size:11px;
border:1px solid #d1d5db;border-radius:8px;padding:8px;white-space:pre;overflow:auto}
.hint{color:#666;font-size:12px;margin:6px 0 0}
table.sum{width:100%;border-collapse:collapse;font-size:13px;margin-top:8px}
table.sum th{background:#f3f4f6;border:1px solid #d1d5db;padding:6px 8px;
text-align:left;white-space:nowrap;font-weight:700}
table.sum td{border:1px solid #e5e7eb;padding:6px 8px;vertical-align:middle}
table.sum tr.urgent{background:#fef2f2}
table.sum td a{text-decoration:none;font-weight:700}
table.sum td a:hover{text-decoration:underline}
@media print{
 body{margin:0;font-size:11pt}
 .bar,#paste,.noprint{display:none !important}
 .card{break-inside:avoid;page-break-inside:avoid;border:1px solid #999}
 .warn{border:2px solid #000;background:#fff;color:#000}
 h3{break-after:avoid;page-break-after:avoid}
 a{color:#000;text-decoration:none}
}"""
    core = [r for r in rows if set(r["직렬태그"].split(";")) & CORE_TAGS]
    samu = [r for r in rows if r not in core]

    if mode == "open":
        head = "지금 지원할 수 있는 공고"
        sub = ("%s 기준으로 <b>접수가 끝나지 않은 공고 %d건</b>입니다. "
               "마감이 가까운 순서로 놓았습니다.<br>"
               "매주 월요일 아침에 저절로 새로 고쳐집니다 — 이 주소만 기억해 두세요."
               % (today, len(rows)))
    else:
        head = "이번 주 새로 올라온 공고"
        sub = ("%s · 새로 올라온 고졸·신입 공고 %d건을 확인해 "
               "<b>%s %d건</b>, 사무 %d건을 찾았습니다."
               % (today, scanned, CORE_LABEL, len(core), len(samu)))
    parts = ["<meta charset='utf-8'><title>%s</title><style>%s</style>" % (head, css)]
    parts.append("<h1>%s</h1>" % head)
    parts.append("<div class='sub'>%s</div>" % sub)

    if mode == "open":
        # 학생·학부모가 보는 화면
        parts.append(
            "<div class='warn'><b>⚠️ 지원하기 전에 반드시 원문 공고를 확인하세요.</b><br>"
            "이 표는 공공기관 채용정보를 자동으로 모아 정리한 것입니다. "
            "<b>자격요건과 마감일이 실제 공고와 다를 수 있고, 공고가 중간에 바뀌거나 "
            "취소되기도 합니다.</b> 각 공고의 <b>‘지원’ 링크</b>를 눌러 기관 채용 사이트에서 "
            "직접 확인하세요.<br>"
            "<b>🔴 ‘학교장추천 필요’ 표시가 있는 공고는 혼자 지원할 수 없습니다.</b> "
            "학교의 추천 절차를 먼저 거쳐야 하니, 반드시 <b>담임선생님이나 취업지원부와 "
            "먼저 상담</b>하세요. 추천 마감이 원서접수보다 한 달 넘게 이른 경우가 있습니다.</div>")
    else:
        # 선생님이 보는 주간 보고 화면
        parts.append(
            "<div class='warn'><b>⚠️ 반드시 원문 공고를 직접 확인하세요.</b><br>"
            "이 표는 공공기관 채용정보 요약을 자동으로 정리한 것입니다. "
            "특히 <b>대상(재학생/졸업생)</b>과 <b>학교장추천</b>은 공고 문구로 <b>추정한 값</b>이며, "
            "<b>학교장 추천 마감일·제출 서류·세부 자격요건은 원자료에 실려 있지 않습니다.</b><br>"
            "학생에게 안내하거나 추천 절차를 진행하기 전에, 각 카드의 <b>‘지원’ 링크</b>로 "
            "기관 채용 사이트의 원문 공고문을 반드시 열어 확인해 주세요. "
            "추천 마감이 원서접수보다 한 달 넘게 앞서는 경우가 있습니다.</div>")

    bar = ["<div class='bar'>",
           "<button onclick='window.print()'>🖨 인쇄 / PDF로 저장</button>",
           "<button class='ghost' onclick='copyTsv()'>📋 구글시트용으로 복사</button>"]
    for label, href in (downloads or []):
        bar.append("<a class='btnlink' href='%s' download>⬇ %s</a>"
                   % (esc(href), esc(label)))
    bar.append("<a class='btnlink' href='archive.html'>📁 지난 주차 보기</a>")
    bar.append("</div>")
    parts.append("".join(bar))

    if not rows:
        parts.append("<div class='none'>이번 주 새로 올라온 관련 공고가 없습니다.</div>")

    def dday_of(r):
        """(표시할 글자, 급한가) — 접수마감이 없으면 빈 값."""
        if not r.get("접수마감"):
            return "", False
        try:
            left = (dt.date.fromisoformat(str(r["접수마감"])) - dt.date.today()).days
        except ValueError:
            return "", False
        if left < 0:
            return "마감", False
        return "D-%d" % left, left <= 7

    ordered = core + samu
    for i, r in enumerate(ordered, 1):
        r["_no"] = i
        r["_id"] = "c%d" % i

    def cards(group):
        for r in group:
            tags = "".join("<span class='tag'>%s</span>" % esc(t) for t in r["직렬태그"].split(";"))
            label, _ = dday_of(r)
            dday = "<span class='d'>%s</span>" % esc(label) if label else ""
            parts.append("<div class='card' id='%s'><h2>%d. %s — %s</h2>%s %s<table>" % (
                r["_id"], r["_no"], esc(r["기관명"]), esc(r.get("_제목", "")), tags, dday))
            if r.get("학교장추천") == "필요":
                parts.append("<tr><td class='k'>학교장추천</td>"
                             "<td class='d'>필요 — 추천 마감을 먼저 확인하세요</td></tr>")
            for k in ("대상", "채용유형", "모집직렬_원문", "인원",
                      "접수시작", "접수마감", "가점·자격", "비고"):
                if r.get(k):
                    parts.append("<tr><td class='k'>%s</td><td>%s</td></tr>"
                                 % (esc(k), esc(r[k])[:400]))
            if r.get("공고링크"):
                parts.append("<tr><td class='k'>지원</td><td><a href='%s'>%s</a></td></tr>"
                             % (esc(r["공고링크"]), esc(r["공고링크"])))
            parts.append("</table></div>")

    if ordered:
        parts.append("<h3>한눈에 보기</h3>"
                     "<p class='hint'>줄을 누르면 아래 상세 내용으로 바로 갑니다.</p>"
                     "<table class='sum'><tr><th>#</th><th>기관</th><th>계열</th>"
                     "<th>대상</th><th>학교장추천</th><th>접수마감</th><th>상태</th></tr>")
        for r in ordered:
            label, urgent = dday_of(r)
            rec = r.get("학교장추천") or ""
            parts.append(
                "<tr%s><td>%d</td>"
                "<td><a href='#%s'>%s</a></td>"
                "<td>%s</td><td>%s</td><td%s>%s</td><td>%s</td><td%s>%s</td></tr>"
                % (" class='urgent'" if urgent else "",
                   r["_no"], r["_id"], esc(r["기관명"]),
                   esc(r["직렬태그"].replace(";", " · ")),
                   esc(r.get("대상") or ""),
                   " class='d'" if rec == "필요" else "",
                   ("🔴 필요" if rec == "필요" else esc(rec)),
                   esc(str(r.get("접수마감") or "")),
                   " class='d'" if (urgent or label == "마감") else "",
                   esc(label)))
        parts.append("</table>")

    if core:
        parts.append("<h3>상세 — %s</h3>" % CORE_LABEL)
        cards(core)
    if samu:
        parts.append("<h3 style='margin-top:26px'>상세 — 사무 (참고)</h3>")
        cards(samu)

    if rows and mode != "open":
        parts.append(
            "<div id='paste'><h3>구글시트에 붙여넣기</h3>"
            "<p class='hint'>위의 <b>📋 구글시트용으로 복사</b>를 누른 뒤, 구글시트 "
            "<b>공고2026</b> 탭의 <b>맨 아래 빈 줄 A열</b>을 클릭하고 <b>Ctrl+V</b> 하세요. "
            "열이 자동으로 나뉩니다. 붙여넣은 뒤 <b>순번·구분</b>을 채우고 비고의 🔵를 지우면 됩니다. "
            "제목줄은 빼고 붙이려면 첫 줄을 지우고 복사하세요.</p>"
            "<textarea id='tsv' readonly>%s</textarea></div>" % esc(tsv_of(rows)))
        parts.append("""<script>
function copyTsv(){
 var t=document.getElementById('tsv');
 if(!t){alert('복사할 내용이 없습니다.');return;}
 var msg='복사했습니다.\\n구글시트 공고2026 탭 맨 아래 빈 줄을 클릭하고 Ctrl+V 하세요.';
 t.removeAttribute('readonly'); t.focus(); t.select();
 t.setSelectionRange(0, t.value.length);
 var ok=false;
 try{ ok=document.execCommand('copy'); }catch(e){ ok=false; }
 t.setAttribute('readonly','readonly');
 if(ok){ alert(msg); return; }
 if(navigator.clipboard){
   navigator.clipboard.writeText(t.value).then(function(){ alert(msg); },
     function(){ alert('복사하지 못했습니다. 아래 칸을 직접 선택해 Ctrl+C 하세요.'); });
   return;
 }
 alert('복사하지 못했습니다. 아래 칸을 직접 선택해 Ctrl+C 하세요.');
}
</script>""")

    if failed:
        links = " · ".join("<a href='%s'>%s</a>" % (VIEW_URL.format(idx=i), esc(i)) for i in failed)
        parts.append("<h3 style='margin-top:26px'>읽지 못한 공고</h3>"
                     "<div class='none'>잡알리오가 %d건을 돌려주지 않았습니다(406). "
                     "다음 주에 다시 시도합니다. 급하면 직접 열어 보세요 — %s</div>"
                     % (len(failed), links))

    save_text(path, "\n".join(parts))


def write_archive_index():
    """지난 주차 리포트를 날짜순으로 모아 보여주는 목록 페이지."""
    names = sorted((f for f in os.listdir(DOCS_DIR)
                    if re.match(r"^\d{4}-\d{2}-\d{2}.*\.html$", f)), reverse=True)
    items = "\n".join(
        "<li><a href='%s'>%s</a></li>" % (html.escape(n), html.escape(n[:-5]))
        for n in names)
    page = (
        "<meta charset='utf-8'><title>지난 주차 모아보기</title>"
        "<style>body{font-family:'Malgun Gothic',sans-serif;margin:28px;color:#111}"
        "h1{font-size:19px}ul{line-height:2;padding-left:18px}a{color:#1d4ed8}"
        ".back{display:inline-block;margin-bottom:14px}</style>"
        "<a class='back' href='index.html'>← 이번 주 보기</a>"
        "<h1>지난 주차 모아보기</h1><ul>%s</ul>" % (items or "<li>아직 없습니다.</li>"))
    save_text(os.path.join(DOCS_DIR, "archive.html"), page)


# ────────────────────────── 누적 이력 ──────────────────────────
def load_history():
    """지난 실행에서 모아 둔 공고를 전부 읽어 온다."""
    if not os.path.exists(HIST_CSV):
        return []
    try:
        with open(HIST_CSV, encoding="utf-8-sig", newline="") as f:
            return [dict(r) for r in csv.DictReader(f)]
    except (OSError, csv.Error) as e:
        log("  이력 파일을 읽지 못했습니다(%s). 이번 것부터 새로 씁니다." % e)
        return []


def save_history(hist):
    """CSV(원본)와 XLSX(보기용)를 함께 남긴다."""
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=HIST_COLS, extrasaction="ignore")
    w.writeheader()
    for h in hist:
        w.writerow({c: ("" if h.get(c) is None else h.get(c)) for c in HIST_COLS})
    save_text(HIST_CSV, buf.getvalue(), encoding="utf-8-sig")

    # 엑셀에서는 날짜가 날짜답게 보이도록 되돌린다
    view = []
    for h in hist:
        row = dict(h)
        for c in DATE_COLS:
            v = row.get(c)
            if isinstance(v, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", v):
                row[c] = dt.date.fromisoformat(v)
        view.append(row)
    view.sort(key=lambda r: (str(r.get("접수마감") or "0000"), str(r.get("기관명") or "")),
              reverse=True)
    write_xlsx(view, HIST_XLSX, cols=HIST_COLS, title="수집이력")


# ─────────────────────────── 본체 ───────────────────────────
# ───────────────────── 공공데이터포털 오픈API 확인 ─────────────────────
# 아직 수집에는 쓰지 않는다. 응답이 어떤 이름으로 오는지 알아내려고
# 한 번만 불러 구조를 로그에 남기는 단계.
#
#   https://apis.data.go.kr/1051000/recruitment/list
#   acbgCondLst 학력(고졸 R7030) / recrutSe 채용구분(신입 R2010)
#   ncsCdLst NCS분류(기계 R600015, 전기.전자 R600019)
#   replmprYn 대체인력여부 / resultType 응답형태
#
# ⚠ 로그 파일은 공개 저장소에 그대로 커밋된다. 인증키가 절대 로그에
#   들어가지 않도록, 키가 붙은 주소는 어떤 경우에도 찍지 않는다.
API_BASE = "https://apis.data.go.kr/1051000/recruitment"


def _api_key():
    """Encoding·Decoding 어느 형태로 넣어도 되게 맞춰 준다."""
    key = (os.environ.get("ALIO_API_KEY") or "").strip()
    if key and "%" in key:                 # Encoding 표기면 한 번 풀어 준다.
        key = urllib.parse.unquote(key)    # (그래야 아래서 다시 인코딩할 때 이중 인코딩이 안 된다)
    return key


def api_get(path, params):
    """오픈API 호출. 인증키가 붙은 주소는 절대 로그에 남기지 않는다."""
    key = _api_key()
    if not key:
        return None
    p = dict(params)
    p["serviceKey"] = key
    p.setdefault("resultType", "json")
    url = API_BASE + path + "?" + urllib.parse.urlencode(p)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=20) as r:
            return json.loads(r.read().decode("utf-8", "ignore"))
    except Exception as e:                  # noqa: BLE001 - 주소를 찍으면 키가 샌다
        log("  오픈API 호출 실패(%s): %s" % (path, type(e).__name__))
        return None


def api_fetch_all():
    """고졸·신입 공고를 오픈API로 모아 온다. 못 쓰면 None(→ 스크래핑으로)."""
    key = _api_key()
    if not key:
        log("오픈API 키 없음 — 스크래핑으로 진행합니다")
        return None

    recs, ok = {}, False
    for se, name in (("R2010", "신입"), ("R2030", "신입+경력")):
        page = 1
        while page <= 20:
            data = api_get("/list", {
                "acbgCondLst": "R7030",      # 고졸
                "recrutSe": se,
                "ncsCdLst": NCS_CODES,
                "replmprYn": "N",            # 대체인력 제외
                "numOfRows": 100,
                "pageNo": page,
            })
            if data is None:
                break
            if str(data.get("resultCode")) not in ("200", "0", "00"):
                log("  오픈API 오류 %s: %s"
                    % (data.get("resultCode"), str(data.get("resultMsg"))[:80]))
                break
            batch = data.get("result") or []
            for r in batch:
                recs[str(r.get("recrutPblntSn"))] = r
            total = int(data.get("totalCount") or 0)
            ok = True
            log("  API %s p%d → %d건 (전체 %d)" % (name, page, len(batch), total))
            if not batch or len(recs) >= total:
                break
            page += 1

    if not ok:
        log("오픈API를 쓰지 못했습니다 — 스크래핑으로 되돌립니다")
        return None
    return list(recs.values())


def api_tags(rec):
    """NCS 대분류로 계열을 정하고, 낱말로 세부 이름표를 덧붙인다."""
    tags = []
    for nm in (rec.get("ncsCdNmLst") or "").split(","):
        tags.extend(NCS_TO_TAGS.get(nm.strip(), []))
    blob = " ".join([rec.get("recrutPbancTtl") or "", rec.get("ncsCdNmLst") or "",
                     rec.get("aplyQlfcCn") or "", rec.get("prefCondCn") or ""])
    for tag, words in EXTRA_TAGS.items():
        if any(w in blob for w in words):
            tags.append(tag)
    return ";".join(dict.fromkeys(tags)) or "기타"


def _ymd(s):
    """'20260908' → date"""
    s = re.sub(r"[^0-9]", "", str(s or ""))
    if len(s) != 8:
        return ""
    try:
        return dt.date(int(s[:4]), int(s[4:6]), int(s[6:]))
    except ValueError:
        return ""


def api_row(rec, today):
    """오픈API 레코드 → 구글시트와 같은 24열 한 줄."""
    detail = {
        "_응시자격": rec.get("aplyQlfcCn") or "",
        "_우대내용": " ".join([rec.get("prefCn") or "", rec.get("prefCondCn") or ""]),
        "_전형절차": rec.get("scrnprcdrMthdExpln") or "",
    }
    대상 = guess_target(detail)
    추천, 근거 = detect_recommendation(detail)
    학력 = rec.get("acbgCondNmLst") or ""
    idx = str(rec.get("recrutPblntSn"))

    비고 = ["🔵 잡알리오 신규 — 시트 반영 검토 필요"]
    if 추천 == "필요":
        비고.append("⚠️ 학교장 추천: " + 근거[:200])
    if 학력:
        비고.append("학력요건: %s%s"
                    % (학력, "" if 학력.strip() == "고졸" else " (고졸 지원 가능, 타 학력과 병행 선발)"))
    if rec.get("workRgnNmLst"):
        비고.append("근무지: " + rec["workRgnNmLst"])
    if detail["_응시자격"]:
        비고.append("응시자격: " + re.sub(r"\s+", " ", detail["_응시자격"])[:180])

    상세 = VIEW_URL.format(idx=idx)
    return {
        "순번": "", "구분": "확인필요",
        "기관명": rec.get("instNm") or "",
        "채용유형": rec.get("hireTypeNmLst") or "",
        "직렬태그": api_tags(rec),
        "모집직렬_원문": rec.get("ncsCdNmLst") or "",
        "인원": rec.get("recrutNope") or "",
        "대상": 대상, "학교장추천": 추천, "추천마감": "",
        "접수시작": _ymd(rec.get("pbancBgngYmd")),
        "접수마감": _ymd(rec.get("pbancEndYmd")),
        "상태": "", "필기일": "", "면접일": "", "최종발표": "",
        "전형절차": re.sub(r"\s+", " ", detail["_전형절차"])[:600],
        "필기과목": "",
        "가점·자격": re.sub(r"\s+", " ", detail["_우대내용"]).strip()[:600],
        "성적요건": "",
        "공고링크": rec.get("srcUrl") or 상세,
        "관련서류": 상세,
        "출처": "공공데이터 오픈API(%s 수집)" % today,
        "비고": " / ".join(비고),
        "_제목": rec.get("recrutPbancTtl") or "",
        "_idx": idx, "_상세URL": 상세, "_학력": 학력,
    }


def write_open_index(hist, today):
    """누적 이력에서 아직 접수 중인 공고만 골라 첫 화면을 만든다.

    주간 리포트는 '이번 주에 새로 뜬 것'이라, 수요일에 들어온 학생에게는
    지난주에 뜬 진행중 공고가 안 보인다. 상설 링크로 쓰려면 이 화면이 맞다.
    """
    rows = []
    for h in hist:
        end = h.get("접수마감") or ""
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", str(end)) or str(end) < today:
            continue
        r = dict(h)
        for c in DATE_COLS:
            v = r.get(c)
            if isinstance(v, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", v):
                r[c] = dt.date.fromisoformat(v)
        r["_제목"] = h.get("제목") or ""
        r["_idx"] = h.get("잡알리오idx") or ""
        r["_상세URL"] = h.get("관련서류") or ""
        r["비고"] = re.sub(r"🔵[^/]*/\s*", "", r.get("비고") or "")   # 교사용 표시는 뺀다
        rows.append(r)

    rows.sort(key=lambda r: (str(r.get("접수마감") or "9999"), r.get("기관명") or ""))
    write_report(rows, os.path.join(DOCS_DIR, "index.html"), today,
                 len(rows), None, None, mode="open")
    log("첫 화면: 모집중 %d건 (docs/index.html)" % len(rows))


def finish(rows, dropped, failed, all_ids, seen_idx, today, run_dir, n_new):
    """두 수집 경로(오픈API·스크래핑)가 공유하는 마무리 — 저장·리포트·이력."""
    # 핵심 계열을 위로, 그 안에서 마감 임박순
    downloads = []
    rows.sort(key=lambda r: (0 if set(r["직렬태그"].split(";")) & CORE_TAGS else 1,
                             r["접수마감"] or "9999", r["기관명"]))
    n_core = sum(1 for r in rows if set(r["직렬태그"].split(";")) & CORE_TAGS)
    log("핵심계열 %d건 / 사무 %d건 / 제외 %d건" % (n_core, len(rows) - n_core, dropped))

    if rows:
        xlsx = unique_path(os.path.join(run_dir, "주간수집_%s.xlsx" % today))
        write_xlsx(rows, xlsx)
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(COLS)
        for r in rows:
            w.writerow([r.get(c, "") for c in COLS])
        csvp = unique_path(os.path.join(run_dir, "주간수집_%s.csv" % today))
        csvp = save_text(csvp, buf.getvalue(), encoding="utf-8-sig")
        txt = unique_path(os.path.join(run_dir, "구글시트_붙여넣기_%s.txt" % today))
        save_text(txt, tsv_of(rows))
        if xlsx:
            downloads.append(("엑셀 내려받기", "files/%s/%s" % (today, os.path.basename(xlsx))))
        if csvp:
            downloads.append(("CSV 내려받기", "files/%s/%s" % (today, os.path.basename(csvp))))
        log("저장: docs/files/%s/%s" % (today, os.path.basename(xlsx or "")))

    # 누적 이력 — 주간 파일과 별개로, 지금까지 모은 것이 한 파일에 계속 쌓인다
    hist = load_history()
    known = {h.get("잡알리오idx") for h in hist}
    added = 0
    for r in rows:
        if r["_idx"] in known:
            continue
        h = {c: r.get(c, "") for c in COLS}
        for c in DATE_COLS:                       # CSV에는 날짜를 글자로 적는다
            h[c] = h[c].isoformat() if isinstance(h[c], dt.date) else (h[c] or "")
        h["상태"] = ""                             # 수식은 이력에 넣지 않는다
        h["잡알리오idx"] = r["_idx"]
        h["제목"] = r.get("_제목", "")
        hist.append(h)
        added += 1
    if added:
        save_history(hist)
    log("누적 이력: %d건 추가 / 전체 %d건" % (added, len(hist)))

    # 그 주 신규는 날짜 파일로 남긴다(지우지 않는다).
    report = unique_path(os.path.join(DOCS_DIR, "%s.html" % today))
    write_report(rows, report, today, n_new, failed, downloads)
    write_archive_index()
    log("주간 리포트: docs/%s" % os.path.basename(report))

    # 첫 화면은 '지금 지원할 수 있는 공고' — 언제 들어와도 쓸모 있게.
    write_open_index(hist, today)

    # 실행 흔적을 남겨 저장소가 '활동 중'으로 유지되게 한다.
    # (깃허브는 60일간 활동이 없으면 예약 워크플로를 꺼 버린다)
    save_text(os.path.join(DATA_DIR, "last_run.txt"),
              "%s 실행 / 신규 %d건 / 누적 %d건\n"
              % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), len(rows), len(seen_idx)))

    if failed:
        log("읽지 못한 공고 %d건 — 다음 주에 다시 시도합니다: %s" % (len(failed), ", ".join(failed)))
    seen_idx.update(k for k in all_ids if k not in set(failed))
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump({"updated": today, "seen_idx": sorted(seen_idx)}, f, ensure_ascii=False, indent=1)
    log("완료 — 누적 %d건 기억" % len(seen_idx))

    # 직접 실행했을 때는 결과 리포트를 바로 띄워 준다 (스케줄러는 조용히 지나간다)
    if "--open" in sys.argv and os.name == "nt":
        try:
            os.startfile(report)
        except OSError as e:
            log("  리포트를 열지 못했습니다: %s" % e)
    return 0


def main():
    today = dt.date.today().isoformat()
    run_dir = os.path.join(FILES_DIR, today)
    os.makedirs(run_dir, exist_ok=True)
    log("=" * 60)
    log("수집 시작")

    seen = {}
    if os.path.exists(STATE_PATH):
        try:
            with open(STATE_PATH, encoding="utf-8") as f:
                seen = json.load(f)
        except (OSError, ValueError):
            log("  상태 파일을 읽지 못해 새로 시작합니다")
    seen_idx = set(seen.get("seen_idx", []))

    # ── 1순위: 공공데이터 오픈API ──────────────────────────────
    api_recs = api_fetch_all()
    if api_recs:
        by_id = {str(r.get("recrutPblntSn")): r for r in api_recs}
        log("API 목록 %d건 확인, 기존 %d건" % (len(by_id), len(seen_idx)))
        new_ids = [i for i in by_id if i not in seen_idx]
        log("새 공고 %d건" % len(new_ids))

        rows, dropped = [], 0
        for i in new_ids:
            row = api_row(by_id[i], today)
            if not (set(row["직렬태그"].split(";")) & TARGET_TAGS):
                dropped += 1
                continue
            rows.append(row)
            log("  + %s | %s | %s"
                % (row["기관명"], row["직렬태그"], row.get("_제목", "")[:40]))
        return finish(rows, dropped, [], set(by_id), seen_idx,
                      today, run_dir, len(new_ids))

    # ── 2순위: 잡알리오 화면 스크래핑 (API 가 안 될 때) ──────────
    found = {}
    for name, q in QUERIES:
        for page in range(1, MAX_PAGES + 1):
            page_html = fetch(LIST_URL.format(page=page, q=q))
            if not page_html:
                break
            items = parse_list(page_html)
            if not items:
                break
            strict = "single" if "single" in q else "multi"
            for it in items:
                prev = found.get(it["idx"])
                if prev is None:
                    it["_strict"] = strict
                    found[it["idx"]] = it
                elif strict == "single":
                    prev["_strict"] = "single"      # 고졸 전용 쪽이 더 정확하다
            log("  %s p%d → %d건" % (name, page, len(items)))
            if len(items) < 50:
                break

    log("목록 %d건 확인, 기존 %d건" % (len(found), len(seen_idx)))
    new_idx = [i for i in found if i not in seen_idx]
    log("새 공고 %d건" % len(new_idx))

    rows, dropped, failed = [], 0, []
    for idx in new_idx:
        page_html = fetch(VIEW_URL.format(idx=idx))
        if not page_html:
            failed.append(idx)      # 다음 주에 다시 시도하도록 '본 것'에 넣지 않는다
            continue
        detail = parse_detail(page_html)
        item = found[idx]

        # 육아휴직·병가 대체인력, 기간제·촉탁 같은 비정규직은 학생 취업처가 아니다
        if detail.get("대체인력여부", "").strip() == "예":
            dropped += 1
            continue
        고용형태 = (detail.get("고용형태", "") or item.get("고용형태", "")).strip()
        if "비정규직" in 고용형태:
            dropped += 1
            continue

        row = build_row(item, detail, today)
        tags = set(row["직렬태그"].split(";"))
        if not (tags & TARGET_TAGS):
            dropped += 1
            continue
        # '고졸 포함' 넓은 검색으로만 걸린 건은 기계·전기·전자·용접일 때만 담는다.
        # (대졸과 함께 뽑는 사무직까지 넣으면 목록이 사무로 뒤덮인다)
        if item.get("_strict") == "multi" and not (tags & CORE_TAGS):
            dropped += 1
            continue

        rows.append(row)
        log("  + %s | %s | %s" % (row["기관명"], row["직렬태그"], row.get("_제목", "")[:40]))

    return finish(rows, dropped, failed, set(found), seen_idx,
                  today, run_dir, len(new_idx))


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:                     # 스케줄러가 조용히 죽는 걸 막는다
        log("오류: %s: %s" % (type(e).__name__, e))
        sys.exit(1)
