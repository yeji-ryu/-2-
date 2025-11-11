# -*- coding: utf-8 -*-
"""
generate_reports_fake_kr_dynamic.py

- HWPX 템플릿에서 토큰 robust 추출
- 업종(30+), 기업규모(스타트업~대기업)별 '의미 있는' 수치/문장 생성
- 표/지표/현금흐름/밸류에이션이 서로 '대략 일관'되도록 계산
- 미지 토큰은 토큰명(영/한 키워드)로 의미 추론하여 합리적 값 생성
- 각 값에는 어떤 워터마크/면책도 넣지 않음 (모든 내용은 허구)


사용:
  python generate_reports_fake_kr_dynamic.py --template ./template.hwpx --out ./out --count 5

필요 패키지:
  pip install faker python-dateutil tqdm
"""

import re, os, json, time, argparse, unicodedata, random
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from glob import glob
from datetime import datetime
from faker import Faker

# ---------------- Basics ----------------
FAKER_LOCALE = "ko_KR"
fake = Faker(FAKER_LOCALE)
NOISE = r"(?:\s|<[^>]+?>|[\u200B\uFEFF\u00A0]|&nbsp;)*"
TOKEN_PLAIN_RE = re.compile(r"([\{\｛]{2}.*?[\}\｝]{2})", re.DOTALL)
TOKEN_FAST_RE  = re.compile(r"{{\s*([A-Za-z0-9\u3131-\u318E\uAC00-\uD7A3_ /()\-\.%:,'\"]+?)\s*}}")

def nfc(s:str)->str:
    return unicodedata.normalize("NFC", s)

def clean_token_text(s:str)->str:
    s = re.sub(r"<[^>]+?>","", s, flags=re.DOTALL)
    s = s.replace("\u200B","").replace("\uFEFF","").replace("\u00A0"," ")
    s = re.sub(r"\s+"," ", s)
    return nfc(s.strip())

def extract_tokens_from_xml_text(xml:str):
    out=set()
    for m in TOKEN_FAST_RE.finditer(xml):
        out.add(nfc(m.group(1).strip()))
    for m in TOKEN_PLAIN_RE.finditer(xml):
        inner=re.sub(r"^[\{\｛]{2}","",m.group(1))
        inner=re.sub(r"[\}\｝]{2}$","",inner)
        inner=clean_token_text(inner)
        if inner: out.add(inner)
    return sorted(out)

def extract_tokens_from_hwpx(path:Path):
    found=set()
    with ZipFile(path,"r") as zin:
        for name in zin.namelist():
            if not name.lower().endswith(".xml"): continue
            try:
                xml=zin.read(name).decode("utf-8",errors="ignore")
            except:
                continue
            found |= set(extract_tokens_from_xml_text(xml))
    return sorted(found)

def build_split_tolerant_pattern(token:str)->re.Pattern:
    open_pat  = r"(?:\{\{|\｛\｛)"
    close_pat = r"(?:\}\}|\｝\｝)"
    parts = [re.escape(ch) for ch in token]
    mid   = NOISE.join(parts)
    full  = rf"{open_pat}{NOISE}{mid}{NOISE}{close_pat}"
    return re.compile(full, flags=re.IGNORECASE|re.DOTALL)

def replace_tokens_in_hwpx(template_hwpx:Path, out_hwpx:Path, mapping:dict):
    compiled = {nfc(k): build_split_tolerant_pattern(nfc(k)) for k in mapping}
    with ZipFile(template_hwpx,"r") as zin, ZipFile(out_hwpx,"w",compression=ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data=zin.read(name)
            if name.lower().endswith(".xml"):
                try: s=data.decode("utf-8",errors="ignore")
                except: zout.writestr(name,data); continue
                for key,pat in compiled.items():
                    val = mapping[key]
                    val = (val.replace("&","&amp;")
                               .replace("<","&lt;")
                               .replace(">","&gt;")
                               .replace('"',"&quot;"))
                    s = pat.sub(val, s)
                zout.writestr(name, s.encode("utf-8"))
            else:
                zout.writestr(name, data)

# ---------------- Industry Profiles ----------------
# 각 업종마다 "제품/고객/성장률/마진/CapEx/규제" 범위를 정의
INDUSTRY_PROFILES = {
    # SaaS/IT
    "클라우드/SaaS": dict(products=["SaaS 구독","API 플랫폼","데이터 파이프라인"],
                        customers=["중견·대기업 IT","스타트업","공공기관"],
                        rev_cagr=(12,35), opm=(8,22), ebitda=(12,30), capex_intensity=(2,6),
                        reg="정보보호·개인정보", price_mult=(4.0, 12.0), ev_ebitda=(8,20)),
    "AI/ML": dict(products=["ML 플랫폼","MLOps","비전/LLM 모듈"], customers=["제조·리테일","금융","플랫폼"],
                  rev_cagr=(20,45), opm=(5,18), ebitda=(8,22), capex_intensity=(3,8),
                  reg="데이터윤리·저작권", price_mult=(5.0,15.0), ev_ebitda=(10,25)),
    "핀테크": dict(products=["결제 게이트웨이","대안신용","자산관리"],
                 customers=["중소상공인","개인","금융사"],
                 rev_cagr=(15,35), opm=(5,15), ebitda=(8,18), capex_intensity=(1,4),
                 reg="전자금융·KYC/AML", price_mult=(3.0,10.0), ev_ebitda=(8,20)),
    "사이버보안": dict(products=["EDR","IAM","클라우드 보안"],
                 customers=["대기업","공공","금융"],
                 rev_cagr=(12,28), opm=(8,20), ebitda=(12,24), capex_intensity=(2,5),
                 reg="정보보호 인증", price_mult=(4.0,12.0), ev_ebitda=(10,22)),
    # 제조/하드
    "반도체": dict(products=["ASIC/컨트롤러","공정장비 모듈"], customers=["파운드리","팹리스"],
               rev_cagr=(6,18), opm=(12,25), ebitda=(18,32), capex_intensity=(12,25),
               reg="수출규제/환경안전", price_mult=(2.0,6.0), ev_ebitda=(6,12)),
    "배터리": dict(products=["셀/모듈/팩","리사이클"], customers=["완성차","ESS"],
               rev_cagr=(12,30), opm=(6,16), ebitda=(12,22), capex_intensity=(15,30),
               reg="환경·안전", price_mult=(1.8,5.5), ev_ebitda=(6,12)),
    "디스플레이": dict(products=["OLED 모듈","검사장비"], customers=["패널사","세트사"],
               rev_cagr=(4,12), opm=(6,14), ebitda=(10,20), capex_intensity=(10,22),
               reg="산업안전", price_mult=(1.5,4.0), ev_ebitda=(5,10)),
    "로보틱스": dict(products=["물류 로봇","협동로봇","자율주행 키트"], customers=["물류센터","제조"],
               rev_cagr=(15,35), opm=(5,15), ebitda=(8,20), capex_intensity=(8,16),
               reg="산업안전/자율주행", price_mult=(2.5,8.0), ev_ebitda=(8,18)),
    "국방/드론": dict(products=["전술 레이다","암호통신","UAV 센서"], customers=["국방부","방산사"],
               rev_cagr=(5,14), opm=(6,14), ebitda=(9,18), capex_intensity=(6,12),
               reg="ITAR/수출규제", price_mult=(1.8,5.0), ev_ebitda=(6,12)),
    "위성/우주항공": dict(products=["소형위성","위성통신페이로드"], customers=["정부/민간 통신"],
               rev_cagr=(10,25), opm=(5,12), ebitda=(8,16), capex_intensity=(18,35),
               reg="우주안전/주파수", price_mult=(2.0,6.0), ev_ebitda=(8,16)),
    # 소비/서비스
    "커머스/리테일": dict(products=["D2C 플랫폼","풀필먼트"], customers=["소비자","셀러"],
               rev_cagr=(6,18), opm=(2,8), ebitda=(4,10), capex_intensity=(2,6),
               reg="전자상거래/공정거래", price_mult=(0.8,2.0), ev_ebitda=(5,10)),
    "여행/호스피탈리티": dict(products=["OTA","숙박/투어","공항 서비스"], customers=["여행자","B2B 여행사"],
               rev_cagr=(8,22), opm=(3,10), ebitda=(6,14), capex_intensity=(3,8),
               reg="관광진흥/항공안전", price_mult=(1.0,3.0), ev_ebitda=(6,12)),
    "미디어/콘텐츠": dict(products=["스트리밍","IP/게임","스튜디오"], customers=["소비자","플랫폼"],
               rev_cagr=(6,18), opm=(5,15), ebitda=(8,18), capex_intensity=(4,10),
               reg="저작권/심의", price_mult=(1.5,4.5), ev_ebitda=(6,14)),
    "게임": dict(products=["모바일게임","PC/콘솔","LiveOps"], customers=["글로벌 유저"],
               rev_cagr=(8,24), opm=(12,28), ebitda=(16,32), capex_intensity=(3,7),
               reg="등급/청불", price_mult=(2.0,6.0), ev_ebitda=(8,16)),
    # 바이오/헬스
    "바이오헬스": dict(products=["진단키트","CDMO"], customers=["병원","제약"],
               rev_cagr=(8,22), opm=(8,18), ebitda=(12,24), capex_intensity=(6,14),
               reg="식약처/임상", price_mult=(2.0,7.0), ev_ebitda=(7,14)),
    "제약": dict(products=["신약개발","제네릭","CMC"], customers=["병원/도매"],
               rev_cagr=(5,14), opm=(8,18), ebitda=(12,24), capex_intensity=(6,12),
               reg="임상/허가", price_mult=(1.8,5.5), ev_ebitda=(6,12)),
    "의료기기": dict(products=["영상/체외진단","웨어러블"], customers=["병원/클리닉"],
               rev_cagr=(7,18), opm=(10,22), ebitda=(14,28), capex_intensity=(6,12),
               reg="인허가/보험수가", price_mult=(2.0,6.0), ev_ebitda=(7,14)),
    "디지털헬스": dict(products=["원격의료","PHR","AI 판독"], customers=["병원","보험사"],
               rev_cagr=(10,26), opm=(6,15), ebitda=(10,22), capex_intensity=(3,7),
               reg="원격의료/데이터", price_mult=(2.0,7.0), ev_ebitda=(8,16)),
    # 에너지/환경/산업
    "신재생에너지": dict(products=["태양광 EPC","ESS","PPA"], customers=["발전사","대기업"],
               rev_cagr=(6,20), opm=(6,14), ebitda=(10,20), capex_intensity=(12,28),
               reg="환경·전력시장", price_mult=(1.6,4.0), ev_ebitda=(5,10)),
    "수소": dict(products=["연료전지","수소충전"], customers=["운수/발전"],
               rev_cagr=(10,28), opm=(4,12), ebitda=(8,16), capex_intensity=(15,30),
               reg="안전·환경", price_mult=(1.8,5.0), ev_ebitda=(6,12)),
    "CCUS": dict(products=["포집 모듈","저장 EPC"], customers=["발전/정유"],
               rev_cagr=(6,20), opm=(6,14), ebitda=(10,18), capex_intensity=(15,32),
               reg="환경", price_mult=(1.6,4.0), ev_ebitda=(6,12)),
    "스마트팩토리": dict(products=["MES","IIoT","검사 비전"], customers=["제조"],
               rev_cagr=(8,22), opm=(8,18), ebitda=(12,24), capex_intensity=(6,12),
               reg="산업안전/정보보호", price_mult=(1.8,5.0), ev_ebitda=(7,14)),
    "산업IoT": dict(products=["센서/게이트웨이","플랫폼"], customers=["제조/물류"],
               rev_cagr=(8,24), opm=(6,16), ebitda=(10,20), capex_intensity=(6,12),
               reg="주파수/전파", price_mult=(1.8,5.0), ev_ebitda=(7,14)),
    "스마트시티": dict(products=["교통/치안/환경 솔루션"], customers=["지자체/공공"],
               rev_cagr=(6,16), opm=(5,12), ebitda=(8,16), capex_intensity=(8,18),
               reg="공공조달/개보법", price_mult=(1.5,4.0), ev_ebitda=(6,12)),
    "건설/인프라": dict(products=["토목/플랜트 EPC"], customers=["공공/대기업"],
               rev_cagr=(3,10), opm=(3,8), ebitda=(6,12), capex_intensity=(8,16),
               reg="안전/환경영향", price_mult=(0.8,2.0), ev_ebitda=(4,9)),
    "자동차부품": dict(products=["전장/샤시/내장"], customers=["OEM/티어1"],
               rev_cagr=(4,12), opm=(4,10), ebitda=(7,14), capex_intensity=(8,16),
               reg="품질/환경", price_mult=(0.8,2.0), ev_ebitda=(5,10)),
    "화학": dict(products=["정밀화학","폴리머"], customers=["전자/소재/소비"],
               rev_cagr=(3,10), opm=(6,14), ebitda=(10,20), capex_intensity=(10,20),
               reg="환경/안전", price_mult=(0.8,2.0), ev_ebitda=(5,10)),
    "철강/소재": dict(products=["특수강/합금","소재"], customers=["자동차/조선"],
               rev_cagr=(2,8), opm=(3,9), ebitda=(6,12), capex_intensity=(10,18),
               reg="환경/무역규제", price_mult=(0.6,1.6), ev_ebitda=(4,8)),
    # 물류/농식품/교육/광고/부동산/통신/블록체인
    "물류/라스트마일": dict(products=["풀필먼트","배송로봇"], customers=["커머스","리테일"],
               rev_cagr=(8,22), opm=(3,9), ebitda=(6,12), capex_intensity=(6,12),
               reg="산업안전/운송", price_mult=(0.8,2.2), ev_ebitda=(5,10)),
    "푸드테크": dict(products=["대체식품","키친 플랫폼"], customers=["식품사/소비자"],
               rev_cagr=(6,18), opm=(3,10), ebitda=(6,12), capex_intensity=(4,10),
               reg="식품위생", price_mult=(0.8,2.0), ev_ebitda=(5,10)),
    "애그리테크": dict(products=["정밀농업","스마트팜"], customers=["농가/식품사"],
               rev_cagr=(6,18), opm=(4,12), ebitda=(8,16), capex_intensity=(6,12),
               reg="농업/환경", price_mult=(0.8,2.2), ev_ebitda=(5,10)),
    "에듀테크": dict(products=["러닝SaaS","콘텐츠"], customers=["학교/개인"],
               rev_cagr=(8,24), opm=(6,16), ebitda=(10,20), capex_intensity=(3,8),
               reg="개인정보/저작권", price_mult=(1.5,4.5), ev_ebitda=(6,12)),
    "애드테크": dict(products=["퍼포먼스광고","CDP"], customers=["광고주/에이전시"],
               rev_cagr=(6,18), opm=(8,18), ebitda=(12,22), capex_intensity=(2,6),
               reg="개인정보/쿠키", price_mult=(1.8,5.0), ev_ebitda=(7,14)),
    "프롭테크": dict(products=["부동산 데이터","거래/자산관리"], customers=["개인/기관"],
               rev_cagr=(4,14), opm=(6,16), ebitda=(10,20), capex_intensity=(3,8),
               reg="부동산/개보법", price_mult=(1.0,3.0), ev_ebitda=(5,10)),
    "통신/네트워크": dict(products=["5G/FTTx","NW 장비"], customers=["통신사/공공"],
               rev_cagr=(3,10), opm=(6,14), ebitda=(14,28), capex_intensity=(12,22),
               reg="전파/안전", price_mult=(0.8,2.0), ev_ebitda=(4,9)),
    "블록체인/웹3": dict(products=["지갑/결제","인증/NFT"], customers=["개인/플랫폼"],
               rev_cagr=(10,30), opm=(5,15), ebitda=(8,18), capex_intensity=(2,6),
               reg="특금법/가상자산", price_mult=(1.5,5.0), ev_ebitda=(7,14)),
}

COMPANY_SCALES = [
    ("스타트업","Seed",          (5, 20)),   # (연매출 억원 범위) 기준선
    ("스타트업","Series A",      (15, 60)),
    ("스타트업","Series B",      (40, 150)),
    ("중소기업","-",             (80, 400)),
    ("중견기업","-",             (300, 1500)),
    ("대기업","-",              (1500, 10000)),
    ("공기업","-",              (500, 5000)),
    ("글로벌기업","-",          (3000, 20000)),
]

# ---------------- Name/phrasing helpers ----------------
import random

# 우선 제공된 기업명 목록 (순서대로 사용)
PREDEFINED_COMPANIES = [
    "JS(통신사)", "샤코(클라우드)", "블츠(AI)", "트페(핀테크)", "에코(사이버보안)",
    "벡스(반도체)", "조나우(배터리)", "오로라(디스플레이)", "바스티(로보틱스)", "애니비(국방/드론)",
    "아우소(위성/우주항공)", "말파리테일(커머스/리테일)", "조이관광(여행)", "세라핀(미디어/콘텐츠)",
    "레나타게임즈(게임)", "세트헬스(바이오헬스)", "실코(제약)", "미유의료(의료기기)", "레오나(디지털헬스)",
    "레르에너지(신재생에너지)", "오공수소(수소)", "자르반스마트(스마트팩토리)", "오른건설(건설)", 
    "하이거(자동차부품)", "신지드(화학)", "제이스오(철강/소재)", "아이번푸드(푸드테크)", 
    "쉬엔교육(에듀테크)", "바드(애드테크)", "니코아(엔터)", "누누식품(식품)", 
    "자르반(AI보안)", "카르마(산업디자인)", "나르(금융)"
]

# 기본 랜덤 조합용 데이터
PRE = ["한빛","네오","메가","넥스","루멘","에코","시그마","엘리시온","브릿지","포커스",
       "라디안","솔라","휴먼","레오","하임","시너지","에이펙스","에이든","스파크","샤코","블츠","트페","에코","벡스","조나우","링크","오로라","바스티","애니비","아우소","말파","조이","세라핀","레나타","세트","오공","실코","미유","레오나","자르반","아이번","오른","니코아","나르","누누","바드","카르마"]
SUF = ["테크","랩스","로보틱스","바이오","에너지","시스템즈","솔루션즈","소프트","디바이스",
       "네트웍스","인베스트먼트","프라임","파트너스","클라우드","의료","스마트","팩토리","헬스","푸드","식품","교육","관광","리테일","게임즈","에너지"]
FORM = ["주식회사","㈜","(유)"]

# 인덱스 추적용
_company_index = 0

def company_name():
    global _company_index
    if _company_index < len(PREDEFINED_COMPANIES):
        # 제공된 목록에서 순서대로 반환
        raw_name = PREDEFINED_COMPANIES[_company_index]
        _company_index += 1
        # 괄호 앞부분만 추출
        clean_name = raw_name.split("(")[0]
        return clean_name
    else:
        # 이후부터는 랜덤 생성
        pre = random.choice(PRE)
        suf = random.choice(SUF)
        form = random.choice(FORM)
        return f"{form} {pre}{suf}"


def person(role=None):
    return f"{fake.name()}" + (f" ({role})" if role else "")

def address():
    return fake.address().replace("\n"," ")

def yr():
    return random.randint(2018, 2032)

def pct(a,b,dec=1):
    return round(random.uniform(a,b),dec)

def money_억(a,b,step=1,decimals=False):
    v = random.randrange(a, b+1, step)
    if decimals:
        return f"{v}.{random.randint(0,9)}억 원"
    return f"{v}억 원"

def choice_sent(cands): return random.choice(cands)

# ---------------- Company Context (industry-aware) ----------------
def gen_context():
    sector = random.choice(list(INDUSTRY_PROFILES.keys()))
    prof = INDUSTRY_PROFILES[sector]
    scale, stage, rev_base = random.choice(COMPANY_SCALES)
    name = company_name()
    # 재무 사이즈 기준선
    base_rev = random.randint(rev_base[0], rev_base[1])
    years = 5
    now_year = datetime.now().year
    series_years = [now_year - (years - i) for i in range(1, years+1)]  # e.g., 2020..2024
    low_cagr, high_cagr = prof["rev_cagr"]
    cagr = pct(low_cagr, high_cagr, 1)
    revs = [base_rev]
    for _ in range(1, years):
        revs.append(int(revs[-1] * (1 + cagr/100.0) + random.uniform(-0.03,0.03)*revs[-1]))
    # 마진/현금흐름
    opm_start, opm_end = prof["opm"]
    opm_s = pct(opm_start, min(opm_start+5, opm_end), 1)
    opm_e = pct(max(opm_start, opm_end-5), opm_end, 1)
    op_profits = [round(revs[i]* (opm_s + (opm_e-opm_s)*i/(years-1))/100,1) for i in range(years)]
    ebitda_m = pct(prof["ebitda"][0], prof["ebitda"][1],1)
    ebitdas = [round(revs[i]*ebitda_m/100,1) for i in range(years)]
    net_m = max(1.0, opm_s-1.5)
    net_profits = [round(revs[i]* (net_m/100),1) for i in range(years)]
    # BS 대략 일관
    assets = [int(revs[i]*random.uniform(1.2,1.8)) for i in range(years)]
    debt_ratio = [pct(60, 140, 0) if scale!="대기업" else pct(30,90,0) for _ in range(years)]
    liabs = [int(assets[i]* (debt_ratio[i]/100)) for i in range(years)]
    equity = [assets[i]-liabs[i] for i in range(years)]
    cur_ratio = [pct(120, 180, 0) if scale!="스타트업" else pct(100,150,0) for _ in range(years)]
    # 현금흐름: 영업CF ~ EBITDA의 40~80%, 투자CF ~ -CapEx, 재무CF 보정
    capex_int = pct(prof["capex_intensity"][0], prof["capex_intensity"][1], 0) # 매출 대비 %
    oper_cf = [int(ebitdas[i]*random.uniform(0.4,0.8)) for i in range(years)]
    invest_cf = [ -int(revs[i]* capex_int/100 * random.uniform(0.6,1.2)) for i in range(years)]
    fin_cf = [int((op_profits[i]-abs(invest_cf[i]))*random.uniform(0.3,0.9)) for i in range(years)]
    cash_end = []
    cash = int(base_rev*random.uniform(0.02,0.08))
    for i in range(years):
        cash = max(1, cash + oper_cf[i] + invest_cf[i] + fin_cf[i])
        cash_end.append(cash)

    # 경쟁사 3개 생성
    comps = []
    for _ in range(3):
        cname = company_name().replace("주식회사 ","").replace("㈜ ","").replace("(유) ","")
        c_rev = int(revs[-1] * random.uniform(0.4, 2.0))
        c_share = round(random.uniform(0.5, 20.0), 1)
        c_tech = random.choice(prof["products"])
        c_clients = random.choice(prof["customers"])
        c_risk = random.choice(["원자재 가격 변동","규제 강화","고객 집중","신규 진입자"])
        comps.append(dict(name=cname, rev=c_rev, share=c_share, tech=c_tech, clients=c_clients, risk=c_risk))

    # Valuation 상식적 범위
    ev_rev = round(random.uniform(INDUSTRY_PROFILES[sector]["price_mult"][0], INDUSTRY_PROFILES[sector]["price_mult"][1]),1)
    ev_ebitda = round(random.uniform(INDUSTRY_PROFILES[sector]["ev_ebitda"][0], INDUSTRY_PROFILES[sector]["ev_ebitda"][1]),1)
    ev = int(revs[-1]*ev_rev)
    eq = int(ev * random.uniform(0.85, 0.95))
    pps = f"{int(random.uniform(5000, 50000))}원"

    return dict(
        sector=sector, prof=prof, scale=scale, stage=stage, name=name,
        years=series_years, revs=revs, op=op_profits, net=net_profits, ebitda=ebitdas,
        assets=assets, liabs=liabs, equity=equity, debt_ratio=debt_ratio, cur_ratio=cur_ratio,
        oper_cf=oper_cf, invest_cf=invest_cf, fin_cf=fin_cf, cash=cash_end,
        cagr=cagr, opm_start=opm_s, opm_end=opm_e, ebitda_m=ebitda_m,
        comps=comps, ev=ev, eq=eq, ev_rev=ev_rev, ev_ebitda=ev_ebitda, pps=pps
    )

# ---------------- Phrase generators (industry-aware) ----------------
def phrase(obj, ctx):
    sec = ctx["sector"]; prof = ctx["prof"]
    if obj=="hq_desc":
        return choice_sent([
            "주요 교통망 인접, 물류 접근성 양호",
            "R&D·생산 일체형 캠퍼스, 인증 설비 구비",
            f"{sec} 특화 클러스터 내 위치, 주요 협력사와 인접"
        ])
    if obj=="purpose_strategic":
        return choice_sent([
            f"{sec} 핵심 역량 내재화 및 포트폴리오 보완",
            "중장기 신사업 밸류체인 진입",
            "국내외 전략 고객 선점 및 장기 계약 기반 구축"
        ])
    if obj=="purpose_financial":
        return choice_sent([
            "3~5년 내 IPO/M&A를 통한 회수",
            "현금흐름 균형 개선과 재무 레버리지 최적화",
            "성장 자본 투입 후 멀티플 리레이팅 기대"
        ])
    if obj=="risk_fin":
        return "부채비율, 단기차입 만기 집중, 환율·원자재 변동에 따른 이익 변동성"
    if obj=="risk_biz":
        return "주요 고객 집중, 수주 변동성, 신규 경쟁 심화"
    if obj=="risk_tech":
        return "핵심 인력 의존, 품질 인증·성능 개선 지연 리스크"
    if obj=="risk_op":
        return "공정 불안정, 납기 준수 리스크, 내부 통제 미흡"
    if obj=="sensitive":
        return "내부거래 비율·지분 매각 검토·비공개 계약 조항 등 민감 사안 존재"
    if obj=="market_pos":
        return choice_sent([
            f"{prof['products'][0]} 특화로 틈새시장 점유",
            "고객 맞춤형 솔루션으로 계약 유지율 우수",
            "기술 차별화 기반 가격 경쟁력 확보"
        ])
    return ""

# ---------------- Token mapping (rule-based + fallback) ----------------
def normalize_key(t:str)->str:
    return nfc(t).strip().lower().replace(" ","_")

def klist(nmin=2,nmax=5, item="회사"):
    if "회사" in item:
        return ", ".join(company_name().replace("주식회사 ","") for _ in range(random.randint(nmin,nmax)))
    return ", ".join(fake.word() for _ in range(random.randint(nmin,nmax)))

def money_from_rev(rev, ratio_lo=0.02, ratio_hi=0.12):
    return f"{int(rev*random.uniform(ratio_lo, ratio_hi))}억 원"

def gen_known_value(key:str, ctx:dict):
    years = ctx["years"]; revs=ctx["revs"]; op=ctx["op"]; net=ctx["net"]; ebitda=ctx["ebitda"]
    assets=ctx["assets"]; liabs=ctx["liabs"]; equity=ctx["equity"]; dr=ctx["debt_ratio"]; cr=ctx["cur_ratio"]
    comps=ctx["comps"]

    mapping = {
        # --- 회사/기본 ---
        "target_company": ctx["name"],
        "company_name": ctx["name"],
        "registration_no": f"{random.randint(100,999)}-{random.randint(10,99)}-{random.randint(10000,99999)}",
        "registration_number": f"{random.randint(100,999)}-{random.randint(10,99)}-{random.randint(10000,99999)}",
        "report_date": f"{datetime.now().year}. {datetime.now().month:02d}. {datetime.now().day:02d}",
        "hq_address": address(),
        "hq_description": phrase("hq_desc", ctx),
        "hq_desc": phrase("hq_desc", ctx),
        "business_sector": ctx["sector"],
        "business_products": ", ".join(ctx["prof"]["products"]),
        "shareholders": f"최대주주 및 특수관계인, 전략투자자, 임직원·소액주주로 구성",
        "major_shareholders": f"{person()} 외 {random.randint(1,3)}인",
        "major_share_ratio": f"{pct(25,60,1)}%",
        "affiliate_company": company_name(),
        "affiliate_ratio": f"{pct(10,45,1)}%",
        "minority_shareholders": "임직원·소액주주",
        "clients": klist(item="회사"),
        "products": ", ".join(ctx['prof']['products'][:min(3,len(ctx['prof']['products']))]),
        "financials": f"{years[-1]}년 매출 {revs[-1]}억 원, 영업이익 {op[-1]}억 원, 임직원 {random.randint(20,2000)}명",
        # --- 목적/배경/포인트 ---
        "strategic_objectives": phrase("purpose_strategic", ctx),
        "financial_objectives": phrase("purpose_financial", ctx),
        "policy_environment": f"{ctx['prof']['reg']} 관련 정책 변화 모니터링 필요",
        "tech_competitiveness": choice_sent([
            "핵심 기술 내재화 및 특허 포트폴리오 보유",
            "제품 신뢰성·효율 개선에 따른 TCO 강점",
            "고부가가치 제품 믹스 전환 진행"
        ]),
        "growth_potential": choice_sent([
            f"핵심 고객군 확대 및 {ctx['sector']} 수요 성장",
            f"해외 시장 진출 가속 및 파트너십 확대",
            "제품 라인업 확장과 단가 개선 여지"
        ]),
        "hr_advantage": choice_sent([
            "핵심 연구인력 다수 보유, 이직률 낮음",
            "산업 도메인 전문가 중심의 조직",
            "내부 교육·채용 파이프라인 확보"
        ]),
        "market_position": phrase("market_pos", ctx),
        # --- 리스크 요약 ---
        "risk_financial": phrase("risk_fin", ctx),
        "risk_business": phrase("risk_biz", ctx),
        "risk_tech": phrase("risk_tech", ctx),
        "risk_operational": phrase("risk_op", ctx),
        "sensitive_risks": phrase("sensitive", ctx),
        # --- 투자 구조/조건 ---
        "investment_structure": random.choice(["Equity","Convertible","Mezzanine","혼합(신주+구주)"]),
        "investment_amount": money_from_rev(revs[-1], 0.15, 0.6),
        "investment_preference_terms": "우선주 배당/청산우선권·보호조항 등 투자자 권리 확보",
        "drag_tag_rights": "동반매도·매수권 반영(주요 거래 시 의무 적용)",
        "governance_participation": "이사회 1석·중요안건 동의권",
        "additional_conditions": "기술·법무·세무 DD 완료 후 집행",
        # --- Exit ---
        "exit_ipo_domestic": random.choice(["KOSDAQ","KOSPI"])+" 상장 추진",
        "exit_mna_strategic": "전략적 투자자 대상 매각 가능성",
        "exit_secondary": "재무적 투자자 대상 블록딜",
        "exit_share_buyback": "현금흐름 호전 시 자사주 매입",
        "exit_risk": "시장 침체/규제 변수로 일정 지연 가능",
        # --- 타임라인 ---
        "deal_review_start": "내부 투자심의 승인 후 비공개 검토 시작",
        "deal_review_date": f"{datetime.now().year}년 {random.randint(1,12)}월",
        "due_diligence": "재무·세무·법무·기술 전 영역 실사",
        "due_diligence_date": f"{datetime.now().year}년 {random.randint(1,12)}월",
        "contract_negotiation": "밸류·우선주 구조·보호조항 협상",
        "contract_negotiation_date": f"{datetime.now().year+1}년 {random.randint(1,12)}월",
        "contract_signing": "SHA/SPA 체결",
        "contract_signing_date": f"{datetime.now().year+1}년 {random.randint(1,12)}월",
        "capital_injection_and_transfer": "납입 및 지분 이전",
        "capital_injection_date": f"{datetime.now().year+1}년 {random.randint(1,12)}월",
        "post_management_date": f"{datetime.now().year+1}년 {random.randint(1,12)}월 이후 정기 모니터링",
        # --- 회사 연혁/지배구조 ---
        "company_foundation": f"{random.randint(1995,2023)}년 설립 → 법인 전환 {random.randint(2000,2024)}년",
        "company_growth_stage": choice_sent(["창업기","성장기","확장기","성숙기"]),
        "company_milestones": "핵심 제품 상용화·주요 고객사 확보·해외 파트너십",
        "company_diversification": "주요 시장 외 인접 영역 확장",
        "company_current_status": "R&D·생산·영업 조직 균형적 운영",
        "board_composition": "이사회 5인(사내 2, 사외 3), 감사 1",
        "board_ceo": person("CEO"),
        "board_cfo": person("CFO"),
        "board_cto": person("CTO"),
        "board_others": person("사외이사"),
        "governance_structure": "오너 중심이나 외부 이사 견제 장치 보유",
        # --- BM/수익 ---
        "revenue_model": choice_sent(["구독+라이선스+서비스","제품판매+유지보수","프로젝트+소모품"]),
        "revenue_b2g": money_from_rev(revs[-1],0.1,0.6),
        "revenue_b2b": money_from_rev(revs[-1],0.2,0.7),
        "revenue_b2c": money_from_rev(revs[-1],0.0,0.4),
        "revenue_service": money_from_rev(revs[-1],0.05,0.3),
        "b2g_share": f"{pct(0,60,1)}%",
        "b2b_share": f"{pct(10,80,1)}%",
        "b2c_share": f"{pct(0,40,1)}%",
        "service_share": f"{pct(5,40,1)}%",
        "key_revenue_sources": ", ".join(ctx["prof"]["products"]),
        "business_strategy": choice_sent(["고마진 제품 믹스 확대","전략 고객 장기계약","해외 파트너 채널 구축"]),
        # --- 제품/서비스 ---
        "product_service_a": ctx["prof"]["products"][0],
        "product_service_b": ctx["prof"]["products"][1] if len(ctx["prof"]["products"])>1 else "신규 라인업",
        "product_service_c": ctx["prof"]["products"][2] if len(ctx["prof"]["products"])>2 else "서비스 패키지",
        "product_service_d": "유지보수/운영 서비스",
        # --- 시장/경쟁 ---
        "total_addressable_market": money_억( int(revs[-1]*20), int(revs[-1]*300), 10),
        "serviceable_available_market": money_억( int(revs[-1]*8), int(revs[-1]*150), 5),
        "serviceable_obtainable_market": money_억( int(revs[-1]*3), int(revs[-1]*60), 5),
        "target_customer_segments": ", ".join(ctx["prof"]["customers"]),
        "additional_market_info": phrase("market_pos", ctx),
        "competitor_a_name": comps[0]["name"], "competitor_b_name": comps[1]["name"], "competitor_c_name": comps[2]["name"],
        "competitor_a_profile": f"{ctx['sector']} 주력, 고객 {comps[0]['clients']}",
        "competitor_b_profile": f"{ctx['sector']} 확대 중, 고객 {comps[1]['clients']}",
        "competitor_c_profile": f"틈새 강점, 기술 {comps[2]['tech']}",
        "competitor_a_internal_note": "제품 라인 안정적이나 혁신 속도 제한",
        "competitor_b_eval_note": "수익성 저하·납기 이슈로 감점",
        "competitor_c_tech_note": "기술 우수하나 영업 네트워크 약함",
        "competitor_a_revenue": f"{comps[0]['rev']}억원",
        "competitor_b_revenue": f"{comps[1]['rev']}억원",
        "competitor_c_revenue": f"{comps[2]['rev']}억원",
        "competitor_a_share": f"{comps[0]['share']}%",
        "competitor_b_share": f"{comps[1]['share']}%",
        "competitor_c_share": f"{comps[2]['share']}%",
        "target_revenue": f"{revs[-1]}억원",
        "competitor_a_clients": comps[0]["clients"], "competitor_b_clients": comps[1]["clients"], "competitor_c_clients": comps[2]["clients"], "target_clients": random.choice(ctx["prof"]["customers"]),
        "competitor_a_tech": comps[0]["tech"], "competitor_b_tech": comps[1]["tech"], "competitor_c_tech": comps[2]["tech"], "target_tech": random.choice(ctx["prof"]["products"]),
        "competitor_a_risk": comps[0]["risk"], "competitor_b_risk": comps[1]["risk"], "competitor_c_risk": comps[2]["risk"], "target_risk": "원가/환율/규제",
        "additional_competitor_info": "신규 진입자·대체재 동향 주기적 점검",
        # --- 규제 ---
        "regulatory_environment_current": f"{ctx['prof']['reg']} 규제·인증 요건 존재",
        "policy_changes_future": "지원정책/규제완화·강화 불확실성",
        "regulatory_risk_summary": "인증·승인 지연 시 매출 인식 지체 가능",
        # --- P/L 표 ---
        "year_1": str(years[-3]), "year_2": str(years[-2]), "year_3": str(years[-1]),
        "revenue_year1": str(revs[-3]), "revenue_year2": str(revs[-2]), "revenue_year3": str(revs[-1]),
        "op_profit_year1": str(int(op[-3])), "op_profit_year2": str(int(op[-2])), "op_profit_year3": str(int(op[-1])),
        "net_profit_year1": str(int(net[-3])), "net_profit_year2": str(int(net[-2])), "net_profit_year3": str(int(net[-1])),
        "op_margin_year1": f"{round(op[-3]/revs[-3]*100,1)}", "op_margin_year2": f"{round(op[-2]/revs[-2]*100,1)}", "op_margin_year3": f"{round(op[-1]/revs[-1]*100,1)}",
        "is_revenue": f"{years[-3]}~{years[-1]} 매출 {revs[-3]}→{revs[-1]}억",
        "is_revenue_trend": f"CAGR {ctx['cagr']}%",
        "is_op_profit": f"영업이익 {int(op[-3])}→{int(op[-1])}억",
        "is_op_profit_reason": "제품 믹스 개선·원가 절감",
        "is_net_profit": f"순이익 {int(net[-3])}→{int(net[-1])}억",
        "is_net_profit_comment": "판관비 효율화·금융비용 변동",
        # --- BS 표 ---
        "balancesheet_year_1": str(years[-3]),
        "balancesheet_year_2": str(years[-2]),
        "balancesheet_year_3": str(years[-1]),
        "balancesheet_totalassets_1": str(assets[-3]),
        "balancesheet_totalassets_2": str(assets[-2]),
        "balancesheet_totalassets_3": str(assets[-1]),
        "balancesheet_totalliabilities_1": str(liabs[-3]),
        "balancesheet_totalliabilities_2": str(liabs[-2]),
        "balancesheet_totalliabilities_3": str(liabs[-1]),
        "balancesheet_equity_1": str(equity[-3]),
        "balancesheet_equity_2": str(equity[-2]),
        "balancesheet_equity_3": str(equity[-1]),
        "balancesheet_debtratio_1": str(int(dr[-3])),
        "balancesheet_debtratio_2": str(int(dr[-2])),
        "balancesheet_debtratio_3": str(int(dr[-1])),
        "balancesheet_currentratio_1": str(int(cr[-3])),
        "balancesheet_currentratio_2": str(int(cr[-2])),
        "balancesheet_currentratio_3": str(int(cr[-1])),
        "fin_performance_keypoint": "수익성 개선·운전자본 효율화 필요",
        # --- 지표 추이 ---
        "roe_start_year": str(years[-5]) if len(years)>=5 else str(years[0]),
        "roe_end_year": str(years[-1]),
        "roe_start_val": f"{pct(3,10,1)}",
        "roe_end_val": f"{pct(6,18,1)}",
        "roe_period": str(len(years)-1),
        "roe_change": f"{pct(1,6,1)}",
        "roa_start_year": str(years[-5]) if len(years)>=5 else str(years[0]),
        "roa_end_year": str(years[-1]),
        "roa_start_val": f"{pct(1,6,1)}",
        "roa_end_val": f"{pct(2,10,1)}",
        "roa_comment": "자산 효율성 점진 개선",
        "ebitda_start_year": str(years[-5]) if len(years)>=5 else str(years[0]),
        "ebitda_start_val": f"{pct(max(5,ctx['ebitda_m']-5), ctx['ebitda_m'],1)}",
        "ebitda_end_year": str(years[-1]),
        "ebitda_end_val": f"{pct(ctx['ebitda_m'], min(35, ctx['ebitda_m']+6),1)}",
        "ebitda_comment": "규모의 경제·원가 구조 개선",
        "debt_ratio_start_year": str(years[-3]),
        "debt_ratio_start": f"{int(dr[-3])}",
        "debt_ratio_end_year": str(years[-1]),
        "debt_ratio_end": f"{int(dr[-1])}",
        "debt_ratio_comment": "리파이낸싱·만기 분산 필요",
        "liquidity_start_year": str(years[-3]),
        "liquidity_start": f"{int(cr[-3])}",
        "liquidity_end_year": str(years[-1]),
        "liquidity_end": f"{int(cr[-1])}",
        "liquidity_comment": "운전자본 관리 강화",
        "debt_dep_start": f"{pct(15,40,1)}",
        "debt_dep_end": f"{pct(15,45,1)}",
        "debt_dep_comment": "차입 의존도 관리 필요",
        "sales_period": str(len(years)-1),
        "sales_cagr": f"{ctx['cagr']}",
        "op_profit_period": str(len(years)-1),
        "op_profit_cagr": f"{pct(max(3,ctx['cagr']-3), ctx['cagr']+3,1)}",
        "years": str(len(years)-1),
        "project_effect": "신규 프로젝트 반영 효과",
        "target_roe_year": str(yr()),
        "target_roe": f"{pct(10,20,1)}",
        "risk_scenario": "수주 지연·환율 급등",
        "low_case_roe": f"{pct(3,8,1)}",
        "target_irr": f"{pct(12,25,1)}",
        "actual_irr": f"{pct(6,18,1)}",
        "impact_comment": "원가·환율 민감도 영향",
        "special_quarter": f"{years[-1]}년 {random.randint(1,4)}분기",
        "special_sales": f"{random.randint(5,50)}",
        "oneoff_project": "일회성 대형 주문",
        # --- CF ---
        "oper_cf_period": str(len(years)-1),
        "oper_cf_trend": choice_sent(["플러스 유지","변동성 확대","완만한 증가"]),
        "oper_cf_start_year": str(years[-3]),
        "oper_cf_start_val": str(ctx["oper_cf"][-3]),
        "oper_cf_end_year": str(years[-1]),
        "oper_cf_end_val": str(ctx["oper_cf"][-1]),
        "oper_cf_event": "운전자본 변동",
        "oper_cf_event_size": f"{random.randint(3,30)}",
        "invest_cf_comment": "설비/개발 투자 확대",
        "invest_cf_outflow": f"{abs(ctx['invest_cf'][-1])}",
        "invest_cf_direction": "지속 집행",
        "invest_cf_year": str(years[-1]),
        "invest_cf_event": "CAPEX/인수",
        "invest_cf_amount": f"{abs(ctx['invest_cf'][-1])}",
        "fin_cf_equity": f"{random.randint(5,80)}",
        "fin_cf_equity_year": str(years[-2]),
        "fin_cf_short_borrow": f"{random.randint(5,80)}",
        "fin_cf_long_repay": f"{random.randint(5,80)}",
        "dividend_amount": f"{random.randint(0,50)}",
        "dividend_year": str(years[-1]),
        "dividend_ratio": f"{pct(0,60,1)}",
        "oper_cf_1": str(ctx["oper_cf"][-5]) if len(years)>=5 else "1",
        "oper_cf_2": str(ctx["oper_cf"][-4]) if len(years)>=4 else "2",
        "oper_cf_3": str(ctx["oper_cf"][-3]),
        "oper_cf_4": str(ctx["oper_cf"][-2]),
        "oper_cf_5": str(ctx["oper_cf"][-1]),
        "invest_cf_1": str(ctx["invest_cf"][-5]) if len(years)>=5 else "-5",
        "invest_cf_2": str(ctx["invest_cf"][-4]) if len(years)>=4 else "-4",
        "invest_cf_3": str(ctx["invest_cf"][-3]),
        "invest_cf_4": str(ctx["invest_cf"][-2]),
        "invest_cf_5": str(ctx["invest_cf"][-1]),
        "fin_cf_1": str(ctx["fin_cf"][-5]) if len(years)>=5 else "3",
        "fin_cf_2": str(ctx["fin_cf"][-4]) if len(years)>=4 else "4",
        "fin_cf_3": str(ctx["fin_cf"][-3]),
        "fin_cf_4": str(ctx["fin_cf"][-2]),
        "fin_cf_5": str(ctx["fin_cf"][-1]),
        "cash_1": str(ctx["cash"][-5]) if len(years)>=5 else "5",
        "cash_2": str(ctx["cash"][-4]) if len(years)>=4 else "6",
        "cash_3": str(ctx["cash"][-3]),
        "cash_4": str(ctx["cash"][-2]),
        "cash_5": str(ctx["cash"][-1]),
        # --- Forecast ---
        "revenue_start_year": str(years[-1]),
        "revenue_end_year": str(years[-1]+5),
        "revenue_start": f"{revs[-1]}억 원",
        "revenue_end": f"{int(revs[-1]*(1+ctx['cagr']/100)**5)}억 원",
        "revenue_cagr": f"{ctx['cagr']}%",
        "profit_drivers": choice_sent(["제품 믹스 개선","단가 인상","수율 향상","규모의 경제"]),
        "profit_margin_start": f"{ctx['opm_start']}%",
        "profit_margin_end": f"{ctx['opm_end']}%",
        "debt_reduction_plan": "리파이낸싱·차입 구조 장기화",
        "financial_restructuring_measures": "비핵심 자산 매각·운전자본 효율화",
        "cagr_base": f"{ctx['cagr']}%",
        "base_case_assumption": "수주 정상 진행",
        "cagr_optimistic": f"{pct(ctx['cagr']+2, ctx['cagr']+8,1)}%",
        "optimistic_assumption": "신규 고객/해외 수주",
        "cagr_pessimistic": f"{pct(max(1,ctx['cagr']-6), max(3,ctx['cagr']-1),1)}%",
        "pessimistic_assumption": "수주 지연·가격 압박",
        "internal_ir_plan": "추가 투자 유치·IR 활동",
        "government_policy_impact": "지원/규제 변화 민감",
        "export_delay_impact": "수출 승인 지연 시 매출 인식 지체",
        # --- Valuation ---
        "dcf_method": "FCFF",
        "dcf_period": f"{years[-1]+1}~{years[-1]+5}",
        "dcf_wacc": f"{pct(6,12,1)}%",
        "dcf_ev_range": f"{int(ctx['ev']*0.9)}억 ~ {int(ctx['ev']*1.1)}억",
        "dcf_terminal_growth": f"{pct(1.0,3.0,1)}%",
        "comps_domestic_peers": klist(),
        "comps_foreign_peers": klist(),
        "comps_ev_ebitda": f"{ctx['ev_ebitda']}배",
        "comps_ev_ebitda_avg": f"{round(ctx['ev_ebitda']*random.uniform(0.9,1.1),1)}배",
        "comps_ev_ebitda_target": f"{round(ctx['ev_ebitda']*random.uniform(0.8,1.1),1)}배",
        "comps_pe_ratio": f"{pct(8,30,1)}배",
        "comps_pb_ratio": f"{pct(0.8,4.0,1)}배",
        "benchmark_name": "업계 평균",
        "benchmark_ev_ebitda": f"{round(ctx['ev_ebitda']*random.uniform(0.95,1.05),1)}배",
        "benchmark_pe": f"{pct(10,28,1)}배",
        "benchmark_pb": f"{pct(1.2,3.0,1)}배",
        "target_name": ctx["name"].replace("주식회사 ",""),
        "target_ev_ebitda": f"{ctx['ev_ebitda']}배",
        "target_pe": f"{pct(8,26,1)}배",
        "target_pb": f"{pct(1.0,3.0,1)}배",
        "precedent_period": "5",
        "precedent_count": str(random.randint(3,7)),
        "precedent_ev_rev_multiple": f"{ctx['ev_rev']}배",
        "precedent_ev_ebitda_multiple": f"{ctx['ev_ebitda']}배",
        "deal_value_1": str(int(ctx['ev']*random.uniform(0.6,1.2))),
        "ev_rev_1": f"{pct(max(0.8,ctx['ev_rev']-0.6), ctx['ev_rev']+0.6,1)}배",
        "ev_ebitda_1": f"{pct(max(4,ctx['ev_ebitda']-2), ctx['ev_ebitda']+2,1)}배",
        "acquirer_1": company_name().replace("주식회사 ",""),
        "target_1": company_name().replace("주식회사 ",""),
        "deal_value_2": str(int(ctx['ev']*random.uniform(0.6,1.2))),
        "ev_rev_2": f"{pct(max(0.8,ctx['ev_rev']-0.6), ctx['ev_rev']+0.6,1)}배",
        "ev_ebitda_2": f"{pct(max(4,ctx['ev_ebitda']-2), ctx['ev_ebitda']+2,1)}배",
        "acquirer_2": company_name().replace("주식회사 ",""),
        "target_2": company_name().replace("주식회사 ",""),
        "deal_value_3": str(int(ctx['ev']*random.uniform(0.6,1.2))),
        "ev_rev_3": f"{pct(max(0.8,ctx['ev_rev']-0.6), ctx['ev_rev']+0.6,1)}배",
        "ev_ebitda_3": f"{pct(max(4,ctx['ev_ebitda']-2), ctx['ev_ebitda']+2,1)}배",
        "acquirer_3": company_name().replace("주식회사 ",""),
        "target_3": company_name().replace("주식회사 ",""),
        # --- Valuation result ---
        "val_method_1": "DCF", "val_method_2": "Comps", "val_method_3": "Precedent",
        "val_ev_1": f"{ctx['ev']}억 원",
        "val_eq_1": f"{ctx['eq']}억 원",
        "val_pps_1": ctx["pps"],
        "val_ev_2": f"{int(ctx['ev']*random.uniform(0.9,1.1))}억 원",
        "val_multiple_2": f"EV/EBITDA {ctx['ev_ebitda']}배",
        "val_ev_3": f"{int(ctx['ev']*random.uniform(0.8,1.0))}억 원",
        "val_multiple_3": f"EV/Rev {ctx['ev_rev']}배",
        "val_range_low": f"{int(ctx['ev']*0.9)}억 원",
        "val_low_comment": "보수적 가정",
        "val_range_high": f"{int(ctx['ev']*1.1)}억 원",
        "val_high_comment": "낙관적 가정",
        "val_range_margin": f"{pct(3,8,1)}%",
        "val_irr_target": f"{pct(12,25,1)}%",
        "val_negotiation_scenario": "투자자-창업주 밸류 괴리 조정 필요",
        # --- 민감도 ---
        "wacc_delta": f"{pct(0.5,2.0,1)}",
        "wacc_impact": f"{pct(3,10,1)}",
        "growth_delta": f"{pct(0.3,1.5,1)}",
        "growth_impact": f"{pct(3,9,1)}",
        "ebitda_delta": f"{pct(1,4,1)}",
        "ev_impact": f"{random.randint(50,500)}",
        "wacc_1": f"{pct(6,8,1)}%", "wacc_2": f"{pct(8,9,1)}%", "wacc_3": f"{pct(9,10,1)}%", "wacc_4": f"{pct(10,12,1)}%",
        "growth_1": f"{pct(1.0,1.8,1)}%", "growth_2": f"{pct(2.0,2.6,1)}%", "growth_3": f"{pct(2.7,3.5,1)}%",
        "val_11": str(random.randint(150,300)), "val_12": str(random.randint(140,280)),
        "val_13": str(random.randint(130,260)), "val_14": str(random.randint(120,240)),
        "val_21": str(random.randint(160,320)), "val_22": str(random.randint(150,300)),
        "val_23": str(random.randint(140,280)), "val_24": str(random.randint(130,260)),
        "val_31": str(random.randint(170,340)), "val_32": str(random.randint(160,320)),
        "val_33": str(random.randint(150,300)), "val_34": str(random.randint(140,280)),
        "sensitivity_conclusion1": "WACC·성장률에 높은 민감도",
        "sensitivity_conclusion2": "EBITDA 마진 변화 시 EV 변동 폭 큼",
        # --- 리스크 상세/대응 ---
        "fin_risk_title1": "차입 구조 불안정", "fin_risk_detail1": "단기차입 만기 집중", "fin_risk_detail2":"리파이낸싱 필요",
        "fin_risk_title2": "현금흐름 변동성", "fin_risk_detail3": "운전자본 소요 확대", "fin_risk_detail4": "계절성 수주",
        "fin_risk_title3": "이익 변동성", "fin_risk_detail5":"환율·원자재 가격 영향", "fin_risk_detail6":"단가 인하 압박",
        "operational_risk_title1":"핵심 인력 의존", "operational_risk_detail1":"R&D 인력 이탈 리스크", "operational_risk_detail2":"프로세스 미비",
        "operational_risk_title2":"공급망 제약", "operational_risk_detail3":"특정 협력사 의존", "operational_risk_detail4":"대체선 부족",
        "operational_risk_title3":"내부 통제", "operational_risk_detail5":"매출 인식 지연", "operational_risk_detail6":"원가 집계 지연",
        "market_risk_title1":"수요 변동성", "market_risk_detail1":"경기·정책 민감", "market_risk_detail2":"고객 이탈 시 리스크",
        "market_risk_title2":"경쟁 심화", "market_risk_detail3":"해외 경쟁사 진입", "market_risk_detail4":"가격 경쟁 격화",
        "market_risk_title3":"규모의 경제 미흡", "market_risk_detail5":"단가 인하 대응력 약화", "market_risk_detail6":"마케팅 투자 필요",
        "domestic_regulation":"국내 규제 강화", "domestic_impact":"인증·비용 증가",
        "policy_variable":"정책 불확실성", "policy_impact":"투자 지연·수요 위축",
        "approval_target":"신제품 인증/허가", "approval_impact":"출시 지연·매출 인식 지체",
        "foreign_region":"해외 시장", "foreign_policy":"데이터/수출 제한", "foreign_impact":"현지 JV 필요",
        "jv_negotiation_status":"지분·지배 구조 협상 진행",
        "region_domestic":"국내", "regulation_domestic":"개보법/ESG", "impact_domestic":"비용 증가",
        "region_agency":"국방부/규제기관", "regulation_agency":"과제/인허가", "impact_agency":"지연 가능",
        "region_foreign1":"미국", "regulation_foreign1":"ITAR", "impact_foreign1":"승인 지연",
        "region_foreign2":"EU", "regulation_foreign2":"데이터 국외반출", "impact_foreign2":"현지화 필요",
        "financial_risk_strategy":"만기 분산·금리 스왑 등",
        "financial_refinancing_status":"주요 은행과 조건 협의",
        "operational_risk_retention":"핵심 인력 리텐션 프로그램",
        "operational_control_improvement":"ERP/내부통제 개선",
        "market_risk_new_clients":"신규 산업군 공략",
        "market_risk_contract_terms":"SLA/위약금 조항 강화",
        "regulatory_risk_consulting":"인증 전문 인력 확충",
        "regulatory_risk_team_setup":"규제 대응 전담팀 신설",
        "risk_high":"고위험(High) : 고객 집중, 단기차입",
        "risk_medium":"중위험(Medium) : 환율·규제 변수",
        "risk_low":"저위험(Low) : 내부 시스템 개선 이슈",
        # --- 투자 구조 상세/계약 ---
        "equity_investment_details":"보통주·우선주 발행 통한 자본 확충",
        "convertible_investment_details":"전환사채 기반 하방 방어",
        "mezzanine_investment_details":"부채·자본 성격 병행",
        "internal_review_result":"보통주 단독은 비효율, 혼합 구조 적정",
        "investment_structure".lower(): "혼합(신주+구주)",
        "new_investors":"신규 투자자",
        "new_investor_share": f"{pct(10,35,1)}%",
        "new_investor_cond": "성과 조건부",
        "founder_share_after": "창업주 지분 희석 관리",
        "strategic_investors_after":"전략투자자 일부 Exit 가능",
        "shareholder_change_note":"지분 구조 변동성 존재",
        "exit_negotiation_note":"블록딜/협상 여지",
        "founder_defense_note":"우호 지분 확보 논의",
        "pref_share_cond":"우선주 권리", "pref_share_value":"배당률/청산우선권",
        "drag_tag_cond":"Drag/Tag", "drag_tag_value":"동반매도/매수권",
        "protective_provisions":"보호조항", "provisions_value":"중요안건 동의",
        "investor_targets":"내부 목표", "targets_value":"IRR/MOIC",
        "valuation_point":"신주 발행가", "valuation_value":f"EV {ctx['ev']}억 기준",
        "lockup_cond":"Lock-up", "lockup_value":f"{random.randint(1,4)}년",
        "option_cond":"옵션", "option_value":"성과 조건부 콜옵션",
        "internal_review":"조건 미일치 시 결렬 위험",
        # --- Exit 시뮬 ---
        "ipo_market": random.choice(["KOSDAQ","KOSPI","NASDAQ"]),
        "ipo_year": str(datetime.now().year+random.randint(2,5)),
        "ipo_valuation": f"{int(ctx['ev']*random.uniform(1.3,1.8))}억 원",
        "ipo_irr": f"{pct(12,28,1)}%",
        "ipo_risk":"시장 침체/규제 변수",
        "mna_foreign_investors": klist(item="회사"),
        "mna_domestic_investors": klist(item="회사"),
        "mna_probability": f"{pct(20,70,1)}%",
        "mna_negotiation_status":"비공개 타진/예비 협의",
        "secondary_desc":"재무적 투자자 대상 지분 매각",
        "secondary_condition":"블록딜 할인률 고려",
        "buyback_feasibility":"현금흐름 제약으로 제한적",
        "buyback_short_term":"단기 실행 현실성 낮음",
        "sim_ipo_irr": f"{pct(12,28,1)}%", "sim_ipo_moic": f"{pct(1.5,2.4,1)}배",
        "sim_mna_irr": f"{pct(10,22,1)}%", "sim_mna_moic": f"{pct(1.3,2.0,1)}배",
        "sim_secondary_irr": f"{pct(8,16,1)}%", "sim_secondary_moic": f"{pct(1.1,1.6,1)}배",
        "sim_ipo_additional":"상장 시기 민감",
        "sim_mna_additional":"지분율·경영권 이슈",
        "sim_secondary_additional":"유동성·할인율 변수",
        # --- 결론/액션 ---
        "final_recommendation": random.choice(["Invest","Conditional Invest","Reject"]),
        "final_recommendation_rationale":"성장성 대비 리스크 균형 고려",
        "condition_1":"차입 구조 개선", "condition_2":"핵심 인력 리텐션", "condition_3":"규제 대응 로드맵",
        "positive_factor_1":"성장 시장·핵심 고객 확보",
        "positive_factor_2":"제품 경쟁력·가격 우위",
        "positive_factor_3":"해외 확장 잠재력",
        "negative_factor_1":"고객 집중·수요 변동",
        "negative_factor_2":"규제·인증 지연 리스크",
        "negative_factor_3":"운전자본 부담",
        "action_short_1":"리파이낸싱 추진", "action_short_2":"계약 재협상", "action_short_3":"규제 컨설팅 착수",
        "action_mid_1":"신규 사업부 KPI 관리", "action_mid_2":"리텐션 프로그램", "action_mid_3":"ESG 체계 구축",
        "action_long_1":"IPO/M&A 준비", "action_long_2":"해외 채널 확장", "action_long_3":"정기 IR 체계",
        "forecast_revenue": f"CAGR {ctx['cagr']}% 가정",
        "forecast_profit": "영업이익률 점진 개선",
        "forecast_debt_ratio": f"{pct(60,90,0)}% 목표",
        "forecast_exit": f"{random.randint(3,5)}년 내 IPO/M&A",
        # --- 부록/인터뷰/출처 ---
        "fin_years": "5",
        "income_sales_start": f"{revs[-5] if len(revs)>=5 else revs[0]}억", "income_sales_end": f"{revs[-1]}억",
        "income_period": f"{len(ctx['years'])}년", "income_cagr": f"{ctx['cagr']}",
        "balance_assets_start": f"{assets[-3]}억", "balance_assets_end": f"{assets[-1]}억",
        "balance_debt_ratio_start": f"{int(dr[-3])}", "balance_debt_ratio_end": f"{int(dr[-1])}",
        "cashflow_operating": f"{ctx['oper_cf'][-1]}억", "cashflow_investing": f"{ctx['invest_cf'][-1]}억", "cashflow_financing": f"{ctx['fin_cf'][-1]}억", "cashflow_year": f"{years[-1]}",
        "division_1_name": ctx['prof']['products'][0], "division_1_sales": f"{int(revs[-1]*random.uniform(0.3,0.6))}억", "division_1_ratio": f"{pct(30,60,1)}", "division_1_margin": f"{pct(8,20,1)}",
        "division_2_name": ctx['prof']['products'][1] if len(ctx['prof']['products'])>1 else "신사업", "division_2_sales": f"{int(revs[-1]*random.uniform(0.2,0.5))}억", "division_2_ratio": f"{pct(20,50,1)}", "division_2_margin": f"{pct(5,15,1)}",
        "division_3_name": "서비스/유지보수", "division_3_sales": f"{int(revs[-1]*random.uniform(0.1,0.3))}억", "division_3_ratio": f"{pct(10,30,1)}", "division_3_margin": f"{pct(5,20,1)}",
        "data_source_domestic":"KOSIS, 산업부, KRX", "data_source_global":"Bloomberg, Capital IQ", "data_source_report":"증권사 산업리포트", "data_source_internal":"내부 설문/조달 데이터",
        "ceo_quote":"핵심 제품 해외 진출 가속", "ceo_note":"전략 고객 확보 최우선",
        "cfo_name": person("CFO"), "cfo_quote":"유동성 리스크 선제 대응", "cfo_note":"만기 분산·금리 헤지",
        "cto_name": person("CTO"), "cto_quote":"제품 신뢰성·효율 고도화", "cto_note":"인증 로드맵 추진",
        "client_quote":"공급 안정성·기술 지원 중요", "client_note":"SLA 강화 필요",
        "supplier_name": company_name().replace("주식회사 ",""), "supplier_quote":"단가·납기 협상 진행",
        "law_name":"관련 법령", "law_year": str(yr()), "court_name":"○○지법",
        "guideline_source":"정부 가이드라인", "guideline_year": str(yr()),
        "email_date": f"{datetime.now().year}.{random.randint(1,12):02d}",
        "cfo_memo":"단기 유동성 관리 필요",
        "ceo_memo":"해외 파트너십·JV 협상 중",
        "report_source":"산업 리포트", "report_year": str(yr()),
        "paper_title":"학술 논문 제목", "paper_source":"저널명", "paper_year": str(yr()),
        "news_title":"산업 동향 기사", "news_source":"Daily Biz", "news_date": f"{datetime.now().year}.{random.randint(1,12):02d}",
        "harvard_citation":"Hong, J. (2024). Market.", "apa_citation":"Kim (2023). Journal, 12(2).",
    }
    return mapping.get(key)

# --- Fallback by token name (semantic heuristics) ---
KW = {
    "money": ["매출","금액","amount","ev","valuation","capex","dividend","투자","현금","cash","deal","price"],
    "pct": ["비율","율","margin","irr","moic","roe","roa","cagr","growth","rate","ratio","점유"],
    "date": ["date","일자","yyyymm","기간","year","연도","timeline","시점"],
    "name": ["company","기업","법인","acquirer","target","competitor","investor","peer","고객","협력사","담당"],
    "person": ["ceo","cfo","cto","이사","대표","책임","manager","담당"],
    "addr": ["address","주소","hq","본사"],
    "list": ["peers","clients","고객사","competitor","주주","shareholder","investors","업체","목록","list"],
    "text": ["note","memo","comment","설명","특징","리스크","전략","요약","배경","환경","구조","방안","조건"],
}
def fallback_by_name(token:str, ctx:dict):
    k = token.lower()
    # money
    if any(x in k for x in KW["money"]):
        return money_from_rev(ctx["revs"][-1], 0.01, 0.4)
    if any(x in k for x in KW["pct"]):
        return f"{pct(1,40,1)}%"
    if any(x in k for x in KW["date"]):
        return f"{datetime.now().year + random.randint(0,5)}-{random.randint(1,12):02d}"
    if any(x in k for x in KW["person"]):
        role = "임원" if "임원" in k else ("CEO" if "ceo" in k else ("CFO" if "cfo" in k else ("CTO" if "cto" in k else "담당")))
        return person(role)
    if any(x in k for x in KW["addr"]):
        return address()
    if any(x in k for x in KW["name"]):
        return company_name().replace("주식회사 ","")
    if any(x in k for x in KW["list"]):
        return klist()
    if any(x in k for x in KW["text"]):
        return choice_sent([
            "내부 검토 결과에 따라 단계적 추진",
            "시장·규제 변동을 반영한 조건부 접근",
            "비용 구조 개선과 고객 다변화 병행"
        ])
    # default sentence
    return choice_sent([
        "핵심 가정 하에서 단계적 성장 가시성 확보",
        "신규 고객·제품 믹스 개선을 통한 수익성 제고",
        "리스크 관리와 사업 확장 간 균형 필요"
    ])

def build_mapping(tokens:list, seed=None):
    if seed is not None:
        random.seed(seed); Faker.seed(seed)
    ctx = gen_context()
    mapping = {}
    # 프롬프트 토큰이 있을 경우 주입용 문자열
    def ai_prompt():
        sample = ", ".join(sorted(tokens)[:25])
        return (f"[지시사항] 모든 값은 전부 허구로 작성한다. 업종='{ctx['sector']}', 규모='{ctx['scale']}', 단계='{ctx['stage']}'. "
                f"실제 기업/인물/주소/뉴스를 참조하지 말고 자연스러운 한국어 투자보고서 톤으로 서술. "
                f"[참고 토큰 일부] {sample} ...")
    for t in tokens:
        key = normalize_key(t)
        val = gen_known_value(key, ctx)
        if val is None:
            # 프롬프트 토큰?
            if key in ["ai_prompt_for_fake","llm_prompt","gen_prompt"]:
                val = ai_prompt()
            else:
                val = fallback_by_name(key, ctx)
        mapping[t] = val
    return mapping, ctx

# ---------------- Main ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--count", type=int, default=5)
    ap.add_argument("--seed", type=int, default=None)
    ap.add_argument("--dump-jsonl", default="")
    args = ap.parse_args()

    template = Path(args.template)
    outdir = Path(args.out); outdir.mkdir(parents=True, exist_ok=True)

    tokens = extract_tokens_from_hwpx(template)
    if not tokens: raise SystemExit("템플릿에서 토큰을 찾지 못했습니다.")
    print(f"[info] tokens: {len(tokens)}개")

    prefix="가짜투자보고서"; ext="hwpx"
    existing=[int(m.group(1)) for p in glob(str(outdir/f"{prefix}_*.{ext}")) if (m:=re.search(rf"{prefix}_(\d+)\.{ext}$", Path(p).name))]
    start=(max(existing) if existing else 0)+1

    log = open(args.dump_jsonl,"a",encoding="utf-8") if args.dump_jsonl else None
    for i in range(args.count):
        seed = (args.seed or int(time.time())) + i
        mapping, ctx = build_mapping(tokens, seed=seed)
        out_path = outdir / f"{prefix}_{start+i:02d}.{ext}"
        replace_tokens_in_hwpx(template, out_path, mapping)
        if log:
            log.write(json.dumps({
                "file": out_path.name,
                "seed": seed,
                "sector": ctx["sector"],
                "scale": ctx["scale"],
                "stage": ctx["stage"],
                "revenue_last": ctx["revs"][-1],
                "cagr": ctx["cagr"]
            }, ensure_ascii=False)+"\n")
        if i==0:
            # 콘솔에 프롬프트 예시 한 번만
            print("\n[LLM PROMPT SAMPLE]\n", (mapping.get("AI_PROMPT_FOR_FAKE") or mapping.get("LLM_PROMPT") or mapping.get("GEN_PROMPT") or "템플릿에 프롬프트 토큰이 없습니다."))
    if log: log.close()
    print("[done] 출력 폴더:", outdir)

if __name__ == "__main__":
    main()
