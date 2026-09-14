# -*- coding: utf-8 -*-
"""
증정 트래킹 대시보드 데이터 생성
  1) 신세계 Cloud POS「영수증별 매출 상세현황」에서 기간 영수증 다운로드
  2) 증정 실제(①) / 이론(②) 대조 후 data/gift.json 생성

사용:  python3 gift_update.py              (설정의 시작일~오늘)
       python3 gift_update.py 2026-09-15 2026-09-21
       python3 gift_update.py --file 받아둔.xlsx      (다운로드 없이 계산만)
"""
import sys, json, math, asyncio, datetime, glob
from pathlib import Path
from collections import defaultdict
import openpyxl

# ════════ 행사 설정 — 행사가 바뀌면 여기만 고치세요 ════════
CONFIG = {
    "행사명":     "골프우승 증정 프로모션",
    "매장코드":   "4012",
    "매장명":     "호우섬 영등포 타임스퀘어점",
    "증정품문자": "[우승기념]",   # 상품명에 이 글자가 들어가면 증정품
    "최소금액":   50000,          # 총매출(할인 전·VAT 포함) 이 금액 미만 → 0개
    "개당금액":   100000,         # 기준 초과 시 이 금액마다 1개(올림) = 객단가 25,000 × 4인
    "시작일":     "2026-09-15",
    "종료일":     "2026-09-21",   # 비우면 오늘까지
}
# ══════════════════════════════════════════════════════

BASE = Path(__file__).resolve().parent
OUT  = BASE / "data" / "gift.json"
DL   = BASE / ".gift_raw"


# ───────────────────────── 집계 ─────────────────────────
def load_rows(paths):
    rows = []
    for p in paths:
        ws = openpyxl.load_workbook(p, data_only=True).active
        for r in range(3, ws.max_row + 1):
            v = [ws.cell(r, c).value for c in range(1, 27)]
            if v[6] is None:
                continue
            rows.append(v)
    return rows


def analyze(rows, cfg):
    gift_mark = cfg["증정품문자"]
    lo, per = cfg["최소금액"], cfg["개당금액"]

    rec = {}
    for v in rows:
        k = (v[4], v[6])                       # 매출일자, 거래번호
        d = rec.setdefault(k, {
            "일자": v[4], "거래": v[6], "pos": v[5], "테이블": v[8] or "",
            "결제": v[9] or "", "주문": v[10] or "",
            "총매출": v[12] or 0, "할인": v[16] or 0, "고객수": v[17] or 0,
            "증정": 0,
        })
        if v[22] and gift_mark in str(v[22]):
            d["증정"] += int(v[24] or 0)

    def gn(x):
        try: return int(x["거래"])
        except Exception: return 0

    rows_s = sorted(rec.values(), key=lambda x: (x["일자"], gn(x)))
    for x in rows_s:
        x["상태"] = ""; x["짝"] = ""

    # 반품 ↔ 취소원거래 순차 매칭 (교차일 허용)
    #   1순위: 주문시간 + 테이블 + 금액   2순위: 테이블 + 금액 (가장 가까운 이전 건)
    for i, rf in enumerate(rows_s):
        if rf["총매출"] >= 0:
            continue
        rf["상태"] = "반품"
        t = -rf["총매출"]
        prev = rows_s[:i]
        c = [o for o in prev if o["총매출"] == t and not o["상태"]
             and o["주문"] == rf["주문"] and o["테이블"] == rf["테이블"]]
        if not c:
            c = [o for o in prev if o["총매출"] == t and not o["상태"]
                 and o["테이블"] == rf["테이블"]]
        if c:
            o = c[-1]
            o["상태"] = "취소원거래"
            o["짝"] = rf["거래"]; rf["짝"] = o["거래"]

    def theory(m):
        return 0 if m < lo else math.ceil(m / per)

    live = []
    for x in rows_s:
        if x["상태"] or x["총매출"] < 0:
            x["이론"] = None; x["판정"] = x["상태"] or "반품"
            continue
        x["이론"] = theory(x["총매출"])
        x["판정"] = ("누락" if x["증정"] < x["이론"]
                     else "초과" if x["증정"] > x["이론"] else "정상")
        live.append(x)

    days = defaultdict(lambda: dict(전체=0, 취소=0, 유효=0, 총매출=0, 고객=0,
                                    대상=0, 이론=0, 실제=0, 누락=0, 초과=0))
    for x in rows_s:
        d = days[x["일자"]]
        d["전체"] += 1
        if x["상태"]:
            d["취소"] += 1; continue
        d["유효"] += 1; d["총매출"] += x["총매출"]; d["고객"] += int(x["고객수"] or 0)
        d["이론"] += int(x["이론"]); d["실제"] += int(x["증정"])
        if x["이론"] > 0: d["대상"] += 1
        if x["판정"] == "누락": d["누락"] += 1
        if x["판정"] == "초과": d["초과"] += 1

    issues = [{
        "일자": x["일자"], "거래": x["거래"], "pos": x["pos"], "테이블": x["테이블"],
        "결제": x["결제"], "총매출": x["총매출"], "할인": x["할인"],
        "실제": x["증정"], "이론": x["이론"], "판정": x["판정"],
    } for x in rows_s if x["판정"] in ("누락", "초과")]

    chains = [{
        "일자": x["일자"], "거래": x["거래"], "총매출": x["총매출"],
        "테이블": x["테이블"], "짝": x["짝"], "상태": x["상태"],
    } for x in rows_s if x["상태"] == "취소원거래"]

    g1 = sum(x["증정"] for x in live)
    g2 = sum(x["이론"] for x in live)
    return {
        "행사명": cfg["행사명"], "매장명": cfg["매장명"],
        "기준": {"최소금액": lo, "개당금액": per, "증정품문자": gift_mark},
        "갱신": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "마지막결제": max((x["결제"] for x in rows_s), default=""),
        "합계": {
            "영수증": len(rows_s), "취소": len(rows_s) - len(live), "유효": len(live),
            "총매출": sum(x["총매출"] for x in live),
            "고객": int(sum(x["고객수"] for x in live)),
            "대상": sum(1 for x in live if x["이론"] > 0),
            "이론": int(g2), "실제": int(g1), "차이": int(g1 - g2),
            "제공율": (g1 / g2) if g2 else None,
            "누락": sum(1 for x in live if x["판정"] == "누락"),
            "초과": sum(1 for x in live if x["판정"] == "초과"),
        },
        "일자별": [dict(일자=k, **v) for k, v in sorted(days.items())],
        "이슈": issues,
        "취소체인": chains,
    }


# ───────────────────────── POS 다운로드 ─────────────────────────
URL = "https://cloudposoffice.shinsegae.com/"
CRED = BASE / ".pos_credentials.json"   # git에 올라가지 않음(.gitignore)


def credentials():
    """POS 사번/비밀번호 — 환경변수 POS_ID/POS_PW 우선, 없으면 .pos_credentials.json"""
    import os
    uid, pw = os.environ.get("POS_ID"), os.environ.get("POS_PW")
    if uid and pw:
        return uid, pw
    if CRED.exists():
        c = json.loads(CRED.read_text(encoding="utf-8"))
        if c.get("empCd") and c.get("password"):
            return c["empCd"], c["password"]
    raise SystemExit(
        f"POS 로그인 정보가 없습니다.\n"
        f"  {CRED} 파일을 아래 형식으로 만드세요:\n"
        f'  {{"empCd": "사번", "password": "비밀번호"}}\n'
        f"  (또는 환경변수 POS_ID / POS_PW 설정)")


PFX = "mf_tac_layout_contents_003002010_body_"
POP = PFX + "TNANT_SEARCH_POP_wframe_"


async def fetch(code, start, end, outdir):
    from playwright.async_api import async_playwright
    Path(outdir).mkdir(parents=True, exist_ok=True)
    s8, e8 = start.replace("-", ""), end.replace("-", "")
    async with async_playwright() as p:
        b = await p.chromium.launch(headless=True)
        page = await (await b.new_context(accept_downloads=True)).new_page()
        await page.goto(URL, wait_until="networkidle", timeout=60000)
        uid, pw = credentials()
        await page.fill("#mf_ibx_empCd", uid); await page.fill("#mf_sct_password", pw)
        await page.locator("a").filter(has_text="로그인").first.click()
        await page.wait_for_timeout(4000)
        await page.locator("#mf_wfm_side_gen_menu1_2_btn_menu1").click(); await page.wait_for_timeout(1500)
        await page.locator("#mf_wfm_side_gen_menu2_1_btn_menu2").click(); await page.wait_for_timeout(1500)
        await page.locator("#mf_wfm_side_gen_menu2_1_gen_menu3_8_btn_menu3").click(); await page.wait_for_timeout(6000)
        await page.evaluate(f"document.getElementById('{PFX}btn_searchTnantPop').click()")
        await page.wait_for_timeout(3500)
        await page.locator(f"#{POP}ibx_program").fill(code)
        await page.locator(f"#{POP}btn_search").click(); await page.wait_for_timeout(4000)
        cell = f"#{POP}grd_list_cell_0_5"
        if not await page.locator(cell).count():
            raise SystemExit(f"매장코드 {code} 조회 결과 없음")
        await page.locator(cell).dblclick(); await page.wait_for_timeout(2500)
        st = await page.evaluate(f"""()=>{{
          const s=(n,v)=>$p.getComponentById('{PFX}'+n).setValue(v);
          s('ica_startDate','{s8}'); s('ica_endDate','{e8}');
          const g=n=>String($p.getComponentById('{PFX}'+n).getValue());
          return {{cd:g('ibx_sTnantCd'), nm:g('ibx_sTnantCdNm')}};}}""")
        print(f"  매장 {st['cd']} {st['nm']} / {start} ~ {end}", flush=True)
        await page.locator(f"#{PFX}btn_SSearch").click()
        rows = 0
        for _ in range(24):
            await page.wait_for_timeout(5000)
            rows = await page.evaluate(
                f"()=>{{try{{return $p.getComponentById('{PFX}grd_main').getRowCount();}}catch(e){{return 0;}}}}")
            if rows: break
        if not rows:
            raise SystemExit("조회 0건")
        async with page.expect_download(timeout=300000) as dl:
            await page.locator(f"#{PFX}btn_PExcelGrp").click()
        d = await dl.value
        dest = Path(outdir) / f"{code}_{start}_{end}.xlsx"
        await d.save_as(str(dest))
        await b.close()
        return dest


def main():
    cfg = dict(CONFIG)
    args = sys.argv[1:]
    if args and args[0] == "--file":
        files = args[1:]
    else:
        start = args[0] if len(args) > 0 else cfg["시작일"]
        end   = args[1] if len(args) > 1 else (cfg["종료일"] or datetime.date.today().isoformat())
        cfg["시작일"], cfg["종료일"] = start, end
        print(f"[1/2] POS 다운로드 {start} ~ {end}", flush=True)
        import shutil
        if DL.exists(): shutil.rmtree(DL)
        files = [asyncio.run(fetch(cfg["매장코드"], start, end, DL))]
    print("[2/2] 집계", flush=True)
    data = analyze(load_rows(files), cfg)
    data["기간"] = f'{cfg["시작일"]} ~ {cfg["종료일"]}' if cfg.get("종료일") else cfg["시작일"]
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    s = data["합계"]
    print(f"  영수증 {s['영수증']}건 / 유효 {s['유효']}건 / 대상 {s['대상']}건")
    print(f"  ② {s['이론']}개 / ① {s['실제']}개 / 차이 {s['차이']:+d}"
          + (f" / 제공율 {s['제공율']*100:.1f}%" if s["제공율"] else ""))
    print(f"  누락 {s['누락']}건 · 초과 {s['초과']}건  →  {OUT}")


if __name__ == "__main__":
    main()
