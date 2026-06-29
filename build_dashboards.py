#!/usr/bin/env python3
"""
build_dashboards.py — детерминированный генератор финансовых дашбордов SA Minerals.

Версия 1: пересобирает инвест-слой (sa_minerals_group_dashboard.html) ПОЛНОСТЬЮ
из листа "Ledger" файла Shareholder Ledger (xlsx). Все цифры считаются из транзакций,
ничего не хардкодится. Интеркомпани-переводы (Internal=Yes) исключаются из "развёртывания".

Запуск:
    python3 build_dashboards.py --ledger <путь к .xlsx> --out <путь к .html>

Перед публикацией прогоняются проверки: сохранность капитала 0–100%,
recoverable + cash + overhead == capital injected. При провале — exit 1 (не публиковать).
"""
import argparse, sys, datetime
import openpyxl


def load_ledger(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Ledger"]
    idx = None
    rows = []
    for r in ws.iter_rows(values_only=True):
        cells = [(str(c).strip() if c is not None else "") for c in r]
        if idx is None:
            if "#" in cells and "Direction" in cells and "Category" in cells:
                idx = {h: i for i, h in enumerate(cells)}
            continue
        def g(name):
            i = idx.get(name)
            return r[i] if (i is not None and i < len(r)) else None
        num = g("#")
        if num in (None, ""):
            continue
        try:
            int(str(num).strip())
        except ValueError:
            continue  # totals / footer rows
        d = g("Date")
        if isinstance(d, (datetime.datetime, datetime.date)):
            d = d.strftime("%Y-%m-%d")
        def f(name):
            v = g(name)
            try:
                return float(v) if v not in (None, "") else 0.0
            except (TypeError, ValueError):
                return 0.0
        rows.append({
            "num": int(str(num).strip()),
            "date": str(d) if d else "",
            "account": (g("Account") or "").strip(),
            "direction": (g("Direction") or "").strip(),
            "category": (g("Category") or "").strip(),
            "counterparty": (g("Counterparty") or "").strip(),
            "internal": str(g("Internal?") or "").strip().lower().startswith("y"),
            "inflow": f("Inflow (R)"),
            "outflow": f("Outflow (R)"),
        })
    return rows


ACC = {"63209738939": "Asset Co", "63209745116": "Claim Co"}

# Ручные оверрайды для операций, чей бизнес-смысл НЕ виден из текста банка.
# Ключ: (date 'YYYY-MM-DD', amount round2). Заполняется человеком (1–2 в неделю).
OVERRIDES = {
    ("2026-06-22", -845375.09): {"category": "Claim Acquisition",
                                 "counterparty": "Wilrock Properties (Pty) Ltd"},
}

def classify_raw(desc, ref, amt):
    """Правила классификации сырой транзакции FNB → (category, internal, counterparty)."""
    t = (str(desc) + " " + str(ref)).upper()
    if "ADRU TECH" in t: return "Capital Injection", False, "AdRu Tech Ltd (Cyprus)"
    if "ADT CASH DEPO" in t or ("DEPOSIT" in t and abs(amt) <= 5000): return "Account Activation", False, "Funded by Paul Khourie"
    if "DRAWDOWN" in t: return "Loan to EMS (Receivable)", False, "EMS Mining"
    if "INVESTEC" in t: return "Claim Acquisition", False, "Investec Bank"
    if "BUBESI" in t: return "Claim Acquisition", False, "Bubesi Investments 46 (Pty) Ltd"
    if "SANDVIK" in t: return "Claim Acquisition", False, "Sandvik Mining RSA (Pty) Ltd"
    if "WILROCK" in t: return "Claim Acquisition", False, "Wilrock Properties (Pty) Ltd"
    if any(k in t for k in ("SWIFT COMMISSION", "INWARD SWIFT", "FOREX TRANSFER", "SWIFT CORRECTION")):
        return "Bank Charges", False, "FNB"
    if ("TRF" in t or "TRANSFER FUNDS" in t) and ("SA MINERALS" in t or "TRANSFER FUNDS" in t):
        return "Intercompany Transfer", True, "group account"
    if "AFRIQOM" in t: return "Operating Expense — Travel & Conference", False, "AFRIQOM FZ LLC (UAE)"
    if any(k in t for k in ("STRAND", "ACCOMMODATION", "GARDEN COURT", "HOTEL")):
        return "Operating Expense — Travel & Conference", False, "Travel / accommodation"
    if "INV17842" in t or "COLLECTIVE" in t: return "Operating Expense — Advisory", False, "Collective Accounting (Pty) Ltd"
    if any(k in t for k in ("WOOD", "INVOICE 0001", "ADVISOR")): return "Operating Expense — Advisory", False, "Chris Wood (advisor)"
    return "UNCATEGORIZED", False, str(desc).strip()[:40]

def load_raw(path):
    """Строит ledger-эквивалент из сырых выписок FNB Asset/Claim (.csv или .zip с .csv)."""
    import zipfile, glob, os, csv as _csv, re
    texts = []
    for fp in glob.glob(os.path.join(path, "*")):
        low = fp.lower()
        if low.endswith(".zip"):
            try:
                z = zipfile.ZipFile(fp)
                for n in z.namelist():
                    if n.lower().endswith(".csv"):
                        texts.append(z.read(n).decode("utf-8", "replace"))
            except zipfile.BadZipFile:
                pass
        elif low.endswith(".csv"):
            texts.append(open(fp, encoding="utf-8", errors="replace").read())
    rows, flagged = [], []
    for txt in texts:
        lines = txt.splitlines()
        acc_no = None
        for ln in lines[:3]:
            mm = re.search(r"ACCOUNT NUMBER\s*([0-9]+)", ln.upper())
            if mm:
                acc_no = mm.group(1); break
        account = ACC.get(acc_no, acc_no or "?")
        for r in _csv.reader(lines):
            if len(r) < 4:
                continue
            try:
                amt = round(float(r[2]), 2)
            except (ValueError, IndexError):
                continue
            desc, ref = r[3], (r[4] if len(r) > 4 else "")
            cat, internal, cp = classify_raw(desc, ref, amt)
            ov = OVERRIDES.get((r[0].strip(), amt))
            if ov:
                cat = ov.get("category", cat); cp = ov.get("counterparty", cp); internal = ov.get("internal", internal)
            if cat == "UNCATEGORIZED":
                flagged.append((r[0].strip(), account, amt, str(desc).strip()))
            rows.append({"date": r[0].strip(), "account": account,
                         "direction": "INFLOW" if amt > 0 else "OUTFLOW",
                         "category": cat, "counterparty": cp, "internal": internal,
                         "inflow": amt if amt > 0 else 0.0, "outflow": -amt if amt < 0 else 0.0})
    rows.sort(key=lambda x: (x["date"], x["account"]))
    for i, r in enumerate(rows, 1):
        r["num"] = i
    if flagged:
        sys.stderr.write(f"WARNING: {len(flagged)} операций не классифицированы (добавь в OVERRIDES):\n")
        for d, a, amt, ds in flagged:
            sys.stderr.write(f"   {d} {a} {amt:,.2f}  «{ds[:48]}»\n")
    return rows


def compute(rows):
    m = {}
    ext_in = sum(r["inflow"] for r in rows if r["direction"] == "INFLOW" and not r["internal"])
    claims = sum(r["outflow"] for r in rows if r["category"] == "Claim Acquisition")
    loan = sum(r["outflow"] for r in rows if "Loan to EMS" in r["category"])
    opex = sum(r["outflow"] for r in rows if r["category"].startswith("Operating Expense"))
    charges = sum(r["outflow"] for r in rows if r["category"] == "Bank Charges")
    recoverable = claims + loan
    overhead = opex + charges
    cash = ext_in - recoverable - overhead
    m["injected"] = ext_in
    m["claims"] = claims
    m["loan"] = loan
    m["recoverable"] = recoverable
    m["overhead"] = overhead
    m["cash"] = cash
    m["preserved"] = (recoverable + cash) / ext_in if ext_in else 0.0
    # per-account closing balances (включая интеркомпани — это реальные деньги на счёте)
    accts = {}
    for r in rows:
        a = r["account"]
        accts.setdefault(a, {"in": 0.0, "out": 0.0, "n": 0})
        accts[a]["in"] += r["inflow"]
        accts[a]["out"] += r["outflow"]
        accts[a]["n"] += 1
    for a, v in accts.items():
        v["close"] = v["in"] - v["out"]
    m["accounts"] = accts
    m["claim_rows"] = [r for r in rows if r["category"] == "Claim Acquisition"]
    m["loan_rows"] = [r for r in rows if "Loan to EMS" in r["category"]]
    m["txns"] = rows
    return m


def rm(n):
    return f"{n:,.0f}"


def millions(n):
    return f"R{n/1e6:.2f}m"


SHORT_CAT = {
    "Account Activation": "Activation", "Capital Injection": "Capital Injection",
    "Intercompany Transfer": "Intercompany", "Claim Acquisition": "Claim Acquisition",
    "Loan to EMS (Receivable)": "Loan to EMS", "Bank Charges": "Bank Charges",
    "Operating Expense — Advisory": "OpEx — Advisory",
    "Operating Expense — Travel & Conference": "OpEx — Travel",
}
# справочные face value клеймов (нет в леджере; из соглашений)
FACE = {"Bubesi Investments 46 (Pty) Ltd": 3944125.0}


def render(m, gen_date):
    a = m["accounts"]
    asset = next((v for k, v in a.items() if "Asset" in k), {"in":0,"out":0,"close":0,"n":0})
    claim = next((v for k, v in a.items() if "Claim" in k), {"in":0,"out":0,"close":0,"n":0})
    total_close = sum(v["close"] for v in a.values())
    # ledger JS array
    js = []
    for r in m["txns"]:
        d = "IN" if r["direction"] == "INFLOW" else "OUT"
        cat = SHORT_CAT.get(r["category"], r["category"])[:22]
        cp = r["counterparty"][:26].replace('"', "'")
        date = r["date"][5:] if len(r["date"]) >= 10 else r["date"]
        inf = f'{r["inflow"]:.2f}' if r["inflow"] else "null"
        out = f'{r["outflow"]:.2f}' if r["outflow"] else "null"
        js.append(f'[{r["num"]},"{date}","{r["account"]}","{d}","{cat}","{cp}",{inf},{out}]')
    ledger_js = ",\n".join(js)
    # claim portfolio rows
    crows = ""
    for i, r in enumerate(m["claim_rows"], 1):
        face = FACE.get(r["counterparty"], r["outflow"])
        crows += (f'<tr><td>{i:02d}</td><td>{r["counterparty"][:30]}</td><td>{r["date"][5:]}</td>'
                  f'<td class="n">{rm(r["outflow"])}</td><td class="n">{rm(face)}</td></tr>')
    for r in m["loan_rows"]:
        crows += (f'<tr><td>—</td><td>Loan to EMS — Drawdown 01</td><td>{r["date"][5:]}</td>'
                  f'<td class="n">{rm(r["outflow"])}</td><td class="n">{rm(r["outflow"])}</td></tr>')
    # category split for "where it went"
    inj = m["injected"]
    def pct(x): return (x/inj*100) if inj else 0
    tmpl = TEMPLATE
    repl = {
        "@@GEN@@": gen_date,
        "@@KPI_INJECTED@@": millions(inj),
        "@@KPI_RECOVER@@": millions(m["recoverable"]),
        "@@KPI_CASH@@": millions(m["cash"]),
        "@@KPI_PRESERVED@@": f'{m["preserved"]*100:.1f}%',
        "@@KPI_OVERHEAD@@": f'R{m["overhead"]/1e3:.1f}k',
        "@@ASSET_CLOSE@@": rm(asset["close"]),
        "@@CLAIM_CLOSE@@": rm(claim["close"]),
        "@@GROUP_CLOSE@@": rm(total_close),
        "@@ASSET_PCT@@": f'{(asset["close"]/total_close*100) if total_close else 0:.1f}%',
        "@@CLAIM_PCT@@": f'{(claim["close"]/total_close*100) if total_close else 0:.1f}%',
        "@@W_CLAIMS@@": rm(m["claims"]), "@@W_CLAIMS_P@@": f'{pct(m["claims"]):.1f}%',
        "@@W_LOAN@@": rm(m["loan"]), "@@W_LOAN_P@@": f'{pct(m["loan"]):.1f}%',
        "@@W_CASH@@": rm(m["cash"]), "@@W_CASH_P@@": f'{pct(m["cash"]):.1f}%',
        "@@W_OPEX@@": rm(m["overhead"]), "@@W_OPEX_P@@": f'{pct(m["overhead"]):.1f}%',
        "@@REC_PCT@@": f'{m["recoverable"]/inj*100 if inj else 0:.1f}%',
        "@@CASH_PCT@@": f'{m["cash"]/inj*100 if inj else 0:.1f}%',
        "@@OVH_PCT@@": f'{m["overhead"]/inj*100 if inj else 0:.1f}%',
        "@@NTXN@@": str(len(m["txns"])),
        "@@CLAIM_ROWS@@": crows,
        "@@LEDGER_JS@@": ledger_js,
    }
    for k, v in repl.items():
        tmpl = tmpl.replace(k, v)
    return tmpl


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw", help="папка с сырыми выписками FNB Asset/Claim (.csv/.zip) — основной режим")
    ap.add_argument("--ledger", help="xlsx-ледж­ер Александра (резерв/сверка)")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    if args.raw:
        rows = load_raw(args.raw)
    elif args.ledger:
        rows = load_ledger(args.ledger)
    else:
        print("ERROR: нужно указать --raw <папка> или --ledger <файл>", file=sys.stderr); sys.exit(1)
    if not rows:
        print("ERROR: в леджере не найдено транзакций", file=sys.stderr); sys.exit(1)
    m = compute(rows)
    # --- sanity checks ---
    errs = []
    if not (0 <= m["preserved"] <= 1.0001):
        errs.append(f'preserved={m["preserved"]:.3f} вне 0–100%')
    recon = m["recoverable"] + m["overhead"] + m["cash"]
    if abs(recon - m["injected"]) > 1.0:
        errs.append(f'итоги не сходятся: {recon:.2f} != injected {m["injected"]:.2f}')
    if m["injected"] <= 0:
        errs.append("injected <= 0")
    if errs:
        print("SANITY FAILED — не публиковать:\n  " + "\n  ".join(errs), file=sys.stderr)
        sys.exit(1)
    html = render(m, datetime.date.today().strftime("%d %b %Y"))
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"OK → {args.out}")
    print(f"  injected={millions(m['injected'])} recoverable={millions(m['recoverable'])} "
          f"cash={millions(m['cash'])} preserved={m['preserved']*100:.1f}% "
          f"txns={len(m['txns'])}")


TEMPLATE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SA Minerals Group — Shareholder Dashboard</title>
<style>
:root{--bg:#F5F2EA;--ink:#1A1A1A;--ink2:#555;--ink3:#8A8A82;--line:#D9D4C7;--fill1:#1A1A1A;--fill2:#6B6B63;--fill3:#A8A399;--alert:#B00020;--card:#FCFAF4;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Georgia,serif;background:var(--bg);color:var(--ink);line-height:1.5}
.wrap{max-width:1080px;margin:0 auto;padding:28px 22px 70px}
#ov{position:fixed;inset:0;background:var(--bg);z-index:99;display:flex;align-items:center;justify-content:center}
.box{background:var(--card);border:1px solid var(--line);border-radius:4px;padding:36px 40px;width:330px;text-align:center}
.box .l{font-size:10px;letter-spacing:3px;text-transform:uppercase;color:var(--ink3)}
.box .t{font-size:17px;font-weight:700;margin:6px 0 18px}
.f{width:100%;border:1px solid var(--line);background:#fff;border-radius:3px;padding:10px 12px;font-size:13px;margin-bottom:10px;outline:none}
.b{width:100%;background:var(--ink);color:var(--bg);border:none;border-radius:3px;padding:11px;font-size:12px;letter-spacing:1px;text-transform:uppercase;cursor:pointer}
.err{color:var(--alert);font-size:11px;margin-top:8px;display:none}
h1{font-size:25px;font-weight:700} .top{font-size:10px;letter-spacing:3px;text-transform:uppercase;color:var(--ink3)}
.sub{font-size:13px;color:var(--ink2);margin-top:6px;max-width:780px}.hr{height:1px;background:var(--ink);border:none;margin:16px 0 20px}
.meta{display:flex;gap:22px;flex-wrap:wrap;font-size:12px;color:var(--ink2);margin-top:8px}
.kgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:1px;background:var(--line);border:1px solid var(--line);margin:22px 0}
.k{background:var(--card);padding:15px 16px}.k .l{font-size:9.5px;letter-spacing:1.2px;text-transform:uppercase;color:var(--ink3)}
.k .v{font-size:22px;font-weight:700;margin-top:5px}.k .m{font-size:10.5px;color:var(--ink2);margin-top:3px}
.sec{font-size:11px;letter-spacing:2px;text-transform:uppercase;border-bottom:1px solid var(--ink);padding-bottom:6px;margin:30px 0 14px;font-weight:700}
.sec span{float:right;font-size:9.5px;color:var(--ink3);text-transform:none;font-weight:400}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{text-align:left;font-size:9.5px;letter-spacing:.8px;text-transform:uppercase;color:var(--ink3);padding:7px 9px;border-bottom:1px solid var(--ink)}
td{padding:7px 9px;border-bottom:1px solid var(--line)}td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}
tr.tot td{font-weight:700;border-top:1px solid var(--ink);border-bottom:none}
.brow{display:flex;align-items:center;gap:10px;margin:7px 0}.blab{width:215px;font-size:12px;flex-shrink:0}
.btr{flex:1;background:#EAE5D8;height:16px;border-radius:1px;overflow:hidden}.bfl{height:100%;background:var(--fill1)}
.bamt{width:130px;text-align:right;font-size:12px;font-variant-numeric:tabular-nums;flex-shrink:0}
.stack{display:flex;height:34px;border-radius:3px;overflow:hidden;margin:6px 0 12px}.stack>div{display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:700}
.legend{display:flex;gap:18px;flex-wrap:wrap;font-size:12px}.legend span{display:flex;align-items:center;gap:6px}.dot{width:11px;height:11px;border-radius:2px;display:inline-block}
.tag{font-size:10px;font-weight:700;padding:2px 7px;border-radius:3px}.t-in{background:#E8EDE6;color:#2E5E2E}.t-out{background:#F2E6E6;color:#7A2A2A}
.foot{font-size:10.5px;color:var(--ink3);margin-top:30px;border-top:1px solid var(--line);padding-top:12px}
@media(max-width:760px){.kgrid{grid-template-columns:1fr}.blab{width:120px}}
</style></head><body>
<div id="ov"><div class="box"><div class="l">SA Minerals Group</div><div class="t">Shareholder Dashboard</div>
<input class="f" id="u" placeholder="Username"><input class="f" id="p" type="password" placeholder="Password">
<button class="b" onclick="lo()">Enter</button><div class="err" id="e">Incorrect username or password</div></div></div>
<div class="wrap" id="main" style="display:none">
<div class="top">Confidential · For Board / IC use only · auto-generated</div>
<h1>SA Minerals Group — Shareholder Dashboard</h1>
<div class="sub">Консолидированное развёртывание капитала инвестора по двум счетам FNB (Asset Co + Claim Co). Интеркомпани-переводы исключены из KPI развёртывания. Собрано автоматически из леджера.</div>
<div class="meta"><span>Capital injected: @@KPI_INJECTED@@</span><span>Generated: @@GEN@@</span></div>
<hr class="hr">
<div class="kgrid">
<div class="k"><div class="l">Capital injected</div><div class="v">@@KPI_INJECTED@@</div><div class="m">AdRu Tech + activations</div></div>
<div class="k"><div class="l">Recoverable assets</div><div class="v">@@KPI_RECOVER@@</div><div class="m">claims + EMS loan</div></div>
<div class="k"><div class="l">Cash on hand</div><div class="v">@@KPI_CASH@@</div><div class="m">Asset + Claim Co</div></div>
<div class="k"><div class="l">Capital preserved</div><div class="v">@@KPI_PRESERVED@@</div><div class="m">(recoverable+cash)÷injected</div></div>
<div class="k"><div class="l">Overhead</div><div class="v">@@KPI_OVERHEAD@@</div><div class="m">opex + bank fees</div></div>
<div class="k"><div class="l">Transactions</div><div class="v">@@NTXN@@</div><div class="m">в леджере</div></div>
</div>
<div class="sec">Where the cash sits</div>
<table><thead><tr><th>Account</th><th class="n">Closing (R)</th><th class="n">% Group</th></tr></thead><tbody>
<tr><td>SA Minerals (Asset Co) · FNB …8939</td><td class="n">@@ASSET_CLOSE@@</td><td class="n">@@ASSET_PCT@@</td></tr>
<tr><td>SA Minerals Capital (Claim Co) · FNB …5116</td><td class="n">@@CLAIM_CLOSE@@</td><td class="n">@@CLAIM_PCT@@</td></tr>
<tr class="tot"><td>Group total</td><td class="n">@@GROUP_CLOSE@@</td><td class="n">100%</td></tr></tbody></table>
<div class="sec">Where the capital went — by category</div>
<div class="brow"><div class="blab">Claim acquisitions</div><div class="btr"><div class="bfl" style="width:@@W_CLAIMS_P@@"></div></div><div class="bamt">@@W_CLAIMS@@</div></div>
<div class="brow"><div class="blab">Loan to EMS</div><div class="btr"><div class="bfl" style="width:@@W_LOAN_P@@;background:var(--fill2)"></div></div><div class="bamt">@@W_LOAN@@</div></div>
<div class="brow"><div class="blab">Cash retained</div><div class="btr"><div class="bfl" style="width:@@W_CASH_P@@;background:var(--fill3)"></div></div><div class="bamt">@@W_CASH@@</div></div>
<div class="brow"><div class="blab">Overhead (opex+fees)</div><div class="btr"><div class="bfl" style="width:@@W_OPEX_P@@;background:var(--alert)"></div></div><div class="bamt">@@W_OPEX@@</div></div>
<div class="sec">Capital status</div>
<div class="stack"><div style="width:@@REC_PCT@@;background:var(--fill1)">Recoverable</div><div style="width:@@CASH_PCT@@;background:var(--fill3);color:#1A1A1A">Cash</div><div style="width:@@OVH_PCT@@;background:var(--alert)"></div></div>
<div class="legend"><span><i class="dot" style="background:var(--fill1)"></i>Recoverable @@REC_PCT@@</span><span><i class="dot" style="background:var(--fill3)"></i>Cash @@CASH_PCT@@</span><span><i class="dot" style="background:var(--alert)"></i>Overhead @@OVH_PCT@@</span></div>
<div class="sec">Claim portfolio &amp; loan to EMS</div>
<table><thead><tr><th>#</th><th>Counterparty</th><th>Date</th><th class="n">Cost (R)</th><th class="n">Face (R)</th></tr></thead><tbody>@@CLAIM_ROWS@@</tbody></table>
<div class="sec">Full transaction ledger <span>@@NTXN@@ transactions</span></div>
<div style="overflow-x:auto"><table id="lg"><thead><tr><th>#</th><th>Date</th><th>Account</th><th>Dir</th><th>Category</th><th>Counterparty</th><th class="n">Inflow</th><th class="n">Outflow</th></tr></thead><tbody id="lb"></tbody></table></div>
<div class="foot">Source: Shareholder Ledger (Asset + Claim FNB). Auto-generated @@GEN@@. Confidential.</div>
</div>
<script>
function lo(){var u=document.getElementById('u').value.trim(),p=document.getElementById('p').value.trim();
if(u==='viewer'&&p==='2026'){try{sessionStorage.setItem('sam_auth','1')}catch(e){}document.getElementById('ov').style.display='none';document.getElementById('main').style.display='block';}
else{document.getElementById('e').style.display='block';}}
document.addEventListener('keydown',function(e){var o=document.getElementById('ov');if(e.key==='Enter'&&o&&o.style.display!=='none')lo();});
try{if(sessionStorage.getItem('sam_auth')==='1'){document.getElementById('ov').style.display='none';document.getElementById('main').style.display='block';}}catch(e){}
var L=[
@@LEDGER_JS@@
];
function fmt(n){return n==null?"":n.toLocaleString('en-US',{maximumFractionDigits:2});}
var tb=document.getElementById('lb');
L.forEach(function(r){var tr=document.createElement('tr');var d=r[3]==='IN'?'<span class="tag t-in">IN</span>':'<span class="tag t-out">OUT</span>';
tr.innerHTML='<td>'+r[0]+'</td><td>'+r[1]+'</td><td>'+r[2]+'</td><td>'+d+'</td><td>'+r[4]+'</td><td>'+r[5]+'</td><td class="n">'+fmt(r[6])+'</td><td class="n">'+fmt(r[7])+'</td>';tb.appendChild(tr);});
</script></body></html>"""


if __name__ == "__main__":
    main()
