#!/usr/bin/env python3
"""
reconstruct_ledger.py — авто-классификатор: строит ледж­ер-эквивалент SA Minerals
ПРЯМО из сырых выписок FNB (Asset Co + Claim Co), без участия Александра.

Назначение: страховка от bus-factor. Если ледж­ера Александра нет — этот скрипт
восстанавливает классификацию (Category, Internal?, Direction) из описаний банковских
транзакций по правилам, и считает те же метрики, что build_dashboards.py.
Неоднозначные операции (которые требуют бизнес-знания) он НЕ угадывает, а помечает
 uncategorized — чтобы человек посмотрел только их.

Запуск:  python3 reconstruct_ledger.py <папка с CSV-выписками FNB>
Формат CSV: строка-заголовок "...FOR ACCOUNT NUMBER <n>", затем
            Date,SERVICE FEE,Amount,DESCRIPTION,REFERENCE,Balance,...
"""
import sys, os, csv, glob, re

ACC = {"63209738939": "Asset Co", "63209745116": "Claim Co"}

def classify(desc, amt):
    """Возвращает (category, internal:bool). desc — описание из банка, amt — сумма (знак важен)."""
    d = desc.upper()
    # порядок важен — от конкретного к общему
    if "ADRU TECH" in d:
        return "Capital Injection", False
    if "ADT CASH DEPO" in d or ("DEPOSIT" in d and abs(amt) <= 5000):
        return "Account Activation", False
    if "DRAWDOWN" in d:
        return "Loan to EMS (Receivable)", False
    if any(k in d for k in ("INVESTEC", "BUBESI", "SANDVIK", "WILROCK")):
        return "Claim Acquisition", False
    if "SWIFT COMMISSION" in d or "INWARD SWIFT" in d or "FOREX TRANSFER" in d or "SWIFT CORRECTION" in d:
        return "Bank Charges", False
    # интеркомпани: перевод между двумя счетами группы
    if ("TRF" in d or "TRANSFER FUNDS" in d) and ("SA MINERALS" in d or "TRANSFER FUNDS" in d):
        return "Intercompany Transfer", True
    if any(k in d for k in ("AFRIQOM", "STRAND HOTEL", "ACCOMMODATION", "GARDEN COURT",
                            "INV17842", "INVOICE", "ADVISOR", "WOOD", "COLLECTIVE")):
        return "Operating Expense", False
    return "UNCATEGORIZED", False  # требует ручной проверки

def parse_csv(path):
    txt = open(path, encoding="utf-8", errors="replace").read().splitlines()
    acc_no = None
    for line in txt[:3]:
        mm = re.search(r"ACCOUNT NUMBER\s*([0-9]+)", line.upper())
        if mm:
            acc_no = mm.group(1); break
    account = ACC.get(acc_no, acc_no or "?")
    rows = []
    for r in csv.reader(txt):
        if len(r) < 4:
            continue
        try:
            amt = round(float(r[2]), 2)
        except (ValueError, IndexError):
            continue
        desc = r[3]
        cat, internal = classify(desc, amt)
        rows.append({"date": r[0].strip(), "account": account, "amount": amt,
                     "desc": desc.strip(), "category": cat, "internal": internal})
    return rows

def main():
    if len(sys.argv) < 2:
        print("usage: reconstruct_ledger.py <dir-with-fnb-csv>"); sys.exit(1)
    files = glob.glob(os.path.join(sys.argv[1], "*.csv"))
    rows = []
    for f in files:
        rows += parse_csv(f)
    rows.sort(key=lambda x: (x["date"], x["account"]))
    # метрики (как в build_dashboards)
    def s(cond): return sum(r["amount"] for r in rows if cond(r))
    injected = sum(r["amount"] for r in rows if r["amount"] > 0 and not r["internal"]
                   and r["category"] in ("Capital Injection", "Account Activation"))
    claims = -sum(r["amount"] for r in rows if r["category"] == "Claim Acquisition")
    loan = -sum(r["amount"] for r in rows if r["category"] == "Loan to EMS (Receivable)")
    opex = -sum(r["amount"] for r in rows if r["category"] == "Operating Expense")
    charges = -sum(r["amount"] for r in rows if r["category"] == "Bank Charges")
    recoverable = claims + loan
    overhead = opex + charges
    cash = injected - recoverable - overhead
    unc = [r for r in rows if r["category"] == "UNCATEGORIZED"]
    print(f"Транзакций распознано: {len(rows)}  (счета: {sorted(set(r['account'] for r in rows))})")
    print(f"  Capital injected : R{injected:,.2f}")
    print(f"  Recoverable      : R{recoverable:,.2f}  (claims {claims:,.0f} + loan {loan:,.0f})")
    print(f"  Overhead         : R{overhead:,.2f}")
    print(f"  Cash (derived)   : R{cash:,.2f}")
    print(f"  Preserved        : {(recoverable+cash)/injected*100:.1f}%" if injected else "  Preserved: n/a")
    print(f"\nКатегории (авто):")
    from collections import Counter
    for k, v in Counter(r["category"] for r in rows).most_common():
        print(f"  {k:28} {v}")
    if unc:
        print(f"\n⚠ ТРЕБУЮТ РУЧНОЙ ПРОВЕРКИ ({len(unc)}) — бизнес-смысл не выводится из текста:")
        for r in unc:
            print(f"   {r['date']} {r['account']} {r['amount']:>14,.2f}  «{r['desc'][:48]}»")

if __name__ == "__main__":
    main()
