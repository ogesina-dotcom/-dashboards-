#!/usr/bin/env bash
# weekly_build.sh — единый еженедельный пайплайн: пересобирает все три дашборда
# (инвест + сводный + EMS-операции) и пушит одним коммитом в GitHub Pages.
# Запускать в окружении, где локально доступны файлы Drive (Drive for Desktop / выгрузки).
#
#   chmod +x weekly_build.sh
#   ./weekly_build.sh
#
# Для авто-запуска по расписанию — cron (пример в конце файла).

set -uo pipefail

# ================== ПРАВЬ ЭТИ ПУТИ ==================
REPO="$HOME/dashboards/-dashboards-"                       # локальный клон GitHub-репозитория
RAW_FNB="$HOME/GoogleDrive/Weekly finance/01_SA_Minerals/Bank_raw"   # выписки FNB Asset+Claim (.csv/.zip)
EMS_DATA="$HOME/GoogleDrive/Weekly finance/02_EMS/SmartBuilder"      # SmartBuilder xlsx + ems_balances.json + ems_budget.json
EMS_BAL="$EMS_DATA/ems_balances.json"                      # балансы банков EMS (для сводного)

# Куда писать EMS-дашборд. По умолчанию отдельный файл, чтобы НЕ затирать
# текущий рукотворный ems_financial_live.html. Когда сверишь вывод генератора —
# поменяй на "ems_financial_live.html".
OUT_EMS="ems_ops_auto.html"

# Токен для пуша (или настрой git credential helper и оставь пустым).
GITHUB_TOKEN="${GITHUB_TOKEN:-}"
REMOTE="https://github.com/ogesina-dotcom/-dashboards-.git"
# ===================================================

cd "$REPO" || { echo "Нет папки репозитория: $REPO"; exit 1; }
git pull -q || true

# 1) Инвест + сводный (из выписок FNB). Стоп-гейт внутри скрипта:
#    крупная неклассифицированная операция → exit!=0 → НЕ публикуем.
python3 build_dashboards.py --raw "$RAW_FNB" \
  --out sa_minerals_group_dashboard.html \
  --consolidated-out consolidated_overview.html --ems-balances "$EMS_BAL" \
  --report flagged.txt --metrics-out metrics_latest.csv
INV_RC=$?
if [ $INV_RC -ne 0 ]; then
  echo "BLOCK: инвест/сводный не собрались (код $INV_RC) — проверь flagged.txt. Публикация остановлена."
  exit $INV_RC
fi

# 2) EMS-операции (из SmartBuilder)
python3 build_ems_dashboard.py --data "$EMS_DATA" --out "$OUT_EMS" || {
  echo "WARN: EMS-дашборд не собрался — публикую инвест/сводный без него."; }

# 3) Коммит + пуш, только если что-то поменялось
DATE=$(date +%F)
git add sa_minerals_group_dashboard.html consolidated_overview.html "$OUT_EMS" flagged.txt metrics_latest.csv 2>/dev/null
if git diff --cached --quiet; then
  echo "Изменений нет — коммит не нужен."
  exit 0
fi
git commit -q -m "Weekly rebuild $DATE (invest + consolidated + EMS)"
if [ -n "$GITHUB_TOKEN" ]; then
  git push -q "https://x-access-token:${GITHUB_TOKEN}@github.com/ogesina-dotcom/-dashboards-.git" HEAD:main
else
  git push -q
fi
echo "OK: запушен weekly rebuild $DATE. announceIfUpdatedToday() увидит сегодняшний коммит."

# ── cron (пятница 17:00, за час до анонса 18:00) ──
#   crontab -e
#   0 17 * * 5  GITHUB_TOKEN=github_pat_xxx /path/to/weekly_build.sh >> /path/to/weekly_build.log 2>&1
