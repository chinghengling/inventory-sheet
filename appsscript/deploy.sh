#!/usr/bin/env bash
# 推送並重新部署 Apps Script 後端，部署ID固定，SCRIPT_URL/CAN_SCRIPT_URL 不會變動。
# 用法: ./deploy.sh inventory | can | all
set -euo pipefail
cd "$(dirname "$0")"

INVENTORY_DEPLOYMENT_ID="AKfycbw-B0NMCe5bdKrBEfAZGCIHNeLEWWXo1jlfloCjs7JQeg4mXvgkEQzr95r0dnm7wW0dxA"
CAN_DEPLOYMENT_ID="AKfycbwPy1PY5s4j8BSCRNj3m0rRko4EVQDdqiKhJ7iIa-jFQSKdGxBZVUvRbUs9e_Px9ti0"

deploy_one() {
  local dir="$1" deployment_id="$2" label="$3"
  echo "== $label =="
  (cd "$dir" && clasp push -f && clasp deploy -i "$deployment_id" -d "auto deploy $(date '+%Y-%m-%d %H:%M:%S')")
}

case "${1:-all}" in
  inventory) deploy_one inventory "$INVENTORY_DEPLOYMENT_ID" "日盤表後端" ;;
  can) deploy_one can "$CAN_DEPLOYMENT_ID" "易開罐後端" ;;
  all)
    deploy_one inventory "$INVENTORY_DEPLOYMENT_ID" "日盤表後端"
    deploy_one can "$CAN_DEPLOYMENT_ID" "易開罐後端"
    ;;
  *) echo "用法: $0 inventory|can|all" >&2; exit 1 ;;
esac
