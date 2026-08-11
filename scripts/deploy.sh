#!/usr/bin/env bash
# Despliega la app de facturación a Google Apps Script con clasp.
# Uso:
#   ./scripts/deploy.sh                 # primer despliegue (crea proyecto nuevo)
#   ./scripts/deploy.sh <scriptId>      # despliega a un proyecto existente
#   ./scripts/deploy.sh --no-tests      # omite runTests
set -euo pipefail

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

say()  { printf "${GREEN}==>${NC} %s\n" "$1"; }
warn() { printf "${YELLOW}!! %s${NC}\n" "$1"; }
err()  { printf "${RED}XX %s${NC}\n" "$1"; exit 1; }

RUN_TESTS=1
SCRIPT_ID=""
for arg in "$@"; do
  case "$arg" in
    --no-tests) RUN_TESTS=0 ;;
    --help|-h)
      echo "Despliega la app a Google Apps Script usando clasp."
      echo "  ./scripts/deploy.sh                crea un proyecto nuevo + despliega"
      echo "  ./scripts/deploy.sh <scriptId>     despliega a un proyecto existente"
      echo "  ./scripts/deploy.sh --no-tests     omite ejecutar runTests"
      exit 0 ;;
    *) SCRIPT_ID="$arg" ;;
  esac
done

command -v clasp >/dev/null 2>&1 || {
  err "clasp no está instalado. Instálalo: npm install -g @google/clasp"
}

command -v jq >/dev/null 2>&1 || {
  err "jq no está instalado. Instálalo (ej. sudo apt install jq)."
}

say "Verificando login de clasp..."
if ! clasp whoami >/dev/null 2>&1; then
  err "No hay sesión de clasp. Ejecuta primero: clasp login"
fi

say "Verificando archivos del proyecto..."
for f in appsscript.json Code.gs Data.gs Aux.gs index.html pdfTemplate.html; do
  [ -f "$f" ] || err "Falta el archivo $f en el directorio actual."
done

if [ -f .clasp.json ] && [ -z "$SCRIPT_ID" ]; then
  warn "Ya existe .clasp.json. Usando el scriptId configurado."
elif [ -z "$SCRIPT_ID" ]; then
  say "Creando proyecto Apps Script vinculado a un nuevo Google Spreadsheet..."
  clasp create --type sheets --title "Facturación" --rootDir . >/dev/null
else
  say "Vinculando al proyecto existente: $SCRIPT_ID"
  clasp clone "$SCRIPT_ID" --rootDir . >/dev/null
fi

say "Subiendo archivos (clasp push)..."
clasp push -f

if [ "$RUN_TESTS" = "1" ]; then
  say "Ejecutando tests (clasp run runTests)..."
  warn "La primera vez, autoriza el acceso cuando lo pida."
  clasp run runTests
else
  warn "Tests omitidos (--no-tests)."
fi

say "Creando despliegue de Web App..."
DEPLOY_OUT=$(clasp deploy --type web --description "Facturación (despliegue manual)")
echo "$DEPLOY_OUT"

URL=$(echo "$DEPLOY_OUT" | grep -oE 'https://[^ ]+' || true)
if [ -n "$URL" ]; then
  say "App lista: $URL"
else
  warn "No se pudo extraer la URL. Revisa el editor: Implementar → Administrar implementaciones."
fi

say "Recuerda: acceso 'Solo yo' en el editor (Implementar → Nueva implementación → Web App)."
say "Antes de usarla, corre ensureSheets en el editor (o se crea solo en el primer uso)."
