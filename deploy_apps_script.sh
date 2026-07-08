#!/usr/bin/env bash
# Publica o conteúdo de apps-script-atual/ direto no projeto Apps Script.
#
# Requisitos (variáveis de ambiente):
#   SCRIPT_ID                    — ID do projeto Apps Script (Configurações do projeto → ID do script)
#   CLOUDSDK_AUTH_ACCESS_TOKEN   — access token Google com escopo script.projects
#                                  (ex.: gcloud auth application-default print-access-token)
#
# A Apps Script API precisa estar ativada em https://script.google.com/home/usersettings
#
# ATENÇÃO: a API substitui TODO o conteúdo do projeto pelo enviado aqui.
set -euo pipefail

SRC_DIR="$(cd "$(dirname "$0")" && pwd)/apps-script-atual"
API="https://script.googleapis.com/v1/projects/${SCRIPT_ID:?defina SCRIPT_ID}/content"
TOKEN="${CLOUDSDK_AUTH_ACCESS_TOKEN:?defina CLOUDSDK_AUTH_ACCESS_TOKEN}"

# 1) Valida acesso e mostra o estado atual do projeto.
echo "== Verificando acesso ao projeto $SCRIPT_ID =="
HTTP=$(curl -sS -o /tmp/as_current.json -w "%{http_code}" -H "Authorization: Bearer $TOKEN" "$API")
if [ "$HTTP" != "200" ]; then
  echo "ERRO: GET retornou $HTTP — token inválido/expirado ou API desativada." >&2
  cat /tmp/as_current.json >&2
  exit 1
fi
python3 - "$SRC_DIR" <<'EOF'
import json, sys
d = json.load(open('/tmp/as_current.json'))
print("Projeto atual tem %d arquivos." % len(d.get('files', [])))
EOF

# 2) Monta o payload com todos os arquivos locais e publica.
echo "== Publicando arquivos de $SRC_DIR =="
python3 - "$SRC_DIR" <<'EOF' > /tmp/as_payload.json
import json, os, sys
src = sys.argv[1]
files = []
for name in sorted(os.listdir(src)):
    path = os.path.join(src, name)
    if not os.path.isfile(path):
        continue
    with open(path, encoding='utf-8') as fh:
        source = fh.read()
    if name == 'appsscript.json':
        files.append({'name': 'appsscript', 'type': 'JSON', 'source': source})
    elif name.endswith('.js') or name.endswith('.gs'):
        files.append({'name': os.path.splitext(name)[0], 'type': 'SERVER_JS', 'source': source})
print(json.dumps({'files': files}))
EOF

HTTP=$(curl -sS -o /tmp/as_result.json -w "%{http_code}" -X PUT \
  -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  --data-binary @/tmp/as_payload.json "$API")
if [ "$HTTP" != "200" ]; then
  echo "ERRO: PUT retornou $HTTP." >&2
  cat /tmp/as_result.json >&2
  exit 1
fi
python3 - <<'EOF'
import json
d = json.load(open('/tmp/as_result.json'))
print("OK: projeto atualizado com %d arquivos." % len(d.get('files', [])))
for f in d.get('files', []):
    print(" -", f['name'], "(%s)" % f['type'])
EOF
