#!/bin/bash
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

ROOT="$CLAUDE_PROJECT_DIR"
cd "$ROOT"

# PDF 생성 파이프라인이 매번 새 컨테이너에서 필요로 하는 시스템/언어 의존성 설치.
# (LibreOffice의 docx->pdf 변환 필터(libreoffice-writer)와 qpdf는 기본 이미지에 없음 —
#  누락 시 soffice가 "source file could not be loaded"로 조용히 실패한다.)
apt-get update -qq
apt-get install -y -qq libreoffice-writer qpdf

pip install -q --break-system-packages -r requirements.txt

npm install --no-audit --no-fund
echo "export NODE_PATH=\"$ROOT/node_modules\"" >> "$CLAUDE_ENV_FILE"

mkdir -p ~/.fonts
cp assets/fonts/*.otf ~/.fonts/
fc-cache -f >/dev/null
