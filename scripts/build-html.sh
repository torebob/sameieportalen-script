set -euo pipefail
rm -rf dist
mkdir -p dist

# Bygg hver side eksplisitt til dist/
posthtml -i src/pages/index.html   -o dist/index.html   -c posthtml.config.js
posthtml -i src/pages/budget.html  -o dist/budget.html  -c posthtml.config.js
posthtml -i src/pages/reports.html -o dist/reports.html -c posthtml.config.js

# Kopiér statiske ressurser
mkdir -p dist/assets dist/js
cp -r src/assets/* dist/assets/ 2>/dev/null || true
cp -r src/js/*     dist/js/     2>/dev/null || true

echo "==> Bygd filer i dist/:"
ls -la dist
