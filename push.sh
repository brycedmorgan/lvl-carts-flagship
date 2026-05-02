#!/usr/bin/env bash
# One-shot: initialize, commit, and push the LVL Carts demo to GitHub.
# Run from inside /Users/bryce/Documents/intermountaingolfcars/lvl_demo
set -e

cd "$(dirname "$0")"

# Restore the assets/ folder layout that index.html expects
mkdir -p assets
mv -f *.png assets/ 2>/dev/null || true
for f in cart-pearl-white.jpg cart-tiffany-blue.jpg cart-royal-white.jpg cart-atlas-6-blue.jpg hero1.jpg hero2.jpg hero3.jpg lvl-logo.png; do
  [ -f "$f" ] && mv -f "$f" assets/ 2>/dev/null || true
done

# Re-prefix asset references in index.html (in case they were flattened)
# Only re-add the assets/ prefix if it's not already there.
python3 - <<'PY'
import re, pathlib
p = pathlib.Path("index.html")
s = p.read_text()
# Match src="X.jpg|png" or href="X.png" without assets/ prefix
def fix(m):
    quote, name = m.group(1), m.group(2)
    if name.startswith(("http","data:","#","/","assets/")): return m.group(0)
    return f'{m.group(0).split("=")[0]}={quote}assets/{name}{quote}'
s = re.sub(r'(?:src|href)=(["\'])([^"\']+\.(?:png|jpe?g|webp))\1', fix, s)
p.write_text(s)
print("paths fixed")
PY

git init -q 2>/dev/null || true
git checkout -B main 2>/dev/null || true
git add .
git -c user.email="brycedmorgan@gmail.com" -c user.name="Bryce Morgan" commit -m "LVL Carts digital flagship — initial demo" -q || echo "(nothing to commit)"
git remote remove origin 2>/dev/null || true
git remote add origin https://github.com/brycedmorgan/lvl-carts-flagship.git
git push -u origin main --force

echo ""
echo "✅ Pushed to https://github.com/brycedmorgan/lvl-carts-flagship"
echo "Now go to https://vercel.com/new — import the repo — deploy."
