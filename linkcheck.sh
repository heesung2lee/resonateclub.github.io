#!/bin/bash
# linkcheck.sh — 인포사이트 내부링크+이미지 무결성 검사 (유닛테스트 아님, 산출물 검사)
# 사용: ./linkcheck.sh — 깨진 내부링크/없는 이미지 있으면 FAIL 목록 출력
cd "$(dirname "$0")"
/usr/bin/python3 - <<'PYEOF'
import re, os, html
from pathlib import Path
root = Path('.')
pages = sorted(root.glob('*.html')) + sorted((root/'articles').glob('*.html'))
href_re = re.compile(r'''(?:href|src)\s*=\s*["']([^"'#]+)''', re.I)
broken, missing_img, checked = [], [], 0
for p in pages:
    t = p.read_text(encoding='utf-8', errors='ignore')
    basedir = p.parent
    for m in href_re.finditer(t):
        u = html.unescape(m.group(1)).strip()
        if not u or u.startswith(('http', 'mailto:', 'tel:', 'data:', 'javascript:', '#')):
            continue
        u = u.split('#')[0].split('?')[0]
        target = (basedir/u) if not u.startswith('/') else (root/u[1:])
        checked += 1
        if not target.exists():
            is_img = target.suffix.lower() in ('.jpg','.jpeg','.png','.webp','.gif','.svg','.ico','.css','.js')
            (missing_img if is_img else broken).append(f'{p}: {u}')
print(f'checked {checked} refs in {len(pages)} pages')
if broken:
    print(f'FAIL broken_links={len(broken)}:'); [print('  '+b) for b in broken[:20]]
if missing_img:
    print(f'FAIL missing_files={len(missing_img)}:'); [print('  '+b) for b in missing_img[:20]]
print('LINKCHECK FAILED' if (broken or missing_img) else 'LINKCHECK PASSED')
PYEOF
