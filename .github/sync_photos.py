"""Download each item's Drive photo (sheet column M) to images/<driveId>.jpg if it isn't there yet."""
import json, os, re, subprocess

items = json.load(open('/tmp/items.json'))['items']
ids = set()
for it in items:
    m = re.search(r'(?:/d/|id=)([\w-]{20,})', it.get('photo') or '')
    if m:
        ids.add(m.group(1))

os.makedirs('images', exist_ok=True)
for i in sorted(ids):
    out = f'images/{i}.jpg'
    if os.path.exists(out):
        continue
    url = f'https://drive.google.com/thumbnail?id={i}&sz=w1000'
    r = subprocess.run(['curl', '-sSL', '-f', '-o', out, url])
    ok = (r.returncode == 0 and os.path.exists(out) and os.path.getsize(out) > 2000
          and open(out, 'rb').read(3) in (b'\xff\xd8\xff', b'\x89PN'))
    if ok:
        from PIL import Image, ImageOps   # shrink for phones: max 800px wide, JPEG q80
        im = ImageOps.exif_transpose(Image.open(out)).convert('RGB'); im.thumbnail((800, 1067))
        im.save(out, 'JPEG', quality=80, optimize=True, progressive=True)
        print('saved', out, os.path.getsize(out))
    else:
        print('failed', i)
        if os.path.exists(out):
            os.remove(out)
