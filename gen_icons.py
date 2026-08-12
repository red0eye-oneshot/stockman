#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
이 파일을 C:\stock\apk\ 에 복사한 후
python gen_icons.py 실행하면
icon-192.png, icon-512.png, manifest.json 이 생성됩니다.
"""
import struct, zlib, math, json, os

def make_png(size):
    w, h = size, size
    cx, cy = w // 2, h // 2
    r = w * 0.38

    def pixel(x, y):
        dx, dy = x - cx, y - cy
        dist = math.sqrt(dx*dx + dy*dy)
        pr, pg, pb = 10, 10, 15  # 배경

        if dist < r:
            pr, pg, pb = 20, 20, 35  # 원 내부

        # 상승 추세선
        x1, y1 = cx - r*0.5, cy + r*0.28
        x2, y2 = cx + r*0.5, cy - r*0.28
        dx2 = x2 - x1; dy2 = y2 - y1
        t = max(0, min(1, ((x-x1)*dx2 + (y-y1)*dy2) / (dx2*dx2+dy2*dy2)))
        lx = x1 + t*dx2; ly = y1 + t*dy2
        if math.sqrt((x-lx)**2+(y-ly)**2) < max(2, w//28) and dist < r*0.85:
            pr, pg, pb = 255, 51, 85  # 빨간 선

        # 바 차트
        bars = [(-r*0.33, r*0.32), (-r*0.11, r*0.20), (r*0.11, r*0.07)]
        bw = r * 0.13
        base = cy + r * 0.38
        for bx_off, bh in bars:
            bx = cx + bx_off
            if abs(x - bx) < bw/2 and base - bh < y < base + 2:
                pr, pg, pb = 91, 79, 255  # 보라 막대

        return pr, pg, pb

    raw = b''
    for y in range(h):
        row = b'\x00'
        for x in range(w):
            r2, g2, b2 = pixel(x, y)
            row += bytes([r2, g2, b2, 255])
        raw += row

    comp = zlib.compress(raw, 9)

    def chunk(name, data):
        c = name + data
        return struct.pack('>I', len(data)) + c + struct.pack('>I', zlib.crc32(c) & 0xffffffff)

    ihdr = struct.pack('>IIBBBBB', w, h, 8, 6, 0, 0, 0)
    png = b'\x89PNG\r\n\x1a\n'
    png += chunk(b'IHDR', ihdr)
    png += chunk(b'IDAT', comp)
    png += chunk(b'IEND', b'')
    return png

manifest = {
    "name": "주식트래커",
    "short_name": "주식트래커",
    "description": "한국 주식 수익률 및 투자자 동향 트래커",
    "start_url": "/",
    "display": "standalone",
    "background_color": "#0a0a0f",
    "theme_color": "#0a0a0f",
    "orientation": "portrait",
    "icons": [
        {"src": "/icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any maskable"},
        {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"}
    ],
    "lang": "ko",
    "categories": ["finance"]
}

out_dir = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(out_dir, 'icon-192.png'), 'wb') as f:
    f.write(make_png(192))
print("icon-192.png 생성 완료")

with open(os.path.join(out_dir, 'icon-512.png'), 'wb') as f:
    f.write(make_png(512))
print("icon-512.png 생성 완료")

with open(os.path.join(out_dir, 'manifest.json'), 'w', encoding='utf-8') as f:
    json.dump(manifest, f, ensure_ascii=False, indent=2)
print("manifest.json 생성 완료")

print("\n완료! 이제 git add / commit / push 하세요.")
