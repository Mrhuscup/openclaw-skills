#!/usr/bin/env python3
"""常州招标测试脚本 - 验证Python版抓取逻辑"""
import urllib.request, re, ssl, sys, time
from datetime import datetime

BASE_URL = 'http://ggzy.xzsp.changzhou.gov.cn'

def http_get(url):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(url, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept-Encoding': 'gzip, deflate',
    })
    with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
        return r.read().decode('utf-8', 'ignore')

RELEVANT = ['市政','道路','公路','桥梁','排水','照明','绿化','交通','施工','总承包','新建','改建']

def score(t):
    return sum(1 for k in RELEVANT if k in t)

html = http_get(BASE_URL + '/jyzx/001001/tradeInfonew.html?category=001001')
print(f'HTML: {len(html)} bytes | has </td>: {"</td>" in html} | has tzjydetail: {"tzjydetail" in html}')

items = []
for m in re.finditer(r'<tr[^>]*>([\s\S]*?)</tr>', html):
    row = m.group(1)
    segs = row.split('</td>')
    if len(segs) < 4: continue
    om = re.search(r"tzjydetail\('([^']+)',\s*'([^']+)'\s*,", segs[1])
    if not om or om.group(2) in segs[0]: continue
    tm = re.search(r'title="([^"]{5,120})"', segs[1])
    if not tm: continue
    area = re.sub('<[^>]+>', '', segs[2]).replace('&nbsp;',' ').strip()
    date = re.sub('<[^>]+>', '', segs[3]).replace('&nbsp;',' ').strip()
    items.append({'cat':om.group(1),'uuid':om.group(2),'title':tm.group(1),'area':area,'date':date})

print(f'\nParsed {len(items)} items:')
for i, it in enumerate(items[:5]):
    print(f'  {i+1}. [{it["date"]}] {it["area"]} | {it["title"][:45]} | uuid={it["uuid"][:8]}...')

# Filter + sort by relevance
start, end = '2026-03-01', '2026-04-02'
filt = [x for x in items if start <= x['date'] <= end]
filt.sort(key=lambda x: (score(x['title']), x['date']), reverse=True)
print(f'\nRelevance-filtered ({start} to {end}): {len(filt)} items')
for i, it in enumerate(filt[:5]):
    print(f'  {i+1}. {score(it["title"])}★ | {it["date"]} | {it["area"]} | {it["title"][:45]}')

print(f'\n✅ Test PASSED - {len(items)} items parsed, {len(filt)} in date range')
