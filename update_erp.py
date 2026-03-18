#!/usr/bin/env python3
"""
ISKA Metal ERP - Otomatik Güncelleme Scripti
GitHub Actions tarafından çalıştırılır.
data/ klasörüne Excel yüklenince otomatik çalışır.
"""

import pandas as pd
import glob, os, json, re
from datetime import datetime

TODAY = datetime.now()

PRICE_KEYS = ['FİYAT','FIYAT','PRICE','UNIT PRICE','BİRİM FİYAT','BIRIM FIYAT']
QTY_KEYS   = ['ADET','QUANTITY','QUANTITIY','QTY','MİKTAR','MIKTAR']
DONE_VALS  = ['GÖNDERİLDİ','GONDERILDI','SHIPPED','TAMAMLANDI','COMPLETED','SEVK EDİLDİ']
SHIP_KEYS  = ['SHIPMENT','SEVKİYAT','SEVKIYAT','DELIVERY','TESLİMAT']

MUSTERI_MAP = {
    'CIVAK':'CIVAK','ADITEG':'ADITEG','AL FILTER':'AL FILTER',
    'AV INDUSTRIAL':'AV INDUSTRIAL','C-FLEX':'C-FLEX','C.C.RUBBER':'C-FLEX',
    'FAI FILTRI':'FAI FILTRI','FRANKSA':'FRANKSA','GETECH SRL':'GETECH',
    'GUFERO':'GUFERO','GUMMI TECHNIK':'GUMMI TECHNIK',
    'HABERKORN':'HABERKORN','ITH GMBH':'ITH GMBH','JAGER':'JAGER',
    'LA MERCE':'LA MERCE','LLC MACHINE':'LLC MACHINE',
    'MBM FILTER':'MBM FILTER','MENGIA SA':'MENGIA','NDR SRL':'NDR SRL',
    'ORMANT SRL':'ORMANT','OTTO MAIERHOFER':'OTTO MAIERHOFER',
    'PANTECNICA SPA':'PANTECNICA','SARL SAFI':'SARL SAFI','SATTLER':'SATTLER',
    'SCHUHMACHER':'SCHUHMACHER','STÖFFL':'STÖFFL','STO FFL':'STÖFFL',
    'TAROGOMMA SRL':'TAROGOMMA','TENAX RUBBER':'TENAX RUBBER',
    'TRIA 2000':'TRIA 2000','PHILLIPS':'PHILLIPS'
}

def get_musteri(fpath):
    base = os.path.basename(fpath).upper()
    for suffix in [' SİPARİŞ 2026.XLSX',' SİPARİŞLERİ 2026.XLSX',' - 2026.XLSX','.XLSX']:
        base = base.replace(suffix,'')
    base = base.strip()
    for k,v in MUSTERI_MAP.items():
        if k.upper() in base: return v
    return base.title()

def parse_ship_date(text):
    import re as re2
    m = re2.search(r'(\d{1,2})[./](\d{1,2})[./](\d{4})', text)
    if m:
        try: return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except: pass
    months = {'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,
              'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12,
              'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'JUNE':6,
              'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    for k,v in months.items():
        if k in text.upper():
            yr = re2.search(r'20\d\d', text)
            return datetime(int(yr.group()) if yr else TODAY.year, v, 1)
    return None

def parse_orders(files):
    all_orders = []
    order_id = 1
    for fpath in sorted(files):
        musteri = get_musteri(fpath)
        try:
            raw = pd.read_excel(fpath, header=None)
            header_row = None
            for i, row in raw.iterrows():
                vals = [str(v).strip().upper() for v in row.values]
                if any(v in PRICE_KEYS or v in QTY_KEYS for v in vals):
                    header_row = i; break
            if header_row is None: continue

            raw.columns = [str(c).strip().upper() for c in raw.iloc[header_row]]
            df = raw.iloc[header_row+1:].reset_index(drop=True)
            cols = list(df.columns)

            price_col  = next((c for c in cols if c in PRICE_KEYS), None)
            qty_col    = next((c for c in cols if c in QTY_KEYS), None)
            status_col = next((c for c in cols if any(k in c for k in ['DURUM','STATUS','STATÜ'])), None)
            netsis_col = next((c for c in cols if 'NETSIS' in c), None)
            iska_col   = next((c for c in cols if 'İSKA' in c or 'ISKA' in c), None)
            urun_col   = next((c for c in cols if 'AÇILIM' in c or 'DESC' in c or 'ÜRÜN' in c), None)

            current_date = None
            current_durum = 'aktif'

            for _, row in df.iterrows():
                netsis = str(row.get(netsis_col,'') if netsis_col else '').strip()
                if not netsis or netsis.upper() == 'NAN': continue
                if any(k in netsis.upper() for k in SHIP_KEYS):
                    sd = parse_ship_date(netsis)
                    if sd:
                        current_date = sd
                        current_durum = 'gonderildi' if sd <= TODAY else 'aktif'
                    continue

                durum = current_durum
                if status_col:
                    st = str(row.get(status_col,'')).strip().upper()
                    if st in DONE_VALS: durum = 'gonderildi'
                    elif st and st != 'NAN': durum = 'aktif'

                p = float(pd.to_numeric(row.get(price_col,0) if price_col else 0, errors='coerce') or 0)
                q = float(pd.to_numeric(row.get(qty_col,0) if qty_col else 0, errors='coerce') or 0)
                iska = str(row.get(iska_col,'') if iska_col else '').strip()
                urun = str(row.get(urun_col,'') if urun_col else '').strip()
                if iska == 'nan': iska = ''
                if urun == 'nan': urun = ''
                tarih = current_date.strftime('%Y-%m-%d') if current_date else ''

                all_orders.append({
                    'i':order_id,'n':netsis,'m':musteri,
                    'k':iska,'u':urun[:60],'t':'',
                    'f':round(p,4),'a':int(round(float(q))) if not pd.isna(q) else 0,'d':tarih,'r':durum
                })
                order_id += 1
        except Exception as e:
            print(f"❌ {musteri}: {e}")
    return all_orders

def parse_depo(fpath):
    df = pd.read_excel(fpath, header=None)
    # Header satır 3'te sabit
    df.columns = [str(c).strip().upper() for c in df.iloc[3]]
    df = df.iloc[4:].reset_index(drop=True)
    
    kod_col = next((c for c in df.columns if 'NETSİS' in c or 'NETSIS' in c), None)
    mik_col = next((c for c in df.columns if 'KALAN' in c), None)
    raf_col = next((c for c in df.columns if 'RAF' in c), None)
    ad_col  = df.columns[0]

    depo = []
    for _, row in df.iterrows():
        kod = str(row.get(kod_col,'') if kod_col else '').strip()
        if not kod or kod.upper() == 'NAN': continue
        ad  = str(row.get(ad_col,'') if ad_col else '').strip()
        mik = float(pd.to_numeric(row.get(mik_col,0) if mik_col else 0, errors='coerce') or 0)
        raf = str(row.get(raf_col,'') if raf_col else '').strip()
        if ad == 'nan': ad = ''
        if raf == 'nan': raf = ''
        depo.append({'k':kod,'a':ad,'m':int(mik) if not pd.isna(mik) else 0,'r':raf,'c':''})
    return depo

# ── ANA İŞLEM ──────────────────────────────────────────────
print("📦 ISKA Metal ERP - Otomatik Güncelleme")
print(f"📅 Tarih: {TODAY.strftime('%d.%m.%Y %H:%M')}")

# Sipariş dosyaları
order_files = glob.glob('data/siparisler/*.xlsx')
print(f"\n📋 {len(order_files)} sipariş dosyası bulundu")
orders = parse_orders(order_files)
aktif = [x for x in orders if x['r']=='aktif']
print(f"✅ Toplam: {len(orders)} | Aktif: {len(aktif)} | €: {sum(x['f']*x['a'] for x in aktif):,.2f}")

# Depo dosyası
depo_files = glob.glob('data/depo/*.xlsx')
depo = []
if depo_files:
    depo_files.sort(key=os.path.getmtime, reverse=True)
    depo = parse_depo(depo_files[0])
    print(f"🏭 Depo: {len(depo)} ürün ({os.path.basename(depo_files[0])})")

# Mevcut index.html'deki veriyi güncelle
import re as re_mod

with open('index.html','r',encoding='utf-8') as f:
    html = f.read()

html = re_mod.sub(r'var DEPO=(?:\[.*?\]|__DEPO__);', 'var DEPO='+json.dumps(depo, ensure_ascii=False)+';', html, flags=re_mod.DOTALL)
html = re_mod.sub(r'var SIPARISLER=(?:\[.*?\]|__SIPARISLER__);', 'var SIPARISLER='+json.dumps(orders, ensure_ascii=False)+';', html, flags=re_mod.DOTALL)

with open('index.html','w',encoding='utf-8') as f:
    f.write(html)

print(f"\n✅ index.html güncellendi ({len(html)//1024} KB)")


# ── SUPABASE GÜNCELLE ──────────────────────────────────────
import urllib.request, json as json_mod

SB_URL = 'https://wtbvmagacbbmuwfxqrnj.supabase.co'
SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind0YnZtYWdhY2JibXV3Znhxcm5qIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMyNzYxOTcsImV4cCI6MjA4ODg1MjE5N30.5Qr_7D10N77fUJYORTyswLBgYPxWFxBKQBlh0wojleU'

def sb_request(path, method='GET', data=None):
    url = SB_URL + '/rest/v1/' + path
    headers = {
        'apikey': SB_KEY,
        'Authorization': 'Bearer ' + SB_KEY,
        'Content-Type': 'application/json',
        'Prefer': 'resolution=merge-duplicates'
    }
    body = json_mod.dumps(data).encode('utf-8') if data else None
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        urllib.request.urlopen(req, timeout=15)
        return True
    except Exception as e:
        print(f"  Supabase hata: {e}")
        return False

if depo:
    print("\n☁️  Supabase stok güncelleniyor...")
    batch_size = 50
    success = 0
    for i in range(0, len(depo), batch_size):
        batch = depo[i:i+batch_size]
        rows = [{'kod':d['k'],'ad':d['a'],'miktar':d['m'],'raf':d['r'],'kategori':d['c']} for d in batch]
        if sb_request('stok', 'POST', rows):
            success += len(batch)
    print(f"  ✅ {success}/{len(depo)} ürün Supabase'e aktarıldı")
