#!/usr/bin/env python
# coding: utf-8

# In[1]:


# KAP'tan indirilen dosyanın konumunu sistematikleştirmek için; 
import os

# Modüllerin kurulumunda doğru Python ortamının seçilmesi için; 
import sys

# Projenin sürdürülebilir olması amacıyla bilgisayarda olmayan modülün komut satırından oto kurulumu için; 
import subprocess

# İndirdiğimiz veriyi Python'da okuyabilmek ve finansal oran hesaplayabilmek için;
import pandas as pd


# In[2]:


# Okutacağımız verinin dosya yolunu tanımladık;
dosya_yolu = r"C:\Users\HP\Desktop\KAP VERİLERİ\ASELS_1395801_2024_4.xls"


# In[3]:


# KAP'tan temin edilen .xls uzantılı ancak içeriği HTML olan dosyayı okutmak ve liste halinde dfs değişkenine atanması için;
dfs = pd.read_html(dosya_yolu)


# In[4]:


# KAP' tan temin edilen dosya içerisinde bilanço ve gelir tablosu alanlarının otomatik bulunabilmesi için; 

# Başlangıçta None olarak tanımlıyoruz
df_bilanco = None
df_gelir = None


# In[5]:


# Tüm tabloları sırayla gez;
for i, df in enumerate(dfs):
    try:
        # Tablonun ilk birkaç satırını düz metne çevir ve küçük harfe indir
        metin = " ".join(df.astype(str).head(5).values.flatten()).lower()
        
        # Bilanço tespiti
        if "bilanço" in metin and df_bilanco is None:
            df_bilanco = df
            bilanco_index = i
        
        # Gelir tablosu tespiti
        if "gelir tablosu" in metin and df_gelir is None:
            df_gelir = df
            gelir_index = i

        # İkisini de bulduysa döngüyü durdur, tekrar eden isim olduğu durumunda ilkini seçmek için;
        if df_bilanco is not None and df_gelir is not None:
            break

    except:
        continue
# Check        
print(f"✅ Bilanço tablosu: dfs[{bilanco_index}]")
print(f"✅ Gelir tablosu: dfs[{gelir_index}]")


# In[6]:


# Bilanço ve gelir tablolarını bulduk, şimdi tablolar içerisindeki değerlere 
#kalem isimleri ile erişebilmek için tabloları "veriler" sözlüğüne dönüştüreceğiz ;

# 1. Bilanço ve gelir tablosundan doğru sütunları seç (Hesap Adı, Değer)
bilanco_clean = df_bilanco[[1, 3]].dropna().reset_index(drop=True)
gelir_clean = df_gelir[[1, 3]].dropna().reset_index(drop=True)

# 2. Her tabloyu sözlük yapısına çevir
sozluk_bilanco = dict(zip(bilanco_clean[1], bilanco_clean[3]))
sozluk_gelir = dict(zip(gelir_clean[1], gelir_clean[3]))

# 3. İki sözlüğü birleştir
veriler = {**sozluk_bilanco, **sozluk_gelir}

# 4. Sayıları Türkçe formatından temizle (1.234.567,89 → 1234567.89)
def temizle_sayi(x):
    if isinstance(x, str):
        return float(x.replace(".", "").replace(",", ".").replace("−", "-").replace(" ", ""))
    return x

for k in veriler:
    try:
        veriler[k] = temizle_sayi(veriler[k])
    except:
        continue
        
# 5. Sözlük örneği (ilk 10 anahtarı ve değeri göster)
veriler_ornek = {k: veriler[k] for k in list(veriler)[:10]}
veriler_ornek


# In[7]:


# Tablodaki tüm kalem isimlerini görmek için; 
for k in veriler:
    print(k)


# In[8]:


# Rasyo hesaplamaları için verileri al ve her oranı try-except ile güvenli şekilde hesapla;

sonuclar = {}

try:
    sonuclar["Cari Oran"] = veriler["TOPLAM DÖNEN VARLIKLAR"] / veriler["TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER"]
except:
    sonuclar["Cari Oran"] = "Hesaplanamadı"

try:
    sonuclar["Asit Test Oranı"] = (veriler["TOPLAM DÖNEN VARLIKLAR"] - veriler["Stoklar"]) / veriler["TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER"]
except:
    sonuclar["Asit Test Oranı"] = "Hesaplanamadı"
    
try:
    alacak_kalemi = (
        veriler.get("İlişkili Olmayan Taraflardan Ticari Alacaklar")
        or veriler.get("Ticari Alacaklar")
        or veriler.get("TOPLAM DÖNEN VARLIKLAR")
    )
    if alacak_kalemi is None or alacak_kalemi == 0:
        raise ValueError("Geçerli alacak verisi yok")

    alacak_devir_hizi = veriler["Hasılat"] / alacak_kalemi
    sonuclar["Alacak Devir Hızı"] = alacak_devir_hizi
    sonuclar["Alacak Tahsil Süresi"] = 365 / alacak_devir_hizi

except:
    sonuclar["Alacak Devir Hızı"] = "Hesaplanamadı"
    sonuclar["Alacak Tahsil Süresi"] = "Hesaplanamadı"

try:
    stok_devir_hizi = abs(veriler["Satışların Maliyeti"]) / veriler["Stoklar"]
    sonuclar["Stok Devir Hızı"] = stok_devir_hizi
    sonuclar["Stok Devir Süresi"] = 365 / stok_devir_hizi
except:
    sonuclar["Stok Devir Hızı"] = "Hesaplanamadı"
    sonuclar["Stok Devir Süresi"] = "Hesaplanamadı"

try:
    borc_kalemi = (
        veriler.get("İlişkili Olmayan Taraflara Ticari Borçlar")
        or veriler.get("Ticari Borçlar")
        or veriler.get("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER")
    )
    if borc_kalemi is None or borc_kalemi == 0:
        raise ValueError("Geçerli borç verisi yok")

    borc_devir_hizi = abs(veriler["Satışların Maliyeti"]) / borc_kalemi
    sonuclar["Borç Devir Hızı"] = borc_devir_hizi
    sonuclar["Borç Ödeme Süresi"] = 365 / borc_devir_hizi
except:
    sonuclar["Borç Devir Hızı"] = "Hesaplanamadı"
    sonuclar["Borç Ödeme Süresi"] = "Hesaplanamadı"

try:
    sonuclar["Faaliyet Süresi"] = sonuclar["Alacak Tahsil Süresi"] + sonuclar["Stok Devir Süresi"]
    sonuclar["Finansman Süresi"] = sonuclar["Faaliyet Süresi"] - sonuclar["Borç Ödeme Süresi"]
except:
    sonuclar["Faaliyet Süresi"] = "Hesaplanamadı"
    sonuclar["Finansman Süresi"] = "Hesaplanamadı"

try:
    sonuclar["Kaldıraç Oranı"] = veriler["TOPLAM YÜKÜMLÜLÜKLER"] / veriler["TOPLAM VARLIKLAR"]
except:
    sonuclar["Kaldıraç Oranı"] = "Hesaplanamadı"

try:
    sonuclar["Öz Kaynak Oranı"] = veriler["TOPLAM ÖZKAYNAKLAR"] / veriler["TOPLAM VARLIKLAR"]
except:
    sonuclar["Öz Kaynak Oranı"] = "Hesaplanamadı"

try:
    sonuclar["Aktif Karlılık"] = veriler["Net Dönem Karı veya Zararı"] / veriler["TOPLAM VARLIKLAR"]
except:
    sonuclar["Aktif Karlılık"] = "Hesaplanamadı"

try:
    sonuclar["Ekonomik Rantabilite"] = (veriler["Net Dönem Karı veya Zararı"] +- veriler["Finansman Giderleri"]) / veriler["TOPLAM KAYNAKLAR"]
except:
    sonuclar["Ekonomik Rantabilite"] = "Hesaplanamadı"

try:
    # Alternatif anahtarlarla verileri güvenli şekilde al
    net_kar = veriler.get("Net Dönem Karı veya Zararı") or veriler.get("DÖNEM KARI (ZARARI)")
    finansman = veriler.get("Finansman Giderleri")
    vergi = veriler.get("Dönem Vergi (Gideri) Geliri") or veriler.get("Vergi Gideri")
    amortisman = veriler.get("Amortisman ve İtfa Gideri İle İlgili Düzeltmeler") or 0  # Eksikse 0
    satis = veriler.get("Hasılat") or veriler.get("Net Satışlar")

    # Temel bileşenlerin varlığını kontrol et
    if None in [net_kar, finansman, vergi, satis]:
        raise ValueError("Bazı temel bileşenler eksik")

    # EBITDA hesapla
    ebitda = net_kar +- finansman +- vergi + amortisman

    # Sıfıra bölme hatasından kaçın
    if ebitda == 0 or satis == 0:
        raise ZeroDivisionError("EBITDA veya Hasılat sıfır olamaz")

    # Sonuçları kaydet
    sonuclar["EBITDA"] = ebitda
    sonuclar["EBITDA Marjı"] = ebitda / satis

except:
    sonuclar["EBITDA"] = "Hesaplanamadı"
    sonuclar["EBITDA Marjı"] = "Hesaplanamadı"

try:
    net_finansal_borc = veriler["Kısa Vadeli Borçlanmalar"] + veriler['Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları'] + veriler["Uzun Vadeli Borçlanmalar"] - veriler["Nakit ve Nakit Benzerleri"]
    sonuclar["Net Finansal Borç / EBITDA"] = net_finansal_borc / ebitda
except:
    sonuclar["Net Finansal Borç / EBITDA"] = "Hesaplanamadı"
    
# Sayıları biçimlendir: EBITDA hariç tümü 2 ondalıklı, EBITDA binlik ayraçlı tam sayı ve Biçimlendirilmiş değerleri ayrı listeye al

bicimlenmis_sonuclar = []

for oran, deger in sonuclar.items():
    if isinstance(deger, (int, float)):
        if oran == "EBITDA":
            deger_str = f"{deger:,.0f}".replace(",", ".") + " TL"  # binlik ayraçlı ve TL
        else:
            deger_str = f"{deger:.2f}"  # ondalıklı


    else:
        deger_str = deger  # "Hesaplanamadı" gibi metinler
    bicimlenmis_sonuclar.append((oran, deger_str))

# DataFrame'e çevirmek için;
df_sonuclar = pd.DataFrame(bicimlenmis_sonuclar, columns=["Oran", "Değer"])

# Excel dosyasına kaydet
df_sonuclar.to_excel("C:/Users/HP/Desktop/finansal_rasyolar.xlsx", index=False)

# Ekrana da yazdırmak için;
print(df_sonuclar)


# In[ ]:





# In[ ]:





# In[ ]:




