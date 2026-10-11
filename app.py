import re
from datetime import date, timedelta
from urllib.parse import urlencode, quote

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Ekip İçi Günlük İş Akışı", page_icon="📬", layout="wide")
st.title("📬 Ekip İçi Günlük İş Akışı Otomasyonu")
st.caption("Excel'den otomatik alıcılı, kişiye özel Outlook e-posta taslakları hazırlar. Gönderim kullanıcı onayıyla yapılır.")

KATEGORILER = ["İzinli Listesi", "Geciken İşler", "Bugünkü İşler", "Yaklaşan İşler"]
GEREKLI = ["Çalışan", "E-posta", "İzin Durumu", "İş / Görev", "Son Tarih", "Durum", "Öncelik"]


def temiz(v):
    return "" if pd.isna(v) else str(v).strip()


def gecerli_eposta(adres):
    return bool(re.fullmatch(r"[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+", adres))


def outlook_link(alicilar, konu, govde):
    return "https://outlook.live.com/mail/0/deeplink/compose?" + urlencode(
        {"to": ";".join(alicilar), "subject": konu, "body": govde},
        quote_via=quote,
    )


uploaded = st.file_uploader("Ekip takip Excel dosyasını yükleyin", type=["xlsx"])
if uploaded is None:
    st.info("Başlamak için 'Ekip Takip' sayfası bulunan Excel dosyasını yükleyin.")
    st.stop()

try:
    df = pd.read_excel(uploaded, sheet_name="Ekip Takip")
except Exception as exc:
    st.error(f"'Ekip Takip' sayfası okunamadı: {exc}")
    st.stop()

eksik = [s for s in GEREKLI if s not in df.columns]
if eksik:
    st.error("Excel'de eksik sütunlar: " + ", ".join(eksik))
    st.stop()

for s in ["Çalışan", "E-posta", "İzin Durumu", "İş / Görev", "Durum", "Öncelik"]:
    df[s] = df[s].map(temiz)
df["Son Tarih"] = pd.to_datetime(df["Son Tarih"], errors="coerce").dt.date
df = df[df["Çalışan"] != ""].copy()
if df.empty:
    st.warning("Excel'de çalışan bulunamadı.")
    st.stop()

# Tek çalışan birden fazla görevde yer alabilir, ancak e-posta eşleşmesi tutarlı olmalı.
eslesme = df.groupby("Çalışan")["E-posta"].nunique()
celiskili = eslesme[eslesme > 1].index.tolist()
if celiskili:
    st.error("Birden fazla e-posta tanımlanan çalışanlar: " + ", ".join(celiskili))
    st.stop()

kisiler = df[["Çalışan", "E-posta"]].drop_duplicates("Çalışan").copy()
hatali = kisiler[~kisiler["E-posta"].map(gecerli_eposta)]
if not hatali.empty:
    st.error("E-posta adresi eksik/geçersiz çalışanlar: " + ", ".join(hatali["Çalışan"]))
    st.stop()

bugun = date.today()
yedi_gun = bugun + timedelta(days=7)
tarih = bugun.strftime("%d.%m.%Y")
izinliler = df[df["İzin Durumu"].str.casefold().ne("çalışıyor") & df["İzin Durumu"].ne("")].drop_duplicates("Çalışan")
aktif = df[df["Durum"].str.casefold().ne("tamamlandı") & df["İş / Görev"].ne("")].copy()
geciken = aktif[aktif["Son Tarih"].notna() & (aktif["Son Tarih"] < bugun)]
bugunku = aktif[aktif["Son Tarih"] == bugun]
yaklasan = aktif[aktif["Son Tarih"].notna() & (aktif["Son Tarih"] > bugun) & (aktif["Son Tarih"] <= yedi_gun)]

c1, c2, c3, c4 = st.columns(4)
c1.metric("🏖️ İzinli", len(izinliler))
c2.metric("🔴 Geciken İş", len(geciken))
c3.metric("🟠 Bugünkü İş", len(bugunku))
c4.metric("📅 7 Gün İçinde", len(yaklasan))
st.divider()

taslaklar = []
izin_satirlari = []
for _, r in izinliler.iterrows():
    tur = "Yıllık İzin" if r["İzin Durumu"].casefold() == "izinli" else r["İzin Durumu"]
    izin_satirlari.append(f"• {r['Çalışan']} — {tur}")
izin_detay = "\n".join(izin_satirlari) if izin_satirlari else "Bugün izinli çalışan bulunmamaktadır."
izin_govde = (
    f"Günaydın,\n\nBugün izinli olan çalışanlar:\n\n{izin_detay}\n\n"
    f"Toplam ekip: {len(kisiler)} kişi\n"
    f"İzinli kişi sayısı: {len(izinliler)}\n"
    f"Aktif çalışan sayısı: {len(kisiler) - len(izinliler)}\n\nİyi çalışmalar."
)
taslaklar.append({
    "kategori": "İzinli Listesi", "kisi": "Tüm ekip",
    "alicilar": kisiler["E-posta"].drop_duplicates().tolist(),
    "konu": f"Günlük İzinli Listesi – {tarih}", "govde": izin_govde,
})

raporlar = [
    ("Geciken İşler", geciken),
    ("Bugünkü İşler", bugunku),
    ("Yaklaşan İşler", yaklasan),
]
for kategori, tablo in raporlar:
    for kisi, grup in tablo.groupby("Çalışan", sort=True):
        adres = kisiler.loc[kisiler["Çalışan"] == kisi, "E-posta"].iloc[0]
        satirlar = []
        for _, r in grup.iterrows():
            gorev = r["İş / Görev"]
            if kategori == "Geciken İşler":
                ek = f"{(bugun - r['Son Tarih']).days} gün gecikmiş"
            elif kategori == "Bugünkü İşler":
                ek = f"Öncelik: {r['Öncelik'] or 'Belirtilmedi'}"
            else:
                ek = f"Son tarih: {r['Son Tarih'].strftime('%d.%m.%Y')}"
            satirlar.append(f"• {gorev} — {ek}")
        detay = "\n".join(satirlar)
        govde = (
            f"Merhaba {kisi},\n\n{kategori} kapsamında takip etmen gereken görevler:\n\n"
            f"{detay}\n\nToplam iş: {len(grup)}\n\nİyi çalışmalar."
        )
        taslaklar.append({
            "kategori": kategori, "kisi": kisi, "alicilar": [adres],
            "konu": f"{kategori} – {tarih}", "govde": govde,
        })

st.subheader("✉️ Günlük Mail Taslakları")
st.caption(f"Toplam {len(taslaklar)} taslak hazırlandı. Kategoriyi seçerek kişiye özel mailleri inceleyebilirsin.")
secim = st.radio("Hazırlanacak mail:", KATEGORILER, horizontal=True)
gosterilen = [t for t in taslaklar if t["kategori"] == secim]
if not gosterilen:
    st.info("Bu kategoride gönderilecek görev bulunmuyor.")

for i, taslak in enumerate(gosterilen):
    with st.expander(f"{taslak['kategori']} | {taslak['kisi']} | {len(taslak['alicilar'])} alıcı", expanded=(i == 0)):
        st.write("**Alıcılar:** " + "; ".join(taslak["alicilar"]))
        st.write("**Konu:** " + taslak["konu"])
        st.text_area("Mesaj önizlemesi", taslak["govde"], height=230, key=f"govde_{secim}_{i}")
        demo_modu = st.session_state.get("demo_modu", True)

        alicilar = (
           [st.session_state.get("demo_adresi", "test@example.com")]
            if demo_modu
            else taslak["alicilar"]
        )

        url = outlook_link(alicilar, taslak["konu"], taslak["govde"])
        if len(url) > 1800:
            st.warning("Bu taslak bağlantı için uzun olabilir. Outlook'ta eksik açılırsa metni kopyalayarak kontrol edin.")
        st.link_button("📨 Outlook'ta Taslağı Aç", url)

st.info("Bu sürüm Outlook taslaklarını hazırlar, otomatik göndermez. example.com adresleri örnektir; gerçek e-posta gönderimi için kullanılmaz.")

st.divider()
st.subheader("🧪 Demo Modu")

demo_modu = st.toggle(
    "Test gönderimlerini kendi Outlook adresime yönlendir",
    value=False,
    key="demo_modu"
)

demo_adresi = st.text_input(
    "Test e-posta adresi",
    value="",
    placeholder="ornek@outlook.com",
    key="demo_adresi"
)

if demo_modu:
    st.info(
        f"Demo modu açık. Test e-postaları {demo_adresi} adresine yönlendirilecek."
    )
