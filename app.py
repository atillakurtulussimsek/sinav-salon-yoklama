#!/usr/bin/env python3
"""Web arayüzü: Excel yükle, seçenekleri belirle, PDF indir.

Çalıştırma: uvicorn app:app --host 0.0.0.0 --port 8000
"""
import copy
import io
import json
import os
import re
import warnings
from datetime import datetime

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse

import main

warnings.filterwarnings("ignore", message="Workbook contains no default style")

KOK = os.path.dirname(os.path.abspath(__file__))
STATIK = os.path.join(KOK, "static")
MAX_BOYUT = 10 * 1024 * 1024  # 10 MB

app = FastAPI(title="Sınav Salon Yoklama Listesi")


def varsayilan_ayar():
    return main.ayar_yukle(main.VARSAYILAN_AYAR)


def bayrak(v):
    return str(v).lower() in ("1", "true", "on", "evet")


def dosya_adi(metin):
    t = metin.translate(str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU"))
    t = re.sub(r"[^A-Za-z0-9\s-]", "", t).strip()
    t = re.sub(r"\s+", "_", t)
    return t or "yoklama"


@app.get("/")
def anasayfa():
    return FileResponse(os.path.join(STATIK, "index.html"))


@app.get("/api/ayarlar")
def ayarlar():
    return JSONResponse(varsayilan_ayar())


@app.post("/api/olustur")
async def olustur(
    excel: UploadFile = File(...),
    kurumAd: str = Form(""),
    sinavAd: str = Form(""),
    listeAdi: str = Form("Yoklama Listesi"),
    tarih: str = Form(""),
    saat: str = Form(""),
    bilgiGoster: str = Form("true"),
    bilgiSalon: str = Form("true"),
    bilgiOgrenciSayisi: str = Form("true"),
    bilgiTarihSaat: str = Form("true"),
    bilgiGozetmen: str = Form("false"),
    sutunlar: str = Form("[]"),
    siralama: str = Form("ad"),
    uyariGoster: str = Form("true"),
    uyariBaslik: str = Form(""),
    uyariMaddeler: str = Form(""),
    imzaGoster: str = Form("true"),
    imzaBaslik: str = Form("Salon Gözetmeni"),
    imzaSatirlar: str = Form(""),
    cerceve: str = Form("true"),
    altBilgi: str = Form(""),
    sayfaNumarasi: str = Form("true"),
    olusturmaTarihi: str = Form("true"),
    yaziBoyutu: str = Form("9.5"),
    zebra: str = Form("true"),
    renkKoyu: str = Form(""),
):
    if not excel.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Yalnızca .xlsx dosyası yükleyin.")
    icerik = await excel.read()
    if len(icerik) > MAX_BOYUT:
        raise HTTPException(400, "Dosya 10 MB'tan büyük.")

    ayar = copy.deepcopy(varsayilan_ayar())
    ayar["kurum"]["ad"] = kurumAd.strip() or ayar["kurum"]["ad"]
    ayar["sinav"].update({"ad": sinavAd.strip(), "listeAdi": listeAdi.strip(),
                          "tarih": tarih.strip(), "saat": saat.strip()})
    ayar["bilgiKutusu"] = {
        "goster": bayrak(bilgiGoster), "salon": bayrak(bilgiSalon),
        "ogrenciSayisi": bayrak(bilgiOgrenciSayisi),
        "tarihSaat": bayrak(bilgiTarihSaat), "gozetmen": bayrak(bilgiGozetmen),
    }
    try:
        secili = set(json.loads(sutunlar))
    except json.JSONDecodeError:
        raise HTTPException(400, "Sütun listesi okunamadı.")
    ayar["sutunlar"] = {k: (k in secili) for k in main.SUTUN_TANIM}
    if not any(ayar["sutunlar"].values()):
        raise HTTPException(400, "En az bir sütun seçin.")
    ayar["siralama"] = siralama
    ayar["uyarilar"] = {
        "goster": bayrak(uyariGoster), "baslik": uyariBaslik.strip(),
        "maddeler": [m.strip() for m in uyariMaddeler.splitlines() if m.strip()],
    }
    ayar["gozetmenImza"] = {
        "goster": bayrak(imzaGoster), "baslik": imzaBaslik.strip() or "Salon Gözetmeni",
        "satirlar": [s.strip() for s in imzaSatirlar.splitlines() if s.strip()]
        or ["Adı Soyadı :", "İmza :"],
    }
    try:
        boyut = float(yaziBoyutu)
    except ValueError:
        boyut = 9.5
    ayar["sayfa"] = {
        "cerceve": bayrak(cerceve), "altBilgi": altBilgi.strip(),
        "sayfaNumarasi": bayrak(sayfaNumarasi), "olusturmaTarihi": bayrak(olusturmaTarihi),
        "yaziBoyutu": max(7.0, min(boyut, 14.0)), "zebra": bayrak(zebra),
    }
    if re.fullmatch(r"#[0-9a-fA-F]{6}", renkKoyu.strip()):
        ayar["renkler"]["koyu"] = renkKoyu.strip()

    try:
        salonlar = main.salonlari_oku(io.BytesIO(icerik), ayar)
    except SystemExit as e:
        raise HTTPException(400, str(e))
    except Exception:
        raise HTTPException(400, "Excel dosyası okunamadı.")
    if not salonlar:
        raise HTTPException(400, "Excel dosyasında liste bulunamadı.")

    pdf = io.BytesIO()
    main.pdf_olustur(salonlar, pdf, ayar)
    pdf.seek(0)
    ad = f"{dosya_adi(sinavAd)}_{datetime.now():%Y%m%d}.pdf"
    return StreamingResponse(
        pdf, media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{ad}"',
                 "X-Salon-Sayisi": str(len(salonlar))})
