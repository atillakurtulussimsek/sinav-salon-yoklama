# Sınav Salon Yoklama Listesi

Excel sınav yerleşim listesinden, her salon için ayrı sayfada **yoklama listesi PDF'i** üretir.
Komut satırından veya web arayüzünden kullanılabilir.

Her salon sayfasında:

- Kurum ve sınav başlığı
- Bilgi kutusu (salon, öğrenci sayısı, tarih/saat, gözetmen)
- Öğrenci tablosu (seçilebilir sütunlar, imza alanı)
- Salon gözetmenine uyarılar
- Sağ altta gözetmen adı soyadı / imza bloğu

## Excel formatı

İlk satır başlık olmalı. Gerekli sütunlar:

| Sütun | Açıklama |
|---|---|
| `Ad Soyad` | Öğrenci adı |
| `Okul Numarası` | Öğrenci numarası |
| `Sınıf` | Sınıf / şube |
| `Salon adı` | Salon; her farklı değer ayrı sayfa olur |
| `Sira` | Koltuk sırası |
| `Sütun` | Koltuk sütunu |

İsteğe bağlı: `Kitapcık türü`, `Öğrenci Telefonu`, `Veli Adı Soyadı`, `Veli Telefonu`.
Sütun adları büyük/küçük harfe duyarsızdır ve `ayarlar.json` → `excelSutunlari` bölümünden değiştirilebilir.

## Kurulum

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Türkçe karakterler için TTF font gerekir. macOS'ta Arial, Linux'ta DejaVu Sans
(`fonts-dejavu-core`), Windows'ta Arial otomatik bulunur.

## Komut satırı

```bash
.venv/bin/python main.py liste.xlsx                      # liste.pdf üretir
.venv/bin/python main.py liste.xlsx cikti.pdf            # çıktı adı
.venv/bin/python main.py liste.xlsx --ayar tyt.json      # farklı ayar dosyası
```

`--ayar` verilmezse `main.py` yanındaki `ayarlar.json` kullanılır.

## Web arayüzü

```bash
.venv/bin/uvicorn app:app --host 0.0.0.0 --port 8000
```

`http://localhost:8000` adresinde Excel dosyasını seçin, seçenekleri belirleyin, PDF'i indirin.
Form varsayılanları `ayarlar.json` dosyasından gelir. Yüklenen dosyalar diske yazılmaz.

## Docker / Dokploy

```bash
docker build -t yoklama .
docker run -p 8000:8000 yoklama
```

Dokploy'da: depoyu bağlayın, derleme tipi **Dockerfile**, port **8000**.

## Ayar dosyası (`ayarlar.json`)

| Bölüm | İçerik |
|---|---|
| `kurum.ad` | Başlıktaki kurum adı |
| `sinav` | `ad`, `listeAdi`, `tarih`, `saat` |
| `bilgiKutusu` | Hangi bilgi alanlarının gösterileceği |
| `sutunlar` | Tablo sütunları aç/kapa: `koltuk`, `ogrenciNo`, `adSoyad`, `sinif`, `kitapcik`, `ogrenciTelefonu`, `veliAdi`, `veliTelefonu`, `imza` |
| `sutunBasliklari` | Sütun başlık metinleri |
| `siralama` | `ad` (ad soyada göre) veya `koltuk` (sıra-sütun) |
| `uyarilar` | `goster`, `baslik`, `maddeler` (markdown: `**kalın**`, `*italik*`) |
| `gozetmenImza` | İmza bloğu başlığı ve satırları |
| `sayfa` | Çerçeve, altbilgi, sayfa numarası, tarih, yazı boyutu, zebra satırlar |
| `renkler` | `koyu`, `acik`, `zebra`, `cizgi` (hex) |
| `excelSutunlari` | Excel başlıklarının eşlemesi |

Farklı sınavlar için `ayarlar.json` kopyalanıp `--ayar` ile verilebilir.

## Dosyalar

```
main.py        PDF üretimi ve komut satırı
app.py         FastAPI web uygulaması
static/        Web arayüzü
ayarlar.json   Varsayılan ayarlar
Dockerfile     Konteyner imajı
```

## Lisans

MIT
