# Wattaw A.\u015e. Kartvizit Tarama ve G\u00f6r\u00fc\u015fme Takip Uygulamas\u0131

Bu klas\u00f6r Wattaw A.\u015e. ekibi i\u00e7in haz\u0131rlanan masa\u00fcst\u00fc kartvizit kay\u0131t ve g\u00f6r\u00fc\u015fme takip arac\u0131n\u0131 i\u00e7erir. Uygulama Tkinter tabanl\u0131d\u0131r ve kartvizit g\f6rsellerini Wattaw logolar\u0131yla beraber g\u00f6sterir.

## \u00d6zellikler

- Kartvizit g\u00f6rsellerini i\u00e7eri aktararak m\u00fc\u015fteri kayd\u0131 olu\u015fturma
- Sekt\u00f6r, ileti\u015fim bilgileri ve kart \u00fczerindeki notlar\u0131 saklama
- Kartvizit dosyalar\u0131n\u0131 otomatik olarak ar\u015fiv klas\u00f6r\u00fcne kopyalama
- G\u00f6r\u00fc\u015fme tarih/saat, konu ve notlar\u0131 ayri Excel dosyas\u0131nda tutma
- "customers.xlsx" ve "meetings.xlsx" dosyalar\u0131n\u0131 otomatik olarak g\u00fcncelleme
- Gereken Python paketlerini uygulama a\u00e7\u0131l\u0131\u015f\u0131nda otomatik y\u00fckleme/g\u00fcncelleme

## Kullan\u0131m

1. Python 3.9+ y\u00fckl\u00fc oldu\u011fundan emin olun.
2. Bu klas\u00f6re gidin ve uygulamay\u0131 \u015fu komutla ba\u015flat\u0131n:

   ```bash
   python app.py
   ```

3. Uygulama a\u00e7\u0131ld\u0131\u011f\u0131nda eksik paketler varsa otomatik olarak kurulacakt\u0131r.
4. "Kartvizit Kayd\u0131" sekmesinden kartviziti se\u00e7erek ilgili alanlar\u0131 doldurun ve kaydedin.
5. "G\u00f6r\u00fc\u015fme Notlar\u0131" sekmesinden g\u00f6r\u00fc\u015fme detaylar\u0131n\u0131 ekleyin.

Kartvizit kayıtlar\u0131 `data/customers.xlsx`, g\u00f6r\u00fc\u015fme notlar\u0131 ise `data/meetings.xlsx` dosyalar\u0131nda saklan\u0131r. Kart g\u00f6rselleri `data/business_cards/` klas\u00f6r\u00fcne kopyalan\u0131r.

## Ba\u011f\u0131ml\u0131l\u0131klar

- pandas
- openpyxl
- Pillow
- cairosvg

Bu paketlerin tamam\u0131 uygulama \u00e7al\u0131\u015f\u0131rken otomatik olarak kontrol edilir ve gerekirse y\u00fcklenir.

## Notlar

- OCR tabanl\u0131 otomatik bilgi okuma deste\u011fi dahil edilmemi\u015ftir. Kartvizit bilgilerini kay\u0131t esnas\u0131nda manuel olarak doldurabilirsiniz.
- Excel dosyalar\u0131n\u0131n format\u0131n\u0131 de\u011fi\u015ftirmeden kullanman\u0131z tavsiye edilir.
- Uygulama penceresi Tkinter \u00fczerinden \u00e7al\u0131\u015ft\u0131\u011f\u0131 i\u00e7in Windows, macOS ve Linux ortamlar\u0131nda Python kurulu olmas\u0131 yeterlidir.
