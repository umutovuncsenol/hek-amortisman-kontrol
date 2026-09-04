# HEK / Amortisman Eklenti Kontrolü

HEK Formu ve Amortisman Excel dosyalarını süreç numaralarına göre eşleştirir.

Uygulama sarı–mavi arayüz kullanır; başlıkta şeffaf PNG Renault logosu ve
Windows uygulama ikonunda şeffaf Renault simgesi bulunur.

## Girdiler

1. Excel'den tek sütun olarak kopyalanıp satır numaralı alana yapıştırılan süreç numaraları
2. HEK Formu (`AK`: Process ID, `H`: Envanter, `I`: Eklenti, `P`: Kısmi Çıkış Oranı / `txtKismiCikisOrani`)
3. Amortisman dosyası (`A`: Immobilisation, `B`: Numero subsidiaire)

## Çıktı

Uygulama üç sayfalı bir Excel dosyası üretir:

- **Kontrol Özeti:** Eklenti tamamlama oranları ve uyarılar
- **Filtrelenmiş Amortisman:** Eşleşen Amortisman satırlarının tüm sütunları ve HEK kontrolü
- **Eşleşmeyen Kayıtlar:** HEK veya Amortisman tarafında bulunamayan kayıtlar

Aynı envantere bağlı süreç numaraları tek blokta birleştirilir. Amortisman'daki
`0` ana ürün, diğer alt numaralar eklenti kabul edilir.
