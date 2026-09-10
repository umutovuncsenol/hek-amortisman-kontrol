# HEK / Amortisman Eklenti Kontrolü — Kullanım ve Teknik Dokümantasyon

## Amaç

Uygulama, verilen süreç numaralarını HEK Formu’nda bulur; envanter ve eklenti numaralarını Amortisman dosyasıyla eşleştirir. Hurdalanacak ve kalan eklentileri, kısmi çıkış oranlarını ve eşleşmeyen kayıtları gösteren üç sayfalı bir Excel raporu üretir.

## Gereksinimler

Son kullanıcı için Windows EXE paketi tercih edilir. Kaynak koddan çalıştırmak için Python 3 ve `openpyxl>=3.1,<4` gerekir. Girdiler `.xlsx` veya `.xlsm` olmalıdır.

Uygulama yerel çalışır, dış servise veri göndermez ve giriş dosyalarını değiştirmez.

## Girdi veri sözleşmesi

Sütunlar başlık adına göre değil, sabit konuma göre okunur.

| Dosya | Sütun | Alan |
|---|---|---|
| HEK Formu | H | Envanter numarası |
| HEK Formu | I | Eklenti numarası |
| HEK Formu | P | `txtKismiCikisOrani` |
| HEK Formu | AK | Process ID |
| Amortisman | A | Immobilisation |
| Amortisman | B | Numero subsidiaire |

Kaynak şablonda bu kolonlar değişirse uygulama kodu ve testler güncellenmelidir.

## Kullanım

1. Excel’den süreç numarası sütununu kopyalayın.
2. Uygulamada **Panodan Yapıştır** düğmesine basın. Her satırda tek süreç olmalıdır.
3. Sarı tekrar satırlarını ve kırmızı, birden fazla sütun içeren geçersiz satırları kontrol edin.
4. HEK Formu ve Amortisman dosyalarını seçin.
5. İki dosyada doğru sayfa ve başlık satırını seçin.
6. **Analizi Başlat** düğmesine basın.
7. Sonuç `.xlsx` dosyasını kaydedin.

## İş kuralları

- Tekrarlanan süreç yalnızca bir kez analiz edilir.
- HEK’te bulunmayan süreçler eşleşmeyen kayıtlara yazılır.
- HEK satırında envanter veya eklenti boşsa kayıt eşleşmeyenlere yazılır.
- Eşleşme, envanter + alt numara çiftiyle yapılır.
- Amortisman alt numarası `0` ana ürün, diğer değerler eklentidir.
- Amortisman’da bulunup HEK talebinde olmayan eklenti “Kalan eklenti”dir.
- Tamamlanma oranı `hurdalanacak eklenti / toplam eklenti` olarak hesaplanır.
- Eklenti bulunmayan envanter `%100` tamamlanmış kabul edilir.
- `1`, `100` ve `100%` eşdeğer biçimleri `%100` kabul edilir.
- Aynı envantere bağlı süreçler tek sonuç bloğunda birleştirilir.

## Çıktı sayfaları

### Kontrol Özeti

Envanter bazında süreçleri, toplam/hurdalanacak/kalan eklentileri, tamamlanma yüzdesini, kısmi oran kontrolünü ve genel durumu gösterir. Yeşil tamam, turuncu uyarı, kırmızı kalan eklenti anlamındadır.

### Filtrelenmiş Amortisman

İlgili Amortisman satırlarını bütün kaynak sütunlarıyla taşır; HEK talep durumu, talep eden süreç, kısmi çıkış oranı ve kontrol açıklaması ekler.

### Eşleşmeyen Kayıtlar

HEK’te bulunmayan süreçleri, eksik anahtarlı HEK satırlarını ve Amortisman’da karşılığı olmayan HEK kayıtlarını listeler.

## Teknik akış

1. `parse_process_input` süreçleri temizler.
2. `collect_hek` ilgili HEK kayıtlarını toplar.
3. `collect_amortisation` ilgili envanterlerin Amortisman satırlarını tarar.
4. `build_analysis` eşleşmeleri ve iş kurallarını uygular.
5. `write_result_workbook` biçimlendirilmiş üç sayfalı raporu üretir.
6. `run_analysis` tüm süreci yönetir.

## Test

```text
python tests/test_app.py
```

Test; tekrar süreç, geçersiz çok sütunlu yapıştırma, eksik süreç, kalan eklenti, `%100` olmayan oran, büyük Amortisman taraması ve çalışma kitabı yapısını doğrular.

## Windows EXE üretimi

Windows’ta `WINDOWS_EXE_OLUSTUR.bat` çalıştırıldığında `dist/HEKAmortismanKontrol.exe` ve ZIP paketi üretilir. GitHub’daki **Windows EXE Olustur** Actions iş akışı elle veya `v*` etiketiyle tetiklenebilir; etiketli sürüm GitHub Release oluşturur.

Derleme Python 3.12 ve `pyinstaller==6.22.2` kullanır. EXE imzasızsa Windows bilinmeyen yayıncı uyarısı gösterebilir; kurumsal kod imzalama değerlendirilmelidir.

## Dosya ve varlıklar

- `hek_amortisman_kontrol.py`: ana uygulama
- `tests/test_app.py`: regresyon testi
- `.github/workflows/windows-exe.yml`: Windows paketleme hattı
- `assets/HEK_Amortisman_Ornek_Cikti.xlsx`: yardım ekranındaki örnek
- `assets/renault-logo-*`: arayüz ve paketleme görselleri
- `KULLANIM.txt`: kısa son kullanıcı rehberi

Logo/marka kullanım yetkisi şirket tarafından doğrulanmalıdır. Gerçek şirket verileri, `node_modules`, `__pycache__`, sanal ortamlar ve geçici çıktılar depoya eklenmemelidir.

## Bakım

Sabit kolonlar dosyanın başındaki `HEK_*_COL` ve `AMORT_*_COL` sabitlerinde; oran yorumu `ratio_info`; iş kuralları `build_analysis`; Excel sayfa/renk düzeni `write_result_workbook` içinde yönetilir.

Şablon, alan anlamı veya iş terminolojisi değiştiğinde kod, test, örnek çıktı ve bu belge birlikte güncellenmelidir.

## Sorun giderme

- Süreç bulunamıyorsa doğru HEK sayfasını ve AK kolonunu kontrol edin.
- Çok sayıda eşleşmeyen kayıt varsa H/I ve A/B kolonlarının değişmediğini doğrulayın.
- Formül sonuçları eskiyse dosyayı Excel’de hesaplatıp kaydedin.
- EXE açılmıyorsa Windows güvenlik/antivirüs kayıtlarını kontrol edin.
- Örnek çıktı yoksa ilgili Excel’in `assets` klasöründe ve PyInstaller paket listesinde olduğunu doğrulayın.

## Devir notları

- Teslim edilen GitHub etiketi/commit’i kaydedilmelidir.
- Depo ve Actions yönetimi şirket hesabına verilmelidir.
- Bakım sorumlusu, hata bildirim kanalı ve sürümleme yöntemi belirlenmelidir.
- Üretilen raporlar kurumsal veri sınıflandırma ve saklama politikasına göre korunmalıdır.
