# Data Process

**Kayıtlarınızı tek bir masaüstü panelinden yönetin, görselleştirin ve Excel’e aktarın.**

Data Process; kategorilere ayrılmış sayısal verileri takip etmek için geliştirilen, Türkçe arayüze sahip bir masaüstü uygulamasıdır. PyQt5 ile oluşturulan arayüz, SQLite üzerinde yerel veri saklama ve Matplotlib grafikleriyle çalışır.

## Özellikler

- **Özet paneli:** Toplam değer, kayıt sayısı ve en yüksek kaydı görüntüleyin.
- **Kategori filtresi:** Özet kartlarını, grafikleri ve kayıt tablosunu birlikte filtreleyin.
- **Veri görselleştirme:** Kategori dağılımını halka grafikle, kategori toplamlarını yatay çubuklarla ve kayıtların zaman içindeki değişimini çizgi grafikle inceleyin.
- **Kayıt yönetimi:** Yeni kayıt ekleyin, tablodan seçtiğiniz kaydı silin veya onay vererek tüm kayıtları temizleyin.
- **Excel aktarımı:** Tüm kayıtları biçimlendirilmiş bir `.xlsx` dosyasına kaydedin.
- **Yerel depolama:** Verilerinizi uygulama klasöründeki SQLite veritabanında tutun.
- **Türkçe sayı girişi:** Ondalık değerlerde virgül veya nokta kullanın.

## Teknolojiler

| Bileşen | Teknoloji |
| --- | --- |
| Programlama dili | Python |
| Masaüstü arayüzü | PyQt5 |
| Veritabanı | SQLite |
| Grafikler | Matplotlib |
| Excel çıktısı | openpyxl |

## Kurulum

Python **3.10 veya üzeri** gereklidir. Aşağıdaki komutlar Windows / PowerShell içindir.

1. Depoyu ZIP olarak indirin veya Git ile klonlayın, ardından proje klasöründe bir terminal açın.
2. Sanal ortam oluşturun:

   ```powershell
   python -m venv .venv
   ```

3. Bağımlılıkları yükleyin:

   ```powershell
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

4. Uygulamayı başlatın:

   ```powershell
   .\.venv\Scripts\python.exe main.py
   ```

`data.db` dosyası ve gerekli tablo, mevcut değilse ilk çalıştırmada otomatik oluşturulur. Mevcut kayıtlar sonraki açılışlarda yüklenir.

## Kullanım

### Kayıt ekleme

Sol panelde **Kategori** ve **Değer** alanlarını doldurup **Kayıt ekle** düğmesine basın. Değer alanında `Enter` tuşunu da kullanabilirsiniz.

Örnek: `Satış` kategorisi için `1250,50` veya `1250.50` girin. Sayıları binlik ayırıcı olmadan yazın. Yeni kayıt eklendiğinde filtre tüm kategorilere döner.

### Verileri inceleme

Üst bölümdeki kategori listesinden bir seçim yapın. Kartlar, grafikler ve tablo seçilen kategoriye göre güncellenir. **Tüm kategoriler** seçeneği genel görünümü geri getirir.

- Kayıt tablosu en son eklenen kaydı üstte gösterir.
- Karşılaştırma grafiği, toplam değere göre ilk 8 kategoriyi gösterir.
- Halka grafik, pozitif paya sahip kategori sayısı 5’i aşarsa kalanları **Diğer kategoriler** altında toplar.
- Zaman grafiği kayıtları tarih sırasına göre gösterir.
- Sıfır ve negatif değerler karşılaştırma ve zaman grafiklerinde gösterilir. Halka grafik için kategori toplamlarının negatif olmaması ve genel toplamın sıfırdan büyük olması gerekir.

### Kayıt silme

Tablodan bir satır seçip **Seçileni sil** düğmesine basın. **Tümünü sil**, onayınızın ardından filtre dışında kalanlar dahil bütün kayıtları kalıcı olarak siler.

### Excel’e aktarma

**Excel’e aktar** düğmesine basıp dosyanın kaydedileceği konumu seçin. Aktarım, etkin kategori filtresinden bağımsız olarak **tüm kayıtları** içerir. Dosyada kategori, değer ve kayıt tarihi sütunları bulunur.

## Proje yapısı

```text
Data Process/
├── main.py            # Uygulama, arayüz ve veri işlemleri
├── requirements.txt   # Python bağımlılıkları
├── test_ui.py         # Arayüz ve işlev testleri
├── README.md          # Proje dokümantasyonu
└── data.db            # Yerel veritabanı (otomatik oluşturulur)
```

## Testler

Bağımlılıkları yükledikten sonra test paketini çalıştırın:

```powershell
.\.venv\Scripts\python.exe test_ui.py
```

Testler bellekte oluşturulan ayrı bir veritabanı kullanır; mevcut `data.db` dosyasını değiştirmez. Kapsam:

- Kategori filtreleme ve doğru kaydın silinmesi
- Geçersiz girişlerin reddedilmesi
- Boş veri, sıfır ve negatif değerlerde grafik oluşturma
- Toplu silme onayı
- Excel aktarımının tüm kayıtları içermesi
- İki pencere boyutunda arayüz görüntüsü oluşturma

Görünüm testi, proje klasörüne `ui-preview.png` ve `ui-preview-small.png` dosyalarını yazar.

## Sorun giderme

**Bağımlılık bulunamıyor:** Uygulamayı, bağımlılıkları yüklediğiniz sanal ortamın Python yorumlayıcısıyla çalıştırdığınızdan emin olun.

**Windows Uygulama Denetimi DLL dosyasını engelliyor:** Kurumsal güvenlik politikaları Qt veya Matplotlib bileşenlerinin yüklenmesini engelleyebilir. Onaylı bir Python ortamı kullanın veya sistem yöneticinizle görüşün.

## Katkıda bulunma

Hata bildirimleri ve geliştirme önerileri için bir issue açabilirsiniz. Hata bildirimine kullandığınız Python sürümünü, işletim sistemini ve sorunu yeniden oluşturma adımlarını ekleyin. Kod değişikliklerinde ilgili testleri çalıştırıp sonucu pull request açıklamasında paylaşın.
