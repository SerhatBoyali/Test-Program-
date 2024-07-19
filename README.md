# Test Programı
Modbus Protokolü ile veri okuyup işleyen Test Uygulaması.
serhat.boyali@outlook.com mail adresi ile benimle iletişime geçebilirsiniz.

# Uygulamanın detayları
Tasarladığımız test cihazını modbus protokolü kullanarak bilgisayara bir usb dönüştürücüsüyle bağlıyoruz.
Cihazdan Voltaj, Akım ve NTC ısı verileri gelmektedir ve akım ile voltaj çarpımından WATT (güç) değerini ölçüyorum. Veriler 2 adet 8 er bit şeklinde gelmekte (1 veri 16 bit). Uygulamada bu iki veriyi 1 float verisine dönüştürüyorum.
10 tane cihaza kadar test edebilecek şekilde tasarladım. Cihazın ID sine göre tabpage ve ZedGraph kütüphanesini kullandığım grafik ekranım var. ZedGraph grafik ekranı grafiği  kaydetme, yakınlaştırma, uzaklaştırma, yazdırma, ölçeği sıfırlama ve verilerin değerlerini gösterme işlemlerini kolayca yapmaktadır.

# Kullanım Talimatı

1. Modbus protokolü ile tasarladığınız test cihazınızı bir USB dönüştürücü ile bilgisayarınızı takınız, USB dönüştürücünün sürücüsünü indirmeyi unutmayınız.
2. Uygulamayı çalıştırdıktan sonra port ve bautrate seçimi yapınız ardından Porta Bağlan butonuna basınız. Bağlantı başarılı olursa olumlu veya olumsuz mesaj gelecek ve "Port Bağlantısı:" durumu Açık veya Kapalı diye güncellenecek.
3. Port bağlantısı kurulduktan sonra test edeceğiniz cihazın ID numarasına göre Slave ID sayfası seçin ve Connect diyerek bağlanın. Bağlantının başarılı yada başarısız olma durumu mesaj bildirimi ile gelecektir ve "Bağlantı Açık" veya "Bağlantı Kapalı" diye yazacaktır.
4. Veriler grafik ekranına ve textbox'larda görünmeye başlayacaktır.
5. Disconnect butonu ile grafiğe veri aktarma işlemini durdurabilirsiniz ve Bağlantıyı Kes butonu ile port bağlantısını durdurabilirsiniz.
6. Verileri Kaydet butonuna tıklayarak ölçtüğünüz verileri bir excell sayfası halinde kaydedebilirsiniz. Veriler "Sistem Saati	Test Süresi	VOLTAJ	AKIM	NTC1	NTC2	NTC3	NTC4	NTC5	NTC6	WATT" sütunlarında saniyede bir kaydedilir.
7. Formu temizle butonu ile ayarları sıfırlayabilirsiniz ve Programı kapat butonu ile direkt uygulamayı kapatabilirsiniz.

   
![test](https://github.com/user-attachments/assets/1b729616-2e20-42b8-98a1-a195dcd76fa7)

Excell de veriler şu şekilde kayıt altına alınmaktadır;

![image](https://github.com/user-attachments/assets/2b778f01-1bf6-45fb-a589-0948d68edd43)

Bir testin grafik ekranı görüntüsü ; 

![image](https://github.com/user-attachments/assets/accb0d71-b2a8-495a-8e45-b57fc8d55361)
