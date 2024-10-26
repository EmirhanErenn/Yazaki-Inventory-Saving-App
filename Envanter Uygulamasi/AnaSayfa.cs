using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace Envanter_Uygulamasi
{
    public partial class AnaSayfa : Form
    {
        OleDbConnection baglanti;
        OleDbCommand komut;
        OleDbDataAdapter da1;

        public AnaSayfa()
        {
            InitializeComponent();
            this.AutoScaleMode = AutoScaleMode.Dpi;

            baglanti = new OleDbConnection();
            komut = new OleDbCommand();
            da1 = new OleDbDataAdapter();

            string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            baglanti.ConnectionString = baglan;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sorgu = "INSERT INTO envanter ([Asset_No], [Lokasyon], [Yeni_Hostname], [Eski_Hostname], [Kullanici], [Kategori], [Marka], [Model], [Seri_No], [IP_No], [Bulundugu_Bolge], [Mac_Adres], [Mac_Adress_2], [Wireless_Mac_Adres], [Tedarik_Firmasi], [Alis_Tarihi], [Garanti_Suresi], [Eski_Kullanici], [Docking_Station_MAC_Adress], [Docking_Station_IP_Adress], [Switch], [Port], [Isletim_Sistemi], [Virus], [Islemci], [Bellek], [LVA], [Domain], [Office], [Bitlocker], [Zimmet_Baslangic], [Zimmet_Bitis], [Kiralama_Baslangic], [Kiralama_Bitis], [Kiralanan_Firma], [Durumu], [Aciklama]) VALUES (@asset_no, @lokasyon, @yeni_hostname, @eski_hostname, @kullanici, @kategori, @marka, @model, @seri_no, @ip_no, @bulundugu_bolge, @mac_adres, @mac_adress_2, @wireless_mac_adres, @tedarik_firmasi, @alis_tarihi, @garanti_suresi, @eski_kullanici, @docking_station_mac_adress, @docking_station_ip_adress, @switch, @port, @isletim_sistemi, @virus, @islemci, @bellek, @lva, @domain, @office, @bitlocker, @zimmet_baslangic, @zimmet_bitis, @kiralama_baslangic, @kiralama_bitis, @kiralanan_firma, @durumu, @aciklama)";

            komut = new OleDbCommand(sorgu, baglanti);

            // Parametrelerin eklenmesi
            komut.Parameters.AddWithValue("@asset_no", textBox1.Text);
            komut.Parameters.AddWithValue("@lokasyon", comboBox2.SelectedItem?.ToString() ?? (object)DBNull.Value);
            komut.Parameters.AddWithValue("@yeni_hostname", textBox4.Text);
            komut.Parameters.AddWithValue("@eski_hostname", textBox3.Text);
            komut.Parameters.AddWithValue("@kullanici", textBox5.Text);
            komut.Parameters.AddWithValue("@kategori", comboBox1.SelectedItem?.ToString() ?? (object)DBNull.Value);
            komut.Parameters.AddWithValue("@marka", textBox6.Text);
            komut.Parameters.AddWithValue("@model", textBox7.Text);
            komut.Parameters.AddWithValue("@seri_no", textBox8.Text);
            komut.Parameters.AddWithValue("@ip_no", textBox12.Text);
            komut.Parameters.AddWithValue("@bulundugu_bolge", textBox13.Text);
            komut.Parameters.AddWithValue("@mac_adres", textBox9.Text);
            komut.Parameters.AddWithValue("@mac_adress_2", textBox10.Text);
            komut.Parameters.AddWithValue("@wireless_mac_adres", textBox11.Text);
            komut.Parameters.AddWithValue("@tedarik_firmasi", textBox21.Text);
            komut.Parameters.AddWithValue("@alis_tarihi", textBox22.Text);
            komut.Parameters.AddWithValue("@garanti_suresi", comboBox3.SelectedItem?.ToString() ?? (object)DBNull.Value);
            komut.Parameters.AddWithValue("@eski_kullanici", textBox14.Text);
            komut.Parameters.AddWithValue("@docking_station_mac_adress", textBox15.Text);
            komut.Parameters.AddWithValue("@docking_station_ip_adress", textBox16.Text);
            komut.Parameters.AddWithValue("@switch", textBox24.Text);
            komut.Parameters.AddWithValue("@port", textBox25.Text);
            komut.Parameters.AddWithValue("@isletim_sistemi", textBox26.Text);
            komut.Parameters.AddWithValue("@virus", textBox27.Text);
            komut.Parameters.AddWithValue("@islemci", textBox28.Text);
            komut.Parameters.AddWithValue("@bellek", textBox33.Text);
            komut.Parameters.AddWithValue("@lva", comboBox6.SelectedItem?.ToString() ?? (object)DBNull.Value);
            komut.Parameters.AddWithValue("@domain", textBox30.Text);
            komut.Parameters.AddWithValue("@office", textBox31.Text);
            komut.Parameters.AddWithValue("@bitlocker", comboBox5.SelectedItem?.ToString() ?? (object)DBNull.Value);
            komut.Parameters.AddWithValue("@zimmet_baslangic", textBox17.Text);
            komut.Parameters.AddWithValue("@zimmet_bitis", textBox18.Text);
            komut.Parameters.AddWithValue("@kiralama_baslangic", textBox19.Text);
            komut.Parameters.AddWithValue("@kiralama_bitis", textBox23.Text);
            komut.Parameters.AddWithValue("@kiralanan_firma", textBox20.Text);
            komut.Parameters.AddWithValue("@durumu", comboBox4.SelectedItem?.ToString() ?? (object)DBNull.Value);
            komut.Parameters.AddWithValue("@aciklama", richTextBox2.Text);

            try
            {
                baglanti.Open(); // veri tabanını açıyoruz
                komut.ExecuteNonQuery(); // verileri kaydetme
                MessageBox.Show("Veri başarıyla eklendi.", "BAŞARILI", MessageBoxButtons.OK , MessageBoxIcon.Information);
            }
            catch (Exception ex) //Hatayı yakalamak için
            {
                MessageBox.Show("Veri eklenirken bir hata oluştu: " + ex.Message);
            }
            finally
            {
                baglanti.Close(); // açık kalması iyi değil, bu yüzden kapatıyoruz
            }

        }



        private void Form1_Load(object sender, EventArgs e)
        {
            string[] kategori = { "Click Share", "Desktop", "Laptop", "Monitör", "N-Com", "Z-Box", "Printer", "Projeksiyon",
        "Shuttle", "Tablet", "TV", "UPS", "Cep Telefonu", "Switch", "Access Point", "Server", "Hand Terminal", "Storage" };

            comboBox1.Items.AddRange(kategori);

            string[] garanti = { "1 Yıl", "2 Yıl", "3 Yıl", "4 Yıl", "5 Yıl", "6 Yıl", "7 Yıl", "8 Yıl", "9 Yıl", "10 Yıl" };
            comboBox3.Items.AddRange(garanti);

            string[] lokasyon = { "YWTT", "YOTG", "YOTK", "YOTT", "YSTTB", "YSTTK", "YKHT", "NURSAN" };
            comboBox2.Items.AddRange(lokasyon);

            string[] durum = { "Aktif", "Pasif", "Hurda", "Arızalı", "Servis", "Servis Bekliyor" };
            comboBox4.Items.AddRange(durum);

            string[] lva = { "Var", "Yok" };
            comboBox5.Items.AddRange(lva);

            string[] bitlocker = { "Var", "Yok" };
            comboBox6.Items.AddRange(bitlocker);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (Kategori form2 = new Kategori())
            {
                if (form2.ShowDialog() == DialogResult.OK)
                {
                    // Form2'den gelen veriyi kullanmak için böyle yaptım
                    string receivedData = form2.Data;
                    comboBox1.Items.Add(receivedData);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Excel'in Arka Planda Açık Olmadığına Emin Olunuz! (Görev Yöneticisinden Kontrol Edin!)", "UYARI", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (result == DialogResult.OK)
            {
                try
                {
                    string excelPath = System.Windows.Forms.Application.StartupPath + "\\envanter.xlsx";
                    bool fileExists = System.IO.File.Exists(excelPath);

                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook;
                    Excel.Worksheet worksheet;

                    if (fileExists)
                    {
                        workbook = excelApp.Workbooks.Open(excelPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        worksheet = (Excel.Worksheet)workbook.Sheets[1];
                    }
                    else
                    {
                        workbook = excelApp.Workbooks.Add();
                        worksheet = (Excel.Worksheet)workbook.Sheets[1];

                        // Başlıkları ekleme (sadece yeni dosya oluşturulduğunda)
                        worksheet.Cells[1, 1] = "Asset No";
                        worksheet.Cells[1, 2] = "Lokasyon";
                        worksheet.Cells[1, 3] = "Yeni Hostname";
                        worksheet.Cells[1, 4] = "Eski Hostname";
                        worksheet.Cells[1, 5] = "Kullanici";
                        worksheet.Cells[1, 6] = "Kategori";
                        worksheet.Cells[1, 7] = "Marka";
                        worksheet.Cells[1, 8] = "Model";
                        worksheet.Cells[1, 9] = "Seri No";
                        worksheet.Cells[1, 10] = "Mac Adres";
                        worksheet.Cells[1, 11] = "Mac Adress 2";
                        worksheet.Cells[1, 12] = "Wireless Mac Adres";
                        worksheet.Cells[1, 13] = "IP No";
                        worksheet.Cells[1, 14] = "Bulunduğu Bölge";
                        worksheet.Cells[1, 15] = "Eski Kullanici";
                        worksheet.Cells[1, 16] = "Zimmet Baslangic";
                        worksheet.Cells[1, 17] = "Zimmet Bitis";
                        worksheet.Cells[1, 18] = "Kiralama Baslangic";
                        worksheet.Cells[1, 19] = "Kiralama Bitis";
                        worksheet.Cells[1, 20] = "Kiralanan Firma";
                        worksheet.Cells[1, 21] = "Docking Station MAC Adress";
                        worksheet.Cells[1, 22] = "Docking Station IP Adress";
                        worksheet.Cells[1, 23] = "Tedarik Firmasi";
                        worksheet.Cells[1, 24] = "Alış Tarihi";
                        worksheet.Cells[1, 25] = "Garanti Süresi";
                        worksheet.Cells[1, 26] = "Switch";
                        worksheet.Cells[1, 27] = "Port";
                        worksheet.Cells[1, 28] = "İşletim Sistemi";
                        worksheet.Cells[1, 29] = "Virüs";
                        worksheet.Cells[1, 30] = "İşlemci";
                        worksheet.Cells[1, 31] = "Bellek";
                        worksheet.Cells[1, 32] = "LVA";
                        worksheet.Cells[1, 33] = "Domain";
                        worksheet.Cells[1, 34] = "Office";
                        worksheet.Cells[1, 35] = "Bitlocker";
                        worksheet.Cells[1, 36] = "Durumu";
                        worksheet.Cells[1, 37] = "Aciklama";
                    }

                    // İlk boş satırı bulma
                    int row = 2;
                    while (((Excel.Range)worksheet.Cells[row, 1]).Value2 != null)
                    {
                        row++;
                    }

                    // Verileri ekleme
                    worksheet.Cells[row, 1] = textBox1.Text;
                    worksheet.Cells[row, 2] = comboBox2.SelectedItem?.ToString() ?? string.Empty;
                    worksheet.Cells[row, 3] = textBox4.Text;
                    worksheet.Cells[row, 4] = textBox3.Text;
                    worksheet.Cells[row, 5] = textBox5.Text;
                    worksheet.Cells[row, 6] = comboBox1.SelectedItem?.ToString() ?? string.Empty;
                    worksheet.Cells[row, 7] = textBox6.Text;
                    worksheet.Cells[row, 8] = textBox7.Text;
                    worksheet.Cells[row, 9] = textBox8.Text;
                    worksheet.Cells[row, 10] = textBox9.Text;
                    worksheet.Cells[row, 11] = textBox10.Text;
                    worksheet.Cells[row, 12] = textBox11.Text;
                    worksheet.Cells[row, 13] = textBox12.Text;
                    worksheet.Cells[row, 14] = textBox13.Text;
                    worksheet.Cells[row, 15] = textBox14.Text;
                    worksheet.Cells[row, 16] = textBox17.Text;
                    worksheet.Cells[row, 17] = textBox18.Text;
                    worksheet.Cells[row, 18] = textBox19.Text;
                    worksheet.Cells[row, 19] = textBox23.Text;
                    worksheet.Cells[row, 20] = textBox20.Text;
                    worksheet.Cells[row, 21] = textBox15.Text;
                    worksheet.Cells[row, 22] = textBox16.Text;
                    worksheet.Cells[row, 23] = textBox21.Text;
                    worksheet.Cells[row, 24] = textBox22.Text;
                    worksheet.Cells[row, 25] = comboBox3.SelectedItem?.ToString() ?? string.Empty;
                    worksheet.Cells[row, 26] = textBox24.Text;
                    worksheet.Cells[row, 27] = textBox25.Text;
                    worksheet.Cells[row, 28] = textBox26.Text;
                    worksheet.Cells[row, 29] = textBox27.Text;
                    worksheet.Cells[row, 30] = textBox28.Text;
                    worksheet.Cells[row, 31] = textBox33.Text;
                    worksheet.Cells[row, 32] = comboBox6.SelectedItem?.ToString() ?? string.Empty;
                    worksheet.Cells[row, 33] = textBox30.Text;
                    worksheet.Cells[row, 34] = textBox31.Text;
                    worksheet.Cells[row, 35] = comboBox5.SelectedItem?.ToString() ?? string.Empty;
                    worksheet.Cells[row, 36] = comboBox4.SelectedItem?.ToString() ?? string.Empty;
                    worksheet.Cells[row, 37] = richTextBox2.Text;

                    // Excel dosyasını kaydetme
                    workbook.SaveAs(excelPath);
                    workbook.Close(false, Type.Missing, Type.Missing);
                    excelApp.Quit();

                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(excelApp);
                    worksheet = null;
                    workbook = null;
                    excelApp = null;
                    GC.Collect();

                    MessageBox.Show("Veriler başarıyla Excel dosyasına kaydedildi.", "BAŞARILI", MessageBoxButtons.OK , MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Veri kaydedilirken bir hata oluştu: ", "HATA!" + MessageBoxIcon.Error + ex.Message);
                }
            }
            else if (result == DialogResult.Cancel)
            {

            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            VeriGoruntule form6 = new VeriGoruntule();
            form6.Show();
            this.Close();
        }
       
        private void button9_Click(object sender, EventArgs e)
        {

            YetkiKontrol form30 = new YetkiKontrol();
            form30.Show();
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            YetkiKontrol2 yetkiKontrol2 = new YetkiKontrol2();
            yetkiKontrol2.Show();
            this.Close();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            // Mevcut tarih ve saat bilgisi
            string currentDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // Outlook'u açarak e-posta gönder
            try
            {
                // E-posta adresi, konu ve içerik
                string subject = "Envanter Kayıt Uygulamasında Bir Hata Bulundu";
                string body = $"Merhaba Geliştirici,\n\n{currentDateTime} tarihinde uygulamanın Ana Sayfa kısmında bir hata yakaladım. Lütfen ilgilenir misiniz?\n";

                // mailto linkini oluştur
                string mailto = $"mailto:emirhanneren24@gmail.com?subject={Uri.EscapeDataString(subject)}&body={Uri.EscapeDataString(body)}";

                // Outlook'u açarak e-posta oluştur
                System.Diagnostics.Process.Start(mailto);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Outlook açılamadı: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            VeriGuncelle form41 = new VeriGuncelle();
            form41.Show();
            this.Hide();
        }
    }
}


//pasif gelecek güncelleme butonu ile başka bir forma (ana sayfadaki güncelleme formuna)
//silinenler dşye bir excel
//EKlenen kategori kalıcı olsun
//silinenlerde hangi tarihte silindi
//Veri filtreleyip excel e kaydetme veri görüntüle de