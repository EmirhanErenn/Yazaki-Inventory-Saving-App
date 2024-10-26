using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Envanter_Uygulamasi
{
    public partial class KullaniciGuncelle : Form
    {
        OleDbConnection baglanti;
        OleDbDataAdapter da1;
        DataSet dataSet; // DataTable yerine DataSet kullanımı
        OleDbCommandBuilder commandBuilder;

        public KullaniciGuncelle()
        {
            InitializeComponent();
        }

        private void KullaniciGuncelle_Load(object sender, EventArgs e)
        {
            string baglan, sorgu;
            baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            sorgu = "SELECT * FROM kullanici";
            baglanti = new OleDbConnection(baglan);

            da1 = new OleDbDataAdapter(sorgu, baglanti);
            commandBuilder = new OleDbCommandBuilder(da1); // Komut oluşturucu

            dataSet = new DataSet();
            da1.Fill(dataSet, "kullanici");
            dataGridView1.DataSource = dataSet.Tables["kullanici"];

            dataGridView1.MultiSelect = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                // Güncellemeleri veritabanına uygula
                da1.Update(dataSet, "kullanici");
                MessageBox.Show("Değişiklikler başarıyla kaydedildi!","BAŞARILI",MessageBoxButtons.OK,MessageBoxIcon.Information);

                AdminPaneli adminPaneli = new AdminPaneli();
                adminPaneli.Show();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            AnaSayfa form = new AnaSayfa();
            form.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AdminPaneli form2 = new AdminPaneli();
            form2.Show();
            this.Close();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            // Seçim değişiklikleri için gerekli işlemler buraya eklenebilir.
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Mevcut tarih ve saat bilgisi
            string currentDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // Outlook'u açarak e-posta gönder
            try
            {
                // E-posta adresi, konu ve içerik
                string subject = "Envanter Kayıt Uygulamasında Bir Hata Bulundu";
                string body = $"Merhaba Geliştirici,\n\n{currentDateTime} tarihinde uygulamanın Admin Paneli kısmında bir hata yakaladım. Lütfen ilgilenir misiniz?\n";

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
    }
}
