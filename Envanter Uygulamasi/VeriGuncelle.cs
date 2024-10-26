using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Envanter_Uygulamasi
{
    public partial class VeriGuncelle : Form
    {
        OleDbConnection baglanti;
        OleDbDataAdapter da1;
        DataSet dataSet; // DataTable yerine DataSet kullanımı
        OleDbCommandBuilder commandBuilder;

        public VeriGuncelle()
        {
            InitializeComponent();
        }

        private void VeriGuncelle_Load(object sender, EventArgs e)
        {
            string baglan, sorgu;
            baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            sorgu = "SELECT * FROM envanter";
            baglanti = new OleDbConnection(baglan);

            da1 = new OleDbDataAdapter(sorgu, baglanti);
            commandBuilder = new OleDbCommandBuilder(da1); // Komut oluşturucu

            dataSet = new DataSet();
            da1.Fill(dataSet, "envanter");
            dataGridView1.DataSource = dataSet.Tables["envanter"];

            dataGridView1.MultiSelect = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                // Güncellemeleri veritabanına uygula
                da1.Update(dataSet, "envanter");
                MessageBox.Show("Değişiklikler başarıyla kaydedildi!", "BAŞARILI", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            YetkiKontrol form30 = new YetkiKontrol();
            form30.Show();
            this.Close();
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
                string body = $"Merhaba Geliştirici,\n\n{currentDateTime} tarihinde uygulamanın Veri Güncelleme kısmında bir hata yakaladım. Lütfen ilgilenir misiniz?\n";

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
