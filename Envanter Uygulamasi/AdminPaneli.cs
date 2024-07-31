using ClosedXML.Excel;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Envanter_Uygulamasi
{
    public partial class AdminPaneli : Form
    {
        OleDbConnection baglanti;
        OleDbCommand komut;
        OleDbDataAdapter da1;

        private int selectedRowIndex = -1;
        private int selectedColumnIndex = -1;

        public AdminPaneli()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            kullanici();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            kullanici();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        
        private void AdminPaneli_Load(object sender, EventArgs e)
        {
            //Asıl bağlantı komutumuz
            string baglan, sorgu;
            baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            sorgu = "SELECT * FROM kullanici";
            baglanti = new OleDbConnection(baglan); //Bağlantıyı yukarıda tanımladık fakat boş oldupu için burda yukarıdaki kodu execute ettik.
            da1 = new OleDbDataAdapter(sorgu, baglanti); //Geçici olarak da e verileri kaydettik.
            DataSet al = new DataSet(); //dataset ise kalıcı olarak veriyi alır, datagridview veriyi burdan çeker.
            da1.Fill(al, "abc"); //Dataadapter dan veriyi alıp dataset e (da1-----al)
            dataGridView1.DataSource = al.Tables[0];

            string[] kullanici = { "sicil", "kullaniciadi", "yazakimail"};
            comboBox1.Items.AddRange(kullanici);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            // ComboBox'dan seçilen kriteri ve TextBox'dan filtre değerini alıyoruz
            string secilenKriter = comboBox1.SelectedItem.ToString();
            string filtreDegeri = textBox1.Text;

            // Filtreleme metodunu çağırıyoruz
            kfiltrele(secilenKriter, filtreDegeri);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AnaSayfa anasayfa = new AnaSayfa();
            anasayfa.Show();
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox4.Text != textBox5.Text)
            {
                MessageBox.Show("Şifreler Aynı Değil!", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // Eğer şifreler aynı değilse, işlemi durdur.
            }

            string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";

            // baglanti nesnesini tanımlayıp başlatıyoruz
            using (OleDbConnection baglanti = new OleDbConnection(baglan))
            {
                // Komut nesnesini başlatıp baglanti nesnesine bağlıyoruz
                OleDbCommand komut = new OleDbCommand("INSERT INTO kullanici (sicil, kullaniciadi, sifre, yazakimail, yetki) VALUES (@sicil, @kullaniciadi, @sifre, @yazakimail, @yetki)", baglanti);
                komut.Parameters.AddWithValue("@sicil", textBox2.Text);
                komut.Parameters.AddWithValue("@kullaniciadi", textBox3.Text);
                komut.Parameters.AddWithValue("@sifre", textBox4.Text);
                komut.Parameters.AddWithValue("@yazakimail", textBox6.Text);
                komut.Parameters.AddWithValue("@yetki", checkBox1.Checked);

                try
                {
                    baglanti.Open(); // veri tabanını açıyoruz
                    komut.ExecuteNonQuery(); // verileri kaydetme
                    MessageBox.Show("Veri başarıyla eklendi.");
                }
                catch (Exception ex) //Hatayı yakalamak için
                {
                    MessageBox.Show("Veri eklenirken bir hata oluştu: " + ex.Message);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Seçili Olan Kullanıcı Silinecek. Emin misiniz?", "UYARI", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (result == DialogResult.OK)
            {
                komut = new OleDbCommand("DELETE FROM kullanici WHERE sicil=@sicil", baglanti);
                komut.Parameters.AddWithValue("@sicil", dataGridView1.CurrentRow.Cells[0].Value); //Dgv deki seçili olan hücrenin ilk hücresinindeki veri kimlik değişkenine eşit oldu.
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
            }
            else if (result == DialogResult.Cancel)
            {
                // İptal işlemleri buraya
            }

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                SadeceYetkiliKullanicilariGetir();
                checkBox3.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox3.Checked)
            {
                SadeceYetkisizKullanicilariGetir();
                checkBox2.Checked = false;
            }
        }

        private void kullanici()
        {
            //Asıl bağlantı komutumuz
            string baglan, sorgu;
            baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            sorgu = "SELECT * FROM kullanici";
            baglanti = new OleDbConnection(baglan); //Bağlantıyı yukarıda tanımladık fakat boş oldupu için burda yukarıdaki kodu execute ettik.
            da1 = new OleDbDataAdapter(sorgu, baglanti); //Geçici olarak da e verileri kaydettik.
            DataSet al = new DataSet(); //dataset ise kalıcı olarak veriyi alır, datagridview veriyi burdan çeker.
            da1.Fill(al, "abc"); //Dataadapter dan veriyi alıp dataset e (da1-----al)
            dataGridView1.DataSource = al.Tables[0];
        }

        private void kfiltrele(string secilenFiltre, string KriterDegeri)
        {
            string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";

            // Sorguyu dinamik olarak oluşturuyoruz
            string sorgu = "SELECT * FROM kullanici WHERE " + secilenFiltre + " = @filtreDegeri";

            OleDbConnection baglanti = new OleDbConnection(baglan);
            OleDbCommand komut = new OleDbCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@filtreDegeri", KriterDegeri);

            OleDbDataAdapter da1 = new OleDbDataAdapter(komut);
            DataSet al = new DataSet();
            da1.Fill(al, "abc");

            dataGridView1.DataSource = al.Tables[0];
        }

        private void SadeceYetkiliKullanicilariGetir()
        {
            string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            string sorgu = "SELECT * FROM kullanici WHERE yetki = True"; // 'yetki' alanı True olanlar

            try
            {
                using (OleDbConnection baglanti = new OleDbConnection(baglan))
                {
                    OleDbCommand komut = new OleDbCommand(sorgu, baglanti);
                    OleDbDataAdapter da = new OleDbDataAdapter(komut);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "yetkiliKullanicilar");

                    dataGridView1.DataSource = ds.Tables[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

        private void SadeceYetkisizKullanicilariGetir()
        {
            string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            string sorgu = "SELECT * FROM kullanici WHERE yetki = False"; // 'yetki' alanı False olanlar

            try
            {
                using (OleDbConnection baglanti = new OleDbConnection(baglan))
                {
                    OleDbCommand komut = new OleDbCommand(sorgu, baglanti);
                    OleDbDataAdapter da = new OleDbDataAdapter(komut);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "yetkisizKullanicilar");

                    dataGridView1.DataSource = ds.Tables[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Tüm Kullanıcılar Silinecek! Emin Misiniz?", "UYARI", MessageBoxButtons.OKCancel ,MessageBoxIcon.Error);

            if (result == DialogResult.OK)
            {

            }else if (result == DialogResult.Cancel)
            {

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            VerileriExcelDosyasinaAktar();
        }

        private void VerileriExcelDosyasinaAktar()
        {
            try
            {
                // Veritabanından verileri al
                string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
                string sorgu = "SELECT * FROM kullanici";
                OleDbDataAdapter da = new OleDbDataAdapter(sorgu, baglanti);
                DataTable dt = new DataTable();
                da.Fill(dt);

                // Verileri Excel dosyasına aktar
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dt, "Kullanıcılar");

                    // Excel dosyasını kaydet
                    SaveFileDialog saveFileDialog = new SaveFileDialog
                    {
                        Filter = "Excel Dosyası|*.xlsx",
                        Title = "Excel Dosyasını Kaydet",
                        FileName = "Kullanıcılar.xlsx"
                    };

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        wb.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Veriler başarıyla Excel dosyasına aktarıldı.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            // DataGridView'de seçili satır olup olmadığını kontrol ediyoruz
            if (dataGridView1.SelectedRows.Count > 0)
            {
                // Eğer seçili satır varsa, silme butonunu etkinleştiriyoruz
                button4.Enabled = true;
            }
            else
            {
                // Eğer seçili satır yoksa, silme butonunu devre dışı bırakıyoruz
                button4.Enabled = false;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            // Seçili hücreyi kontrol ediyoruz
            if (selectedRowIndex >= 0 && selectedColumnIndex >= 0)
            {
                // Seçili hücredeki veriyi temizliyoruz
                dataGridView1.Rows[selectedRowIndex].Cells[selectedColumnIndex].Value = string.Empty;

                kullanici();
            }
            else
            {
                MessageBox.Show("Lütfen önce bir hücre seçin.");
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            // Hücre seçildiğinde ilgili hücreyi kontrol ediyoruz
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // Seçili hücrenin Row ve Column indekslerini kaydediyoruz
                selectedRowIndex = e.RowIndex;
                selectedColumnIndex = e.ColumnIndex;
            }
        }
    }
}
