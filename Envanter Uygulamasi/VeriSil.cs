using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Envanter_Uygulamasi
{
    public partial class VeriSil : Form
    {
        OleDbConnection baglanti;
        OleDbCommand komut;
        OleDbDataAdapter da1;

        private int selectedRowIndex = -1;
        private int selectedColumnIndex = -1;

        public VeriSil()
        {
            InitializeComponent();
            dataGridView1.CellClick += dataGridView1_CellClick;
        }

        private void VeriSil_Load(object sender, EventArgs e)
        {
            envanter();


            string[] kriter = { "Asset_No", "Lokasyon", "Yeni_Hostname", "Eski_Hostname", "Kullanici", "Kategori", "Marka", "Model", "Seri_No", "IP_No", "Bulundugu_Bolge", "Mac_Adres", "Mac_Adress_2", "Wireless_Mac_Adres", "Tedarik_Firmasi", "Alis_Tarihi", "Garanti_Suresi", "Eski_Kullanici", "Docking_Station_MAC_Adress", "Docking_Station_IP_Adress", "Switch", "Port", "Isletim_Sistemi", "Virus", "Islemci", "Bellek", "LVA", "Domain", "Office", "Bitlocker", "Zimmet_Baslangic", "Zimmet_Bitis", "Kiralama_Baslangic", "Kiralama_Bitis", "Kiralanan_Firma", "Durumu", "Aciklama" };
            comboBox1.Items.AddRange(kriter);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AnaSayfa anaSayfa = new AnaSayfa();
            anaSayfa.Show();
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\EXCALIBUR\source\repos\Envanter Uygulamasi\Envanter Uygulamasi\bin\Debug\envanter.xlsx";

            try
            {
                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Dosya açılamadı: {ex.Message}");
            }
        }

        private void envanter()
        {
            //Asıl bağlantı komutumuz
            string baglan, sorgu;
            baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            sorgu = "SELECT * FROM envanter";
            baglanti = new OleDbConnection(baglan); //Bağlantıyı yukarıda tanımladık fakat boş oldupu için burda yukarıdaki kodu execute ettik.
            da1 = new OleDbDataAdapter(sorgu, baglanti); //Geçici olarak da e verileri kaydettik.
            DataSet al = new DataSet(); //dataset ise kalıcı olarak veriyi alır, datagridview veriyi burdan çeker.
            da1.Fill(al, "abc"); //Dataadapter dan veriyi alıp dataset e (da1-----al)
            dataGridView1.DataSource = al.Tables[0];
        }

        private void filtrele(string secilenKriter, string filtreDegeri)
        {
            string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";

            // Sorguyu dinamik olarak oluşturuyoruz
            string sorgu = "SELECT * FROM envanter WHERE " + secilenKriter + " = @filtreDegeri";

            OleDbConnection baglanti = new OleDbConnection(baglan);
            OleDbCommand komut = new OleDbCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@filtreDegeri", filtreDegeri);

            OleDbDataAdapter da1 = new OleDbDataAdapter(komut);
            DataSet al = new DataSet();
            da1.Fill(al, "abc");

            dataGridView1.DataSource = al.Tables[0];
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string secilenKriter = comboBox1.SelectedItem.ToString();
            string filtreDegeri = textBox1.Text;

            // Filtreleme metodunu çağırıyoruz
            filtrele(secilenKriter, filtreDegeri);
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            // DataGridView'de seçili satır olup olmadığını kontrol ediyoruz
            if (dataGridView1.SelectedRows.Count > 0)
            {
                // Eğer seçili satır varsa, silme butonunu etkinleştiriyoruz
                button5.Enabled = true;
            }
            else
            {
                // Eğer seçili satır yoksa, silme butonunu devre dışı bırakıyoruz
                button5.Enabled = false;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Seçili Olan Satır Silinecek. Emin misiniz?", "UYARI", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (result == DialogResult.OK)
            {
                komut = new OleDbCommand("DELETE FROM envanter WHERE kimlik=@kimlik", baglanti);
                komut.Parameters.AddWithValue("@sicil", dataGridView1.CurrentRow.Cells[0].Value); //Dgv deki seçili olan hücrenin ilk hücresinindeki veri kimlik değişkenine eşit oldu.
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();

                envanter();

                MessageBox.Show("Satır Başarıyla Silindi.","BAŞARILI",MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (result == DialogResult.Cancel)
            {
                // İptal işlemleri buraya
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

        private void button8_Click(object sender, EventArgs e)
        {
            // Seçili hücreyi kontrol ediyoruz
            if (selectedRowIndex >= 0 && selectedColumnIndex >= 0)
            {
                // Seçili hücredeki veriyi temizliyoruz
                dataGridView1.Rows[selectedRowIndex].Cells[selectedColumnIndex].Value = string.Empty;

                envanter();
            }
            else
            {
                MessageBox.Show("Lütfen önce bir hücre seçin.");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            envanter();
        }
    }
}
