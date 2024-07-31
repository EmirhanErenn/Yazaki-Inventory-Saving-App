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
    public partial class VeriGoruntule : Form
    {

        OleDbConnection baglanti;
        OleDbCommand komut;
        OleDbDataAdapter da1;

        public VeriGoruntule()
        {
            InitializeComponent();
        }

        private void VeriGoruntule_Load(object sender, EventArgs e)
        {
            string[] kriter = { "Asset_No", "Lokasyon", "Yeni_Hostname", "Eski_Hostname", "Kullanici", "Kategori", "Marka", "Model", "Seri_No", "IP_No", "Bulundugu_Bolge", "Mac_Adres", "Mac_Adress_2", "Wireless_Mac_Adres", "Tedarik_Firmasi", "Alis_Tarihi", "Garanti_Suresi", "Eski_Kullanici", "Docking_Station_MAC_Adress", "Docking_Station_IP_Adress", "Switch", "Port", "Isletim_Sistemi", "Virus", "Islemci", "Bellek", "LVA", "Domain", "Office", "Bitlocker", "Zimmet_Baslangic", "Zimmet_Bitis", "Kiralama_Baslangic", "Kiralama_Bitis", "Kiralanan_Firma", "Durumu", "Aciklama" };
            comboBox1.Items.AddRange(kriter);

            envanter();
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

        private void button5_Click(object sender, EventArgs e)
        {
            envanter();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // ComboBox'dan seçilen kriteri ve TextBox'dan filtre değerini alıyoruz
            string secilenKriter = comboBox1.SelectedItem.ToString();
            string filtreDegeri = textBox1.Text;

            // Filtreleme metodunu çağırıyoruz
            filtrele(secilenKriter, filtreDegeri);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            envanter();
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
    }
}
