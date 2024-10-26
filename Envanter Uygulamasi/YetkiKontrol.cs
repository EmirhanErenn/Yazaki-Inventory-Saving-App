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
    public partial class YetkiKontrol : Form
    {
        OleDbConnection baglanti;
        OleDbCommand komut;
        OleDbDataAdapter da1;

        public YetkiKontrol()
        {
            InitializeComponent();

            baglanti = new OleDbConnection();
            komut = new OleDbCommand();

            string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            baglanti.ConnectionString = baglan;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sicil = textBox1.Text;
            string username = textBox2.Text; // Kullanıcının girdiği kullanıcı adı
            string password = textBox3.Text; // Kullanıcının girdiği şifre

            bool usernameandpassword = false;
            bool adminyetki = false;

            string sorgu = "SELECT yetki FROM kullanici WHERE kullaniciadi = @kullaniciadi AND sifre = @sifre";

            using (OleDbConnection conn = new OleDbConnection(baglanti.ConnectionString))
            {
                try
                {
                    conn.Open();
                    using (OleDbCommand cmd = new OleDbCommand(sorgu, conn))
                    {
                        cmd.Parameters.AddWithValue("@kullaniciadi", username);
                        cmd.Parameters.AddWithValue("@sifre", password);

                        object result = cmd.ExecuteScalar();

                        if (result != null)
                        {
                            usernameandpassword = true;
                            adminyetki = Convert.ToBoolean(result);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Veritabanı hatası: " + ex.Message);
                }
            }

            if (usernameandpassword)
            {
                if (adminyetki)
                {
                    AdminPaneli form30 = new AdminPaneli();
                    form30.Show();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Yetkiniz Yetersiz!", "HATA!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Kullanıcı Adı veya Şifre Hatalı!", "HATA!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AnaSayfa anaSayfa = new AnaSayfa();
            anaSayfa.Show();
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void YetkiKontrol_Load(object sender, EventArgs e)
        {

        }
    }
}

