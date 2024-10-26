using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection.Emit;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace Envanter_Uygulamasi
{
    public partial class GirisYap : Form
    {
        OleDbConnection baglanti;
        OleDbCommand komut;
        OleDbDataAdapter da1;

        public GirisYap()
        {
            InitializeComponent();


            baglanti = new OleDbConnection();
            komut = new OleDbCommand();
            da1 = new OleDbDataAdapter();

            string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            baglanti.ConnectionString = baglan;          
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            textBox1.Focus();
            label6.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Login(textBox1.Text,textBox2.Text, textBox3.Text);
        }

        private bool Login(string sicil, string kullaniciadi, string sifre)
        {

            string baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.Windows.Forms.Application.StartupPath + "\\envanter.mdb";
            baglanti.ConnectionString = baglan;

            string sorgu = "SELECT COUNT(*) FROM kullanici WHERE sicil = @sicil AND kullaniciadi = @kullaniciadi AND sifre = @sifre";

            try
            {
                using (var baglanti = new OleDbConnection(baglan))
                {
                    using (var komut = new OleDbCommand(sorgu, baglanti))
                    {
                        komut.Parameters.AddWithValue("@sicil", sicil);
                        komut.Parameters.AddWithValue("@kullaniciadi", kullaniciadi);
                        komut.Parameters.AddWithValue("@sifre", sifre);
                        

                        baglanti.Open();

                        // Sonuç null olabilir, bu nedenle kontrol ediyoruz.
                        int userCount = (int)komut.ExecuteScalar();

                        if (userCount > 0)
                        {
                            AnaSayfa form1 = new AnaSayfa();
                            form1.Show();
                            this.Hide();
                            return true;
                        }
                        else
                        {
                            MessageBox.Show("Kullanıcı Adı veya Şifre Hatalı!", "UYARI",MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            textBox1.Clear();
                            textBox2.Clear();
                            textBox3.Clear();
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (label6 != null)
                {
                    label6.Visible = true;
                    label6.Text = "Bir hata oluştu: " + ex.Message;
                }
                return false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SifremiUnuttum form5 = new SifremiUnuttum();
            form5.Show();
            this.Hide();
        }
    }
}
