using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Envanter_Uygulamasi
{
    public partial class Uyarilar : Form
    {
        public Uyarilar()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AnaSayfa form1 = new AnaSayfa();
            form1.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Çok Yakında...", "Yazaki Envanter Uygulaması", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
        }

        private void Uyarilar_Load(object sender, EventArgs e)
        {

        }
    }
}
