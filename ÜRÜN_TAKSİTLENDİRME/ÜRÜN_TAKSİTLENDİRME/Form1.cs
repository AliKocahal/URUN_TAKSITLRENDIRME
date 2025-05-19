using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
namespace ÜRÜN_TAKSİTLENDİRME
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection baglan = new OleDbConnection
           ("Provider=Microsoft.Ace.oledb.12.0; Data Source=taksitler.accdb");

        private void Form1_Load(object sender, EventArgs e)
        {
            for (int i = 1; i <= 12; i++)
            {
                comboBox1.Items.Add(i.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double toplamtutar = Convert.ToDouble(textBox1.Text);
            double taksitsayisi = Convert.ToDouble(comboBox1.SelectedItem);
            DateTime baslangictarihi = dateTimePicker1.Value;
            double taksittutarı = toplamtutar / taksitsayisi;
            DateTime taksittarihi;
            string[] aylar = new string[] { "", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık" };
            string Alanlar = ""; string taksitler = "";
            for (int i = 1; i <= taksitsayisi; i++)
            {
                taksittarihi = baslangictarihi.AddMonths(i);
                Alanlar = Alanlar + aylar[taksittarihi.Month] + ", ";
                taksitler += "'" + taksittutarı.ToString("C") + "', ";
            }
            using (OleDbCommand kaydet = new OleDbCommand("INSERT INTO TABLO1 (" + Alanlar + " TC, adsoyad,toplamborc) VALUES (" + taksitler + "'" + textBox3.Text + "', '" + textBox2.Text + "','" + textBox1.Text.ToString() + "') ", baglan))
            {


                baglan.Open();
                kaydet.ExecuteNonQuery();
                MessageBox.Show(" AYLIK TAKSİT ÖDEMELERİ/nBAŞARIYLA SİSTEME YÜKLENDİ ...");
                baglan.Close();
            }
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

       
    }
}
