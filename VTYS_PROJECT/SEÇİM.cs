using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VTYS_PROJECT
{
    public partial class SEÇİM : Form
    {
        public SEÇİM()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        //araç butonu seçildiyse
        private void button1_Click(object sender, EventArgs e)
        {
            ARAÇ arac = new ARAÇ();
            arac.Show();
        }

        //servis butonu seçildiyse
        private void button2_Click(object sender, EventArgs e)
        {
            SERVİS servis = new SERVİS();
            servis.Show();
        }

        //müşteri butonu seçildiyse
        private void button3_Click(object sender, EventArgs e)
        {
            MÜŞTERİ musteri = new MÜŞTERİ();
            musteri.Show();
        }

        //satış butonu seçildiyse
        private void button4_Click(object sender, EventArgs e)
        {
            SATIŞ satis = new SATIŞ();
            satis.Show();
        }

        //ana class
        private void SEÇİM_Load(object sender, EventArgs e)
        {

        }
        //yönetici girişi ile kullanıcı ekleme
        private void button5_Click(object sender, EventArgs e)
        {
            if(UserProperties.userIsAdmin == true)
            {
                YÖNETİCİ satis = new YÖNETİCİ();
                satis.Show();
            }
            else
            {
                MessageBox.Show("Yetkili Değilsiniz.");

            }
        }
    }
}
