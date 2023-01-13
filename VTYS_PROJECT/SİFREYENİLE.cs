using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VTYS_PROJECT
{
    public partial class SİFREYENİLE : Form
    {
        public SİFREYENİLE()
        {
            InitializeComponent();
        }

        //sql bağlantı
        static string constring = ("Data Source=LAPTOP-E4FOKIAN\\SQLEXPRESS;Initial Catalog=VehicleSale;Integrated Security=True");
        SqlConnection baglan = new SqlConnection(constring);

        //şifre güncelleme
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (baglan.State == ConnectionState.Closed)
                {
                    baglan.Open();
                    string kayitgüncelle = "Update Managers set userNameSurname=@userNameSurname,username=@username,userPassword=@userPassword,userEmail=@userEmail,userPhone=@userPhone";

                    SqlCommand kay = new SqlCommand(kayitgüncelle, baglan);

                    kay.Parameters.AddWithValue("@userNameSurname", textBox3.Text);
                    kay.Parameters.AddWithValue("@username", textBox4.Text);
                    kay.Parameters.AddWithValue("@userPassword", textBox5.Text);
                    kay.Parameters.AddWithValue("@userEmail", textBox6.Text);
                    kay.Parameters.AddWithValue("@userPhone", textBox7.Text);

                    kay.ExecuteNonQuery();

                    MessageBox.Show("Şifre Güncellendi");
                }
            }
            catch (Exception hata)
            {
                MessageBox.Show("Şifre Güncelleme Hatali" + hata.Message);
            }
        }

        private void SİFREYENİLE_Load(object sender, EventArgs e)
        {

        }
    }
}
