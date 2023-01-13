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

    public partial class LOGIN : Form
    {
        SqlConnection con;
        SqlCommand cmd;
        SqlDataReader reader;

        public LOGIN()
        {
            InitializeComponent();
        }

        //sql bağlantı
        static string constring = ("Data Source=LAPTOP-E4FOKIAN\\SQLEXPRESS;Initial Catalog=VehicleSale;Integrated Security=True");
        SqlConnection baglan=new SqlConnection(constring);

        //şifre yenileme arayüzüne geçiş
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SİFREYENİLE sifre=new SİFREYENİLE();
            sifre.ShowDialog();
        }

        //kullanıcı girişi
        private void button1_Click(object sender, EventArgs e)
        {
            string user =textBox1.Text; 
            string password =textBox2.Text;
            con = new SqlConnection("Data Source=LAPTOP-E4FOKIAN\\SQLEXPRESS;Initial Catalog=VehicleSale;Integrated Security=True");
            cmd = new SqlCommand();
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = "Select userAddRole, userEditRole, userDeleteRole, userIsAdmin from Managers where username='"+ textBox1.Text+ "'And userPassword='" + textBox2.Text + "'";
            reader=cmd.ExecuteReader();
            if(reader.Read())
            {
                
                UserProperties.userAddRole = reader.GetBoolean(0);
                UserProperties.userEditRole = reader.GetBoolean(1);
                UserProperties.userDeleteRole = reader.GetBoolean(2);
                UserProperties.userIsAdmin = reader.GetBoolean(3);

                MessageBox.Show("Tebrikler Giriş Başarili");
                SEÇİM secim = new SEÇİM();
                secim.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Hatali Giriş.Tekrar Deneyiniz");
            }
            con.Close();
        }


        private void LOGIN_Load(object sender, EventArgs e)
        {

        }
    }
}
