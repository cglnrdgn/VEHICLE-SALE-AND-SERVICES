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
    public partial class YÖNETİCİ : Form
    {
        public YÖNETİCİ()
        {
            InitializeComponent();
        }

        //sql bağlantı
        static string constring = ("Data Source=LAPTOP-E4FOKIAN\\SQLEXPRESS;Initial Catalog=VehicleSale;Integrated Security=True");
        SqlConnection baglan = new SqlConnection(constring);

        //kullanıcı ekleme ve yetkilendirme
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (baglan.State == ConnectionState.Closed)
                {
                    baglan.Open();
                    string ekle = "Insert Into Managers (userNameSurname,username,userPassword,userEmail,userPhone,userAddRole, userEditRole, userDeleteRole,userIsAdmin) values (@userNameSurname,@username,@userPassword,@userEmail,@userPhone,@userAddRole,@userEditRole, @userDeleteRole,@userIsAdmin)";
                    SqlCommand komut = new SqlCommand(ekle, baglan);
                    komut.Parameters.AddWithValue("@userNameSurname", textBox3.Text);
                    komut.Parameters.AddWithValue("@username", textBox4.Text);
                    komut.Parameters.AddWithValue("@userPassword", textBox5.Text);
                    komut.Parameters.AddWithValue("@userEmail", textBox6.Text);
                    komut.Parameters.AddWithValue("@userPhone", textBox7.Text);

                    komut.Parameters.AddWithValue("@userAddRole", checkBox1.Checked);
                    komut.Parameters.AddWithValue("@userEditRole", checkBox2.Checked);
                    komut.Parameters.AddWithValue("@userDeleteRole", checkBox3.Checked);
                    komut.Parameters.AddWithValue("@userIsAdmin", checkBox4.Checked);

                    komut.ExecuteNonQuery();
                    baglan.Close();
                    MessageBox.Show("Kayıt İşlemi Başarili");
                }
            }
            catch (Exception hata)
            {
                MessageBox.Show("Kayıt İşlemi Hatali" + hata.Message);
            }
        }
    }
}
