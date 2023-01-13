using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace VTYS_PROJECT
{
    public partial class MÜŞTERİ : Form
    {
        public MÜŞTERİ()
        {
            InitializeComponent();
        }

        //sql bağlantı
        static string constring = ("Data Source=LAPTOP-E4FOKIAN\\SQLEXPRESS;Initial Catalog=VehicleSale;Integrated Security=True");
        SqlConnection baglan = new SqlConnection(constring);


        //comboboxlara veri getirme ve filtreleme
        public void Kayitlari_getir()
        {
            baglan.Open();
            string getir = "Select * from Customers";
            string where = " WHERE ";

            if (comboBox1.SelectedItem != null)
            {
                string customerName = comboBox1.SelectedItem.ToString();

                where += " customerName = '" + customerName + "' and ";
            }


            if (comboBox2.SelectedItem != null)
            {
                string customersurname = comboBox2.SelectedItem.ToString();

                where += " customersurname = '" + customersurname + "' and ";
            }


            if (comboBox3.SelectedItem != null)
            {
                string phoneNumber = comboBox3.SelectedItem.ToString();

                where += " phoneNumber = '" + phoneNumber + "' and ";
            }


            if (comboBox4.SelectedItem != null)
            {
                string emailAdress = comboBox4.SelectedItem.ToString();

                where += " emailAdress = '" + emailAdress + "' and";
            }


            if (comboBox5.SelectedItem != null)
            {
                string adress = comboBox5.SelectedItem.ToString();

                where += " adress = '" + adress + "' and";
            }


            if (comboBox6.SelectedItem != null)
            {
                string postCode = comboBox6.SelectedItem.ToString();

                where += " postCode = '" + postCode + "' and";
            }


            where = where.Remove(where.Length - 4);

            getir = getir + where;

            SqlCommand cmd = new SqlCommand(getir, baglan);

            SqlDataAdapter sda = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();
            sda.Fill(dt);
            dataGridView1.DataSource = dt;

            baglan.Close();
        }

        //kayıt silme fonksiyonu
        public void kayit_sil(int ID)
        {
            if(UserProperties.userDeleteRole == true)
            {
                string sil = "Delete from Customers where customerID= @id";

                SqlCommand komut = new SqlCommand(sil, baglan);
                baglan.Open();

                komut.Parameters.AddWithValue("@id", ID);

                komut.ExecuteNonQuery();
                baglan.Close();
            }
            else
            {
                MessageBox.Show("Kayıt Silme Yetkiniz Yok");
            }

        }

        //veri ekleme
        private void ekleme_Click(object sender, EventArgs e)
        {
            try
            {
                if(UserProperties.userAddRole == true)
                {
                    if (baglan.State == ConnectionState.Closed)
                    {
                        baglan.Open();
                        string ekle = "Insert Into Customers (customerName,customersurname,phoneNumber,emailAdress,cityID,postCode) values" +
                            " (@customerName,@customersurname,@phoneNumber,@emailAdress,@cityID,@postCode)";
                        SqlCommand komut = new SqlCommand(ekle, baglan);
                        komut.Parameters.AddWithValue("@customerName", textBox1.Text);
                        komut.Parameters.AddWithValue("@customersurname", textBox2.Text);
                        komut.Parameters.AddWithValue("@phoneNumber", textBox3.Text);
                        komut.Parameters.AddWithValue("@emailAdress", textBox4.Text);
                        komut.Parameters.AddWithValue("@cityID", textBox5.Text);
                        komut.Parameters.AddWithValue("@postCode", textBox6.Text);

                        komut.ExecuteNonQuery();

                        MessageBox.Show("Kayıt İşlemi Başarili");
                    }
                }
                else
                {
                    MessageBox.Show("Kayıt Ekleme Yetkiniz Yok");

                }

            }
            catch (Exception hata)
            {
                MessageBox.Show("Kayıt İşlemi Hatali" + hata.Message);
            }

        }

        //veri silme
        private void silme_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow drow in dataGridView1.SelectedRows)
            {
                int ID = Convert.ToInt32(drow.Cells[0].Value);
                kayit_sil(ID);
                Kayitlari_getir();
            }
        }

        //veri güncelleme
        int i = 0;
        private void güncelleme_Click(object sender, EventArgs e)
        {
            if(UserProperties.userDeleteRole == true)
            {
                string kayitgüncelle = "Update Customers set customerName=@customerName,customersurname=@customersurname,phoneNumber=@phoneNumber,emailAdress=@emailAdress,cityID=@cityID,postCode=@postCode";

                SqlCommand kay = new SqlCommand(kayitgüncelle, baglan);

                kay.Parameters.AddWithValue("@customerName", textBox1.Text);
                kay.Parameters.AddWithValue("@customersurname", textBox2.Text);
                kay.Parameters.AddWithValue("@phoneNumber", textBox3.Text);
                kay.Parameters.AddWithValue("@emailAdress", textBox4.Text);
                kay.Parameters.AddWithValue("@cityID", textBox5.Text);
                kay.Parameters.AddWithValue("@postCode", textBox6.Text);
                kay.Parameters.AddWithValue("id", dataGridView1.Rows[i].Cells[0].Value);

                baglan.Close();
                Kayitlari_getir();
            }
            else
            {
                MessageBox.Show("Kayıt Güncelleme Yetkiniz Yok");

            }
           
        }

        //veri listeleme
        private void listeleme_Click(object sender, EventArgs e)
        {
            Kayitlari_getir();
        }

        //ana class
        private void MÜŞTERİ_Load(object sender, EventArgs e)
        {
            SqlDataReader reader;
            baglan.Open();
            string getir = "Select * from Customers";

            SqlCommand cmd = new SqlCommand(getir, baglan);

            reader = cmd.ExecuteReader();

            SqlDataAdapter sda = new SqlDataAdapter(cmd);

            while (reader.Read())
            {
                comboBox1.Items.Add(reader[1].ToString());
                comboBox2.Items.Add(reader[2].ToString());
                comboBox3.Items.Add(reader[3].ToString());
                comboBox4.Items.Add(reader[4].ToString());
                comboBox5.Items.Add(reader[5].ToString());
                comboBox6.Items.Add(reader[6].ToString());
            }
            baglan.Close();
        }

        //dışa aktar excel'e
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dialog = new DialogResult();
                dialog = MessageBox.Show("Bu işlem, veri yoğunluğuna göre uzun sürebilir. Devam etmek istiyor musunuz?", "EXCEL'E AKTARMA", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialog == DialogResult.Yes)
                {
                    Microsoft.Office.Interop.Excel.Application uyg = new Microsoft.Office.Interop.Excel.Application();
                    uyg.Visible = true;
                    Microsoft.Office.Interop.Excel.Workbook kitap = uyg.Workbooks.Add(System.Reflection.Missing.Value);
                    Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)kitap.Sheets[1];
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[1, i + 1];
                        myRange.Value2 = dataGridView1.Columns[i].HeaderText;
                    }

                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Rows.Count; j++)
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                            myRange.Value2 = dataGridView1[i, j].Value;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("İŞLEM İPTAL EDİLDİ.", "İşlem Sonucu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception)
            {

                MessageBox.Show("İŞLEM TAMAMLANMADAN EXCEL PENCERESİNİ KAPATTINIZ.", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //içe aktar excel'den
        private void button6_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;
            OpenFileDialog OpenFile = new OpenFileDialog();

            OpenFile.Filter = "Excel Files|*.xlsx*";
            OpenFile.Title = "DOSYA SEÇİNİZ";
            OpenFile.FilterIndex = 2;
            OpenFile.RestoreDirectory = true;

            if (OpenFile.ShowDialog() == DialogResult.OK)
            {
                filePath = OpenFile.FileName;
            }

            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", filePath);
            OleDbConnection Econ = new OleDbConnection(constr);
            string Query = string.Format("Select  * FROM [{0}]", "Sayfa1$");
            OleDbCommand Ecom = new OleDbCommand(Query, Econ);
            Econ.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
            Econ.Close();
            oda.Fill(ds);

            DataTable Exceldt = ds.Tables[0];

            SqlConnection sqlConnection = new SqlConnection("Data Source=LAPTOP-E4FOKIAN\\SQLEXPRESS;Initial Catalog=VehicleSale;Integrated Security=True");
            SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConnection);
            bulkCopy.DestinationTableName = "Customers";

            bulkCopy.ColumnMappings.Add("customerName", "customerName");
            bulkCopy.ColumnMappings.Add("customersurname", "customersurname");
            bulkCopy.ColumnMappings.Add("phoneNumber", "phoneNumber");
            bulkCopy.ColumnMappings.Add("emailAdress", "emailAdress");
            bulkCopy.ColumnMappings.Add("adress", "adress");
            bulkCopy.ColumnMappings.Add("cityID", "cityID");
            bulkCopy.ColumnMappings.Add("postCode", "postCode");
           

            sqlConnection.Open();
            bulkCopy.WriteToServer(Exceldt);
            sqlConnection.Close();

            MessageBox.Show("Yükleme başarılı");
        }

        //veritabanını yedekle
        private void button2_Click(object sender, EventArgs e)
        {
            var connectionString = "Data Source=LAPTOP-E4FOKIAN\\\\SQLEXPRESS;Initial Catalog=VehicleSale;Integrated Security=True";

            var backupFolder = "c:\\Backup\\";


            var backupFileName = String.Format("{0}{1}-{2}.bak",
                backupFolder, "VehicleSale",
                DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss"));

            var query = String.Format("BACKUP DATABASE {0} TO DISK='{1}'",
                "VehicleSale", backupFileName);

            using (var command = new SqlCommand(query, baglan))
            {
                baglan.Open();
                command.ExecuteNonQuery();
                baglan.Close();

            }
            MessageBox.Show("VERİTABANI YEDEĞİ BAŞARIYLA ALINDI. ( " + backupFileName + " )");
        }

        //veritabanını yedekten döndür
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                string constring = ("Data Source=LAPTOP-E4FOKIAN\\SQLEXPRESS;Initial Catalog=master;Integrated Security=True");
                SqlConnection baglan = new SqlConnection(constring);

                var filePath = string.Empty;
                OpenFileDialog OpenFile = new OpenFileDialog();
                OpenFile.Filter = "Backup File |*.bak";

                OpenFile.RestoreDirectory = true;

                if (OpenFile.ShowDialog() == DialogResult.OK)
                {
                    string sql = "IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = 'VehicleSale')";
                    sql += "ALTER DATABASE VehicleSale SET SINGLE_USER WITH ROLLBACK IMMEDIATE " +
                        "DROP DATABASE VehicleSale RESTORE DATABASE VehicleSale FROM DISK = '" + OpenFile.FileName + "' " +
                        "ALTER DATABASE VehicleSale SET  MULTI_USER WITH ROLLBACK IMMEDIATE";

                    SqlCommand command = new SqlCommand(sql, baglan);

                    baglan.Open();
                    command.ExecuteNonQuery();

                    MessageBox.Show("VERİTABANINIZ YEDEKTEN DÖNDÜRÜLDÜ");
                    baglan.Close();
                    baglan.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}


