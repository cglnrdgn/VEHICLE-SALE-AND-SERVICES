using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace VTYS_PROJECT
{
    public partial class ARAÇ : Form
    {
        public ARAÇ()
        {
            InitializeComponent();
        }

        //sql bağlantısı
        static string constring = ("Data Source=LAPTOP-E4FOKIAN\\SQLEXPRESS;Initial Catalog=VehicleSale;Integrated Security=True");
        SqlConnection baglan = new SqlConnection(constring);

        //combobox'lara kayıt getirme ve filtreleme
        public void Kayitlari_getir()
        {
            baglan.Open();
            string getir = "Select * from Vehicles";
            string where = " WHERE ";

            if (comboBox1.SelectedItem != null)
            {
                string modelid = comboBox1.SelectedItem.ToString();

                where += " modelID = '" + modelid + "' and ";
            }


            if (comboBox2.SelectedItem != null)
            {
                string brandid = comboBox2.SelectedItem.ToString();

                where += " brandID = '" + brandid + "' and ";
            }


            if (comboBox3.SelectedItem != null)
            {
                string price = comboBox3.SelectedItem.ToString();

                where += " price = '" + price + "' and ";
            }


            if (comboBox4.SelectedItem != null)
            {
                string fuelTypeID = comboBox4.SelectedItem.ToString();

                where += " fuelTypeID = '" + fuelTypeID + "' and";
            }


            if (comboBox5.SelectedItem != null)
            {
                string colorID = comboBox5.SelectedItem.ToString();

                where += " colorID = '" + colorID + "' and";
            }


            if (comboBox6.SelectedItem != null)
            {
                string gearBoxTypeID = comboBox6.SelectedItem.ToString();

                where += " gearBoxTypeID = '" + gearBoxTypeID + "' and";
            }

            if (comboBox7.SelectedItem != null)
            {
                string vehicleClassID = comboBox7.SelectedItem.ToString();

                where += " vehicleClassID = '" + vehicleClassID + "' and";
            }


            if (comboBox8.SelectedItem != null)
            {
                string vehicleYear = comboBox8.SelectedItem.ToString();

                where += " vehicleYear = '" + vehicleYear + "' and";
            }


            if (comboBox9.SelectedItem != null)
            {
                string licansePlate = comboBox9.SelectedItem.ToString();

                where += " licansePlate = '" + licansePlate + "' and";
            }


            if (comboBox10.SelectedItem != null)
            {
                string vin = comboBox10.SelectedItem.ToString();

                where += " vin = '" + vin + "' and";
            }


            where = where.Remove(where.Length - 4);

            getir = getir + where;
            
            SqlCommand cmd = new SqlCommand(getir,baglan);

            SqlDataAdapter sda=new SqlDataAdapter(cmd);

            DataTable dt=new DataTable();
            sda.Fill(dt);
            dataGridView1.DataSource=dt;

            baglan.Close();
        }

        //kayıt silme fonksiyonu
        public void kayit_sil(int ID)
        {
            if(UserProperties.userDeleteRole == true)
            {
                string sil = "Delete from Vehicles where vehicleID= @id";

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

        //ana class
        private void ARAÇ_Load(object sender, EventArgs e)
        {
            SqlDataReader reader;
            baglan.Open();
            string getir = "Select * from Vehicles";

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
                comboBox7.Items.Add(reader[7].ToString());
                comboBox8.Items.Add(reader[8].ToString());
                comboBox9.Items.Add(reader[9].ToString());
                comboBox10.Items.Add(reader[10].ToString());
            }
            dataGridView1.DataSource = sda;
            baglan.Close();
        }
        
        // veri ekle
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if(UserProperties.userAddRole == true)
                {
                    if (baglan.State == ConnectionState.Closed)
                    {
                        baglan.Open();
                        string ekle = "Insert Into Vehicles (modelID,brandID,price,fuelTypeID,colorID,gearBoxTypeID,vehicleClassID,vehicleYear,licansePlate,vin)" +
                            " values (@modelID,@brandID,@price,@fuelTypeID,@colorID,@gearBoxTypeID,@vehicleClassID,@vehicleYear,@licansePlate,@vin)";
                        SqlCommand komut = new SqlCommand(ekle, baglan);
                        komut.Parameters.AddWithValue("@modelID", textBox1.Text);
                        komut.Parameters.AddWithValue("@brandID", textBox2.Text);
                        komut.Parameters.AddWithValue("@price", textBox3.Text);
                        komut.Parameters.AddWithValue("@fuelTypeID", textBox4.Text);
                        komut.Parameters.AddWithValue("@colorID", textBox5.Text);
                        komut.Parameters.AddWithValue("@gearBoxTypeID", textBox6.Text);
                        komut.Parameters.AddWithValue("@vehicleClassID", textBox8.Text);
                        komut.Parameters.AddWithValue("@vehicleYear", textBox9.Text);
                        komut.Parameters.AddWithValue("@licansePlate", textBox11.Text);
                        komut.Parameters.AddWithValue("@vin", textBox12.Text);

                        komut.ExecuteNonQuery();

                        MessageBox.Show("Ekleme İşlemi Başarili");
                    }

                }
                else
                {
                    MessageBox.Show("Kayıt Ekleme Yetkiniz Yok");

                }
            }
            catch (Exception hata)
            {
                MessageBox.Show("Ekleme İşlemi Başarisiz" + hata.Message);
            }

        }

        // veri sil
        private void button4_Click(object sender, EventArgs e)
        {
            foreach(DataGridViewRow drow in dataGridView1.SelectedRows)
            {
                int ID = Convert.ToInt32(drow.Cells[0].Value);
                kayit_sil(ID);
                Kayitlari_getir();
            }
        }

        // veri güncelle

        int i = 0;
        private void button5_Click(object sender, EventArgs e)
        {
            if(UserProperties.userEditRole == true)
            {
                string kayitgüncelle = "Update Vehicles set modelID= @modelID,brandID= @brandID" +
                ",price= @price,fuelTypeID= @fuelTypeID,colorID= @colorID,gearBoxTypeID= @gearBoxTypeID" +
                ",vehicleClassID= @vehicleClassID,vehicleYear= @vehicleYear,licansePlate= @licansePlate,vin= @vin";

                SqlCommand kay = new SqlCommand(kayitgüncelle, baglan);

                kay.Parameters.AddWithValue("@modelID", textBox1.Text);
                kay.Parameters.AddWithValue("@brandID", textBox2.Text);
                kay.Parameters.AddWithValue("@price", textBox3.Text);
                kay.Parameters.AddWithValue("@fuelTypeID", textBox4.Text);
                kay.Parameters.AddWithValue("@colorID", textBox5.Text);
                kay.Parameters.AddWithValue("@gearBoxTypeID", textBox6.Text);
                kay.Parameters.AddWithValue("@vehicleClassID", textBox8.Text);
                kay.Parameters.AddWithValue("@vehicleYear", textBox9.Text);
                kay.Parameters.AddWithValue("@licansePlate", textBox11.Text);
                kay.Parameters.AddWithValue("@vin", textBox12.Text);
                kay.Parameters.AddWithValue("id", dataGridView1.Rows[i].Cells[0].Value);

                baglan.Close();
                Kayitlari_getir();
            }
            else
            {
                MessageBox.Show("Kayıt Güncelleme Yetkiniz Yok");
            }
            
        }
        //veri listele
        private void listeleme_Click(object sender, EventArgs e)
        {
            Kayitlari_getir();
        }


        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        //Dışa aktarma excel'e
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

        //İçe aktarma excel'den
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
            bulkCopy.DestinationTableName = "Vehicles";

            bulkCopy.ColumnMappings.Add("modelID", "modelID");
            bulkCopy.ColumnMappings.Add("brandID", "brandID");
            bulkCopy.ColumnMappings.Add("price", "price");
            bulkCopy.ColumnMappings.Add("fuelTypeID", "fuelTypeID");
            bulkCopy.ColumnMappings.Add("colorID", "colorID");
            bulkCopy.ColumnMappings.Add("gearboxTypeID", "gearboxTypeID");
            bulkCopy.ColumnMappings.Add("vehicleClassID", "vehicleClassID");
            bulkCopy.ColumnMappings.Add("vehicleYear", "vehicleYear");
            bulkCopy.ColumnMappings.Add("licansePlate", "licansePlate");
            bulkCopy.ColumnMappings.Add("vin", "vin");

            sqlConnection.Open();
            bulkCopy.WriteToServer(Exceldt);  
            sqlConnection.Close();

            MessageBox.Show("YÜKLEME BAŞARILI");

        }

        //veritabanı yedekleme
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
            MessageBox.Show("VERİTABANI YEDEĞİ BAŞARIYLA ALINDI. ( "+ backupFileName + " )");
        }

        //veritabanı yedekten döndürme
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

