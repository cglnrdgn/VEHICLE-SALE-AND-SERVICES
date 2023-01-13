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
    public partial class SATIŞ : Form
    {
        public SATIŞ()
        {
            InitializeComponent();
        }

        //sql bağlantı
        static string constring = ("Data Source=LAPTOP-E4FOKIAN\\SQLEXPRESS;Initial Catalog=VehicleSale;Integrated Security=True");
        SqlConnection baglan = new SqlConnection(constring);

        //kayıt ekleme
        public void Kayitlari_getir()
        {
            baglan.Open();
            string getir = "Select * from Sales";
            string where = " WHERE ";

            if (dateTimePicker1.Value != null && dateTimePicker2.Value != null)
            {
                string startdate = dateTimePicker1.Value.ToString("MM-dd-yyyy");
                string enddate = dateTimePicker2.Value.ToString("MM-dd-yyyy");

                where += "  saleDate between '" + startdate + "' and '"+enddate+"' ";
            }
            getir = getir + where + " ORDER BY saleDate ASC";
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
  
            if (UserProperties.userDeleteRole == true)
            {
                string sil = "Delete from Sales where saleID= @id";

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

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        //ana class
        private void SATIŞ_Load(object sender, EventArgs e)
        {
           
        }

        //veri ekleme
        private void ekleme_Click(object sender, EventArgs e)
        {
            try
            {
                if (UserProperties.userAddRole == true)
                {
                    if (baglan.State == ConnectionState.Closed)
                    {
                        baglan.Open();
                        string ekle = "Insert Into Sales(saledate,paymentMethodID,customerID,vehicleID) values" +
                            " (@saledate,@paymentMethodID,@customerID,@vehicleID)";
                        SqlCommand komut = new SqlCommand(ekle, baglan);
                        komut.Parameters.AddWithValue("@saledate", textBox1.Text);
                        komut.Parameters.AddWithValue("@paymentMethodID", textBox9.Text);
                        komut.Parameters.AddWithValue("@customerID", textBox12.Text);
                        komut.Parameters.AddWithValue("@vehicleID", textBox11.Text);

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

            if (UserProperties.userDeleteRole == true)
            {
                string kayitgüncelle = "Update Customers set customerName=@customerName,customersurname=@customersurname,phoneNumber=@phoneNumber,emailAdress=@emailAdress,cityID=@cityID,postCode=@postCode";

                SqlCommand kay = new SqlCommand(kayitgüncelle, baglan);

                kay.Parameters.AddWithValue("@saledate", textBox1.Text);
                kay.Parameters.AddWithValue("@paymentMethodID", textBox9.Text);
                kay.Parameters.AddWithValue("@customerID", textBox12.Text);
                kay.Parameters.AddWithValue("@vehicleID", textBox11.Text);
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
            bulkCopy.DestinationTableName = "Sales";

            bulkCopy.ColumnMappings.Add("saleDate", "saleDate");
            bulkCopy.ColumnMappings.Add("paymentMethodID", "paymentMethodID");
            bulkCopy.ColumnMappings.Add("customerID", "customerID");
            bulkCopy.ColumnMappings.Add("vehicleID", "vehicleID");

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

        //veritabanını geri yükle
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

