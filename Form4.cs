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
using MongoDB.Driver;
using MongoDB.Bson;

namespace ChongBuonLauFW
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
            ShowCurrentList();
        }

        string filename = "";
        private void button1_Click(object sender, EventArgs e)
        {
            var collection = DatabaseMongoCollection.GetMongoUserCollection();
            openFileDialog1.ShowDialog();
            filename = openFileDialog1.FileName;
            if (filename == "") return;
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";" + "Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'";
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = dtSchema.Rows[0]["TABLE_NAME"].ToString();

                // Read data from the first sheet
                string query = "SELECT * FROM [" + sheetName + "]";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    dataGridView1.Rows.Clear();
                    dataGridView1.Columns.Clear();

                    dataGridView1.Columns.Add("Họ Tên", "Họ Tên");
                    dataGridView1.Columns.Add("Số giấy tờ", "Số giấy tờ");
                    dataGridView1.Columns.Add("Tên trong hệ thống", "Tên trong hệ thống");
                    dataGridView1.Columns.Add("Số chuyến trong hệ thống", "Số chuyến trong hệ thống");
                    
                    int rowCount = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        rowCount++;
                        if (rowCount == 1)
                        {
                            continue;
                        }
                        string id = row[2].ToString();


                        string name = row[1].ToString();
                        var filter = Builders<Person>.Filter.Eq("IdNum", id + ";");

                        // Find matching documents
                        var documents = collection.Find(filter).ToList();
                        bool hasMatch = documents.Count > 0;
                        var countFlight = 0;
                        string dbname = "-";
                        if (hasMatch)
                        {
                            countFlight = documents[0].FlightList.Count;
                            dbname = documents[0].Name;
                        }
                        DataGridViewRow dataGridViewRow = new DataGridViewRow();
                        dataGridViewRow.CreateCells(dataGridView1, name, id, dbname, countFlight);
                        dataGridView1.Rows.Add(dataGridViewRow);

                    }
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (filename == "")
            {
                System.Windows.Forms.MessageBox.Show("Vui lòng chọn file");
                return;
            }
            var collection = DatabaseMongoCollection.GetDSRRCollection(0);
            var deleteResult = collection.DeleteMany(Builders<BsonDocument>.Filter.Empty);

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string id = row.Cells["Số giấy tờ"].Value?.ToString() + ";";

                if (id == ";") continue;
                var document = new BsonDocument
                {
                    { "IdNum", id }
                };

                collection.InsertOne(document);
            }
            ShowCurrentList();
        }

        private void ShowCurrentList()
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            dataGridView2.Columns.Add("Số giấy tờ", "Số giấy tờ");
            dataGridView2.Columns.Add("Họ Tên", "Họ Tên");
            dataGridView2.Columns.Add("Số chuyến trong hệ thống", "Số chuyến trong hệ thống");
            dataGridView2.Columns.Add("Ghi chú", "Ghi chú");


            var collectionRR = DatabaseMongoCollection.GetDSRRCollection(0);
            var collection = DatabaseMongoCollection.GetMongoUserCollection();
            var documents = collectionRR.Find(Builders<BsonDocument>.Filter.Empty).ToList();
            foreach (var document in documents)
            {
                string id = document["IdNum"].AsString;

                string name = "-";
                int countFlight = 0;
                string note = "";
                var filter = Builders<Person>.Filter.Eq("IdNum", id);
                var person = collection.Find(filter).FirstOrDefault();

                if (person != null)
                {
                    name = person.Name;
                    countFlight = person.FlightList.Count;
                    note = person.Note;
                }

                DataGridViewRow dataGridViewRow = new DataGridViewRow();
                dataGridViewRow.CreateCells(dataGridView2, id, name, countFlight);
                dataGridView2.Rows.Add(dataGridViewRow);
            }
        }
    }
}
