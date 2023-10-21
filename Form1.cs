using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ChongBuonLauFW
{
    public partial class Form1 : Form
    {
        string[] filenames;
        private BackgroundWorker backgroundWorker1 = new BackgroundWorker();


        public Form1()
        {
            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.ProgressChanged += BackgroundWorker1_ProgressChanged;
            backgroundWorker1.DoWork += BackgroundWorker1_DoWork;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorker1_RunWorkerCompleted;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            string selectedPath = folderBrowserDialog1.SelectedPath;
            label1.Text = "Folder: " + selectedPath;
            if (selectedPath == "") return;
            filenames = Directory.GetFiles(selectedPath, "*.xlsx", SearchOption.AllDirectories);
            var filename = filenames[0];
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";" + "Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'";
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                OleDbDataAdapter objDA = new OleDbDataAdapter("SELECT * from [Hành khách$]", conn);
                OleDbDataReader reader = objDA.SelectCommand.ExecuteReader();

                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                dataGridView1.Columns.Add("Ngày bay", "Ngày bay");
                dataGridView1.Columns["Ngày bay"].DefaultCellStyle.Format = "dd/MM/yyyy";
                dataGridView1.Columns.Add("Chuyến bay", "Chuyến bay");
                dataGridView1.Columns.Add("Số ghế", "Số ghế");
                dataGridView1.Columns.Add("Họ tên", "Họ tên");
                dataGridView1.Columns.Add("Giới tính", "Giới tính");
                dataGridView1.Columns.Add("Quốc tịch", "Quốc tịch");
                dataGridView1.Columns.Add("Ngày sinh", "Ngày sinh");
                dataGridView1.Columns.Add("Loại giấy tờ", "Loại giấy tờ");
                dataGridView1.Columns.Add("Số giấy tờ", "Số giấy tờ");
                dataGridView1.Columns.Add("Nơi cấp", "Nơi cấp");
                dataGridView1.Columns.Add("Nơi đi", "Nơi đi");
                dataGridView1.Columns.Add("Nơi đến", "Nơi đến");
                dataGridView1.Columns.Add("Hành lý", "Hành lý");
                int rowCount = 0;
                while (reader.Read() && rowCount < 1000)
                {
                    string id = reader[8].ToString();
                    if (!id.EndsWith(";")) continue;
                    rowCount++;

                    string flightdate = reader[0].ToString();
                    DateTime flightdt;
                    DateTime.TryParse(flightdate, new CultureInfo("en-GB"), DateTimeStyles.None, out flightdt);


                    string des = reader[12].ToString();
                    string sta = reader[11].ToString();
                    string luggage = reader[14].ToString();
                    string seat = reader[2].ToString();
                    string flightNumber = reader[1].ToString();

                    string name = reader[3].ToString();
                    string sex = reader[4].ToString();
                    string nat = reader[5].ToString();
                    string dob = reader[6].ToString();
                    string idtype = reader[7].ToString();
                    string idprov = reader[9].ToString();

                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dataGridView1, flightdate, flightNumber, seat, name, sex, nat, dob, idtype, id, idprov, sta, des, luggage);
                    dataGridView1.Rows.Add(row);
                }

                reader.Close();
                conn.Close();
            }
        }

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var collection = MongoUserCollection.GetMongoUserCollection();
            var bulkOps = new List<WriteModel<Person>>();

            for (int fileIndex = 0; fileIndex < filenames.Count(); fileIndex++)
            {
                var filename = filenames[fileIndex];
                string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";" + "Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'";
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    OleDbCommand countRow = new OleDbCommand("SELECT count(*) FROM [Hành khách$]", conn);
                    int rowCount = (Int32)countRow.ExecuteScalar();

                    OleDbDataAdapter objDA = new OleDbDataAdapter("SELECT * from [Hành khách$]", conn);
                    OleDbDataReader reader = objDA.SelectCommand.ExecuteReader();
                    var curRow = 0;
                    while (reader.Read())
                    {
                        curRow++;

                        string id = reader[8].ToString();
                        if (!id.EndsWith(";")) continue;

                        string flightdate = reader[0].ToString();
                        DateTime flightdt;
                        DateTime.TryParse(flightdate, new CultureInfo("en-GB"), DateTimeStyles.None, out flightdt);
                        flightdt = new DateTime(flightdt.Year, flightdt.Month, flightdt.Day, 0, 0, 0, DateTimeKind.Utc).AddHours(-7);


                        string des = reader[12].ToString();
                        string sta = reader[11].ToString();
                        string luggage = reader[14].ToString();
                        string seat = reader[2].ToString();
                        string flightNumber = reader[1].ToString();

                        string name = reader[3].ToString();
                        string sex = reader[4].ToString();
                        string nat = reader[5].ToString();
                        string dob = reader[6].ToString();
                        string idtype = reader[7].ToString();
                        string idprov = reader[9].ToString();

                        var filter = Builders<Person>.Filter.Eq(p => p.IdNum, id);

                        var update = Builders<Person>.Update
                            .Set(p => p.Name, name)
                            .Set(p => p.Sex, sex)
                            .Set(p => p.Nationality, nat)
                            .Set(p => p.DOB, dob)
                            .Set(p => p.IdType, idtype)
                            .Set(p => p.IdProv, idprov)
                            .AddToSet(p => p.FlightList, new Flight
                            {
                                Origin = sta,
                                Destination = des,
                                Luggage = luggage,
                                Date = flightdt,
                                Seat = seat,
                                FlightNumber = flightNumber
                            });
                            

                        var updateOneModel = new UpdateOneModel<Person>(filter, update)
                        {
                            IsUpsert = true
                        };

                        bulkOps.Add(updateOneModel);

                        if (bulkOps.Count == 1000)
                        {
                            var bulkWriteResult = collection.BulkWriteAsync(bulkOps);
                            bulkOps.Clear();
                        }

                        int progress = fileIndex * 100 / filenames.Count() + curRow * 100 / rowCount / filenames.Count();
                        backgroundWorker1.ReportProgress(progress);
                    }
                    if (bulkOps.Count > 0)
                    {
                        var bulkWriteResult = collection.BulkWrite(bulkOps);
                        bulkOps.Clear();
                    }

                    reader.Close();
                    conn.Close();
                }
            }
            if (bulkOps.Count > 0)
            {
                var bulkWriteResult = collection.BulkWrite(bulkOps);
                bulkOps.Clear();
            }
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Update the progress bar on the UI thread
            progressBar1.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button2.Enabled = true;
            progressBar1.Value = 0;
            if (e.Error != null)
            {
                MessageBox.Show("An error occurred during adding data: " + e.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.Cancelled)
            {
                MessageBox.Show("Data adding was cancelled.", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Data adding completed.", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            this.Hide();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (!backgroundWorker1.IsBusy)
            {
                button2.Enabled = false;
                backgroundWorker1.RunWorkerAsync();
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
