using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChongBuonLauFW
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
            ShowCurrentList();
        }

        private void UpdateList()
        {
            var collection = DatabaseMongoCollection.GetDSRRCollection(1);
            var deleteResult = collection.DeleteMany(Builders<BsonDocument>.Filter.Empty);

            var userCollection = DatabaseMongoCollection.GetMongoUserCollection();

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                string id = row.Cells["Số giấy tờ"].Value?.ToString() + ";";

                if (id == ";") continue;
                var document = new BsonDocument
                {
                    { "IdNum", id }
                };

                collection.InsertOne(document);


                string updatedNote = row.Cells["Ghi chú"].Value?.ToString();

                var filter = Builders<Person>.Filter.Eq("IdNum", id);
                var update = Builders<Person>.Update.Set("Note", updatedNote);

                var result = userCollection.UpdateOne(filter, update);
            }
            ShowCurrentList();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            UpdateList();
        }
        private void ShowCurrentList()
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            dataGridView2.Columns.Add("Số giấy tờ", "Số giấy tờ");
            dataGridView2.Columns.Add("Họ Tên", "Họ Tên");
            dataGridView2.Columns.Add("Số chuyến trong hệ thống", "Số chuyến trong hệ thống");
            dataGridView2.Columns.Add("Ghi chú", "Ghi chú");


            var collectionRR = DatabaseMongoCollection.GetDSRRCollection(1);
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
                dataGridViewRow.CreateCells(dataGridView2, id.Substring(0, id.Length - 1), name, countFlight, note);
                dataGridView2.Rows.Add(dataGridViewRow);
            }
        }

    }
}
