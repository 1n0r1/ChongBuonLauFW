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
            var collection = MongoUserCollection.GetDSRRCollection(1);
            var deleteResult = collection.DeleteMany(Builders<BsonDocument>.Filter.Empty);

            foreach (DataGridViewRow row in dataGridView2.Rows)
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

            var collectionRR = MongoUserCollection.GetDSRRCollection(1);
            var collection = MongoUserCollection.GetMongoUserCollection();
            var documents = collectionRR.Find(Builders<BsonDocument>.Filter.Empty).ToList();
            foreach (var document in documents)
            {
                string id = document["IdNum"].AsString;

                string name = "-";
                int countFlight = 0;

                var filter = Builders<Person>.Filter.Eq("IdNum", id);
                var person = collection.Find(filter).FirstOrDefault();

                if (person != null)
                {
                    name = person.Name;
                    countFlight = person.FlightList.Count;
                }

                DataGridViewRow dataGridViewRow = new DataGridViewRow();
                dataGridViewRow.CreateCells(dataGridView2, id.Substring(0, id.Length - 1), name, countFlight);
                dataGridView2.Rows.Add(dataGridViewRow);
            }
        }

    }
}
