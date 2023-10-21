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
    public partial class Form2 : Form
    {
        public Form2(string id)
        {
            InitializeComponent();

            var collection = MongoUserCollection.GetMongoUserCollection();
            var filter = Builders<Person>.Filter.Eq("IdNum", id);
            var result = collection.Find(filter).ToList();

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

            if (result.Count > 0)
            {
                foreach (var res in result)
                {
                    List<Flight> flightList = res.FlightList;
                    foreach (var flight in flightList)
                    {
                        TimeZoneInfo vstZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
                        DateTime flightdate = TimeZoneInfo.ConvertTimeFromUtc(flight.Date, vstZone);
                        string flightNumber = flight.FlightNumber;
                        string seat = flight.Seat;
                        string name = res.Name;
                        string sex = res.Sex;
                        string nat = res.Nationality;
                        string dob = res.DOB;
                        string idtype = res.IdType;
                        string idnum = res.IdNum;
                        string idprov = res.IdProv;
                        string sta = flight.Origin;
                        string des = flight.Destination;
                        string luggage = flight.Luggage;

                        DataGridViewRow row = new DataGridViewRow();
                        row.CreateCells(dataGridView1, flightdate, flightNumber, seat, name, sex, nat, dob, idtype, idnum, idprov, sta, des, luggage);
                        dataGridView1.Rows.Add(row);
                    }
                }
                dataGridView1.Sort(dataGridView1.Columns["Ngày bay"], ListSortDirection.Ascending);
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportToExcel(dataGridView1);
        }
    }
}
