using MongoDB.Bson;
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
    public partial class Form3 : Form
    {
        public Form3(BsonArray dataArray, string flightdate, string flightNumber, string luggage, string sta, string des, string luggageCount)
        {
            InitializeComponent();

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
            dataGridView1.Columns.Add("Số hành lý", "Hành lý");

            foreach (var res in dataArray)
            {
                string seat = res["Seat"].ToString();
                string name = res["Name"].ToString();
                string sex = res["Sex"].ToString();
                string nat = res["Nationality"].ToString();
                string dob = res["DOB"].ToString();
                string idtype = res["IdType"].ToString();
                string idnum = res["IdNum"].ToString();
                string idprov = res["IdProv"].ToString();
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView1, flightdate, flightNumber, seat, name, sex, nat, dob, idtype, idnum, idprov, sta, des, luggage, luggageCount);
                dataGridView1.Rows.Add(row);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportToExcel(dataGridView1);
        }
    }
}
