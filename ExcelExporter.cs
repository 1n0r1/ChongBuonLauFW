using ClosedXML.Excel;
using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace ChongBuonLauFW
{
    public static class ExcelExporter
    {
        public static void ExportToExcel(DataGridView dataGridView)
        {
            // Create a SaveFileDialog to specify the file path
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                saveFileDialog.Title = "Export to Excel";
                saveFileDialog.FileName = "example.xlsx"; // Default file name

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new XLWorkbook();

                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    for (int i = 0; i < dataGridView.Columns.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = dataGridView.Columns[i].HeaderText;
                    }
                    for (int i = 0; i < dataGridView.Rows.Count; i++)
                    {
                        bool isYellow = false;
                        var colorCell = dataGridView.Columns.Contains("Color") ? dataGridView.Rows[i].Cells["Color"].Value : null;
                        if (colorCell != null && colorCell.ToString() == "yellow") isYellow = true;

                        for (int j = 0; j < dataGridView.Columns.Count; j++)
                        {
                            object cellValue = dataGridView.Rows[i].Cells[j].Value;
                            worksheet.Cell(i + 2, j + 1).Value = cellValue?.ToString() ?? "";
                            if (isYellow)
                            {
                                worksheet.Cell(i + 2, j + 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                            }
                        }
                    }

                    workbook.SaveAs(saveFileDialog.FileName);
                }
            }
        }
        public static void ExportForm(DataGridView dataGridView)
        {
            // Create a SaveFileDialog to specify the file path
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                saveFileDialog.Title = "Export to Excel";
                saveFileDialog.FileName = "example.xlsx"; // Default file name

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new XLWorkbook();

                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    worksheet.Cell(1, 1).Value = "PHIẾU THU THẬP THÔNG TIN QLRR";
                    worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 4)).Merge();
                    IXLRange title = worksheet.Range("A1:D1");
                    title.Style.Font.Bold = true;

                    worksheet.Cell(3, 1).Value = "Tuyến trọng điểm:";
                    worksheet.Range(worksheet.Cell(3, 1), worksheet.Cell(3, 2)).Merge();
                    worksheet.Cell(4, 1).Value = "Chuyến trọng điểm:";
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 2)).Merge();
                    worksheet.Cell(5, 1).Value = "Hành khách trong DS trọng điểm:";
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 3)).Merge();
                    worksheet.Cell(6, 1).Value = "Hành khách rà soát trọng điểm:";
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 3)).Merge();

                    worksheet.Cell(8, 1).Value = "STT";
                    worksheet.Cell(8, 2).Value = "Số chuyến bay";
                    worksheet.Cell(8, 3).Value = "Ngày đến";
                    worksheet.Cell(8, 4).Value = "Sân bay khởi hành";
                    worksheet.Cell(8, 5).Value = "Sân bay đến";
                    worksheet.Cell(8, 6).Value = "Họ tên";
                    worksheet.Cell(8, 7).Value = "Giới tính";
                    worksheet.Cell(8, 8).Value = "Ngày sinh";
                    worksheet.Cell(8, 9).Value = "Quốc tịch";
                    worksheet.Cell(8, 10).Value = "Số giấy tờ";
                    worksheet.Cell(8, 11).Value = "Số thẻ hành lý";
                    IXLRange headerRange = worksheet.Range("A8:K8");
                    headerRange.Style.Font.Bold = true;

                    dataGridView.Sort(dataGridView.Columns["Chuyến bay"], ListSortDirection.Ascending);

                    var index = 1;
                    string previousChuyenBay = "";
                    int startRow = 8;
                    int endRow = 8;

                    foreach (DataGridViewRow row in dataGridView.Rows)
                    {
                        if (row.Cells["Chuyến bay"].Value == null) continue;

                        // if ((row.Cells["Số khách"].Value?.ToString() ?? "") != "1") continue;

                        string currentChuyenBay = row.Cells["Chuyến bay"].Value?.ToString() ?? "";

                        Console.WriteLine($"ChuyenBay {previousChuyenBay} {currentChuyenBay}");
                        if (currentChuyenBay != previousChuyenBay)
                        {
                            Console.WriteLine($"Row {startRow}-{endRow}");
                            if (startRow != endRow)
                            {
                                IXLRange mergedRange = worksheet.Range($"B{startRow}:B{endRow}");
                                IXLRange mmergedRange = worksheet.Range($"C{startRow}:C{endRow}");
                                IXLRange mmmergedRange = worksheet.Range($"E{startRow}:E{endRow}");
                                mergedRange.Merge();
                                mmergedRange.Merge();
                                mmmergedRange.Merge();
                            }

                            previousChuyenBay = currentChuyenBay;
                            startRow = index + 8;
                        }

                        endRow = index + 8;

                        worksheet.Cell(index + 8, 1).Value = index;
                        worksheet.Cell(index + 8, 2).Value = row.Cells["Chuyến bay"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 3).Value = row.Cells["Ngày bay"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 4).Value = row.Cells["Nơi đi"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 5).Value = row.Cells["Nơi đến"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 6).Value = row.Cells["Họ tên"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 7).Value = row.Cells["Giới tính"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 8).Value = row.Cells["Ngày sinh"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 9).Value = row.Cells["Quốc tịch"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 10).Value = row.Cells["Số giấy tờ"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 11).Value = row.Cells["Hành lý"].Value?.ToString() ?? "";

                        index++;
                        Console.WriteLine($"{index}");
                    }

                    DateTime currentDate = DateTime.UtcNow;
                    TimeZoneInfo vietnamZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
                    DateTime vietnamTime = TimeZoneInfo.ConvertTimeFromUtc(currentDate, vietnamZone);
                    string formattedDate = vietnamTime.ToString("'Ngày' dd 'tháng' MM 'năm' yyyy");
                    worksheet.Cell(index + 8, 11).Value = formattedDate;
                    worksheet.Cell(index + 9, 11).Value = "Người lập phiếu";


                    IXLRange range = worksheet.Range($"A8:K8");
                    range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    IXLRange rrange = worksheet.Range($"A8:E{index + 8}");
                    rrange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rrange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    IXLRange rrrange = worksheet.Range($"K{index + 8}:K{index + 9}");
                    rrrange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rrrange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    IXLRange rrrrange = worksheet.Range($"A8:B{index + 8}");
                    rrrrange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rrrrange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    rrrrange.Style.Font.Bold = true;

                    var columns = worksheet.Columns("B:K");
                    columns.AdjustToContents();


                    IXLRange tableRange = worksheet.Range($"A8:K{index + 7}");
                    tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    tableRange.Style.Border.OutsideBorderColor = XLColor.Black;
                    tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    tableRange.Style.Border.InsideBorderColor = XLColor.Black;

                    workbook.SaveAs(saveFileDialog.FileName);
                    dataGridView.Sort(dataGridView.Columns["Số hành lý"], ListSortDirection.Descending);
                }
            }
        }
        public static void ExportFormHK(DataGridView dataGridView, bool merge)
        {
            // Create a SaveFileDialog to specify the file path
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                saveFileDialog.Title = "Export to Excel";
                saveFileDialog.FileName = "example.xlsx"; // Default file name

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new XLWorkbook();

                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    worksheet.Cell(1, 1).Value = "PHIẾU THU THẬP THÔNG TIN QLRR";
                    worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 4)).Merge();
                    IXLRange title = worksheet.Range("A1:D1");
                    title.Style.Font.Bold = true;

                    worksheet.Cell(3, 1).Value = "Tuyến trọng điểm:";
                    worksheet.Range(worksheet.Cell(3, 1), worksheet.Cell(3, 2)).Merge();
                    worksheet.Cell(4, 1).Value = "Chuyến trọng điểm:";
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 2)).Merge();
                    worksheet.Cell(5, 1).Value = "Hành khách trong DS trọng điểm:";
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 3)).Merge();
                    worksheet.Cell(6, 1).Value = "Hành khách rà soát trọng điểm:";
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 3)).Merge();

                    worksheet.Cell(8, 1).Value = "STT";
                    worksheet.Cell(8, 2).Value = "Số chuyến bay";
                    worksheet.Cell(8, 3).Value = "Ngày đến";
                    worksheet.Cell(8, 4).Value = "Sân bay khởi hành";
                    worksheet.Cell(8, 5).Value = "Sân bay đến";
                    worksheet.Cell(8, 6).Value = "Họ tên";
                    worksheet.Cell(8, 7).Value = "Giới tính";
                    worksheet.Cell(8, 8).Value = "Ngày sinh";
                    worksheet.Cell(8, 9).Value = "Quốc tịch";
                    worksheet.Cell(8, 10).Value = "Số giấy tờ";
                    worksheet.Cell(8, 11).Value = "Số thẻ hành lý";
                    IXLRange headerRange = worksheet.Range("A8:K8");
                    headerRange.Style.Font.Bold = true;

                    dataGridView.Sort(dataGridView.Columns["Chuyến bay"], ListSortDirection.Ascending);

                    var index = 1;
                    string previousChuyenBay = "";
                    int startRow = 8;
                    int endRow = 8;

                    foreach (DataGridViewRow row in dataGridView.Rows)
                    {
                        bool isYellow = false;
                        var colorCell = dataGridView.Columns.Contains("Color") ? row.Cells["Color"].Value : null;
                        if (colorCell != null && colorCell.ToString() == "yellow") isYellow = true;

                        if (row.Cells["Chuyến bay"].Value == null) continue;

                        string currentChuyenBay = row.Cells["Chuyến bay"].Value?.ToString() ?? "";

                        Console.WriteLine($"ChuyenBay {previousChuyenBay} {currentChuyenBay}");
                        if (merge && currentChuyenBay != previousChuyenBay)
                        {
                            Console.WriteLine($"Row {startRow}-{endRow}");
                            if (startRow != endRow)
                            {
                                IXLRange mergedRange = worksheet.Range($"B{startRow}:B{endRow}");
                                IXLRange mmergedRange = worksheet.Range($"C{startRow}:C{endRow}");
                                IXLRange mmmergedRange = worksheet.Range($"E{startRow}:E{endRow}");
                                mergedRange.Merge();
                                mmergedRange.Merge();
                                mmmergedRange.Merge();
                            }

                            previousChuyenBay = currentChuyenBay;
                            startRow = index + 8;
                        }

                        endRow = index + 8;

                        worksheet.Cell(index + 8, 1).Value = index;
                        worksheet.Cell(index + 8, 2).Value = row.Cells["Chuyến bay"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 3).Value = row.Cells["Ngày bay"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 4).Value = row.Cells["Nơi đi"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 5).Value = row.Cells["Nơi đến"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 6).Value = row.Cells["Họ tên"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 7).Value = row.Cells["Giới tính"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 8).Value = row.Cells["Ngày sinh"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 9).Value = row.Cells["Quốc tịch"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 10).Value = row.Cells["Số giấy tờ"].Value?.ToString() ?? "";
                        worksheet.Cell(index + 8, 11).Value = row.Cells["Hành lý"].Value?.ToString() ?? "";

                        if (isYellow)
                        {
                            for (int i = 1; i <= 11; i++)
                            {
                                worksheet.Cell(index + 8, i).Style.Fill.BackgroundColor = XLColor.Yellow;
                            }
                        }

                        index++;
                    }

                    DateTime currentDate = DateTime.UtcNow;
                    TimeZoneInfo vietnamZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
                    DateTime vietnamTime = TimeZoneInfo.ConvertTimeFromUtc(currentDate, vietnamZone);
                    string formattedDate = vietnamTime.ToString("'Ngày' dd 'tháng' MM 'năm' yyyy");
                    worksheet.Cell(index + 8, 11).Value = formattedDate;
                    worksheet.Cell(index + 9, 11).Value = "Người lập phiếu";


                    IXLRange range = worksheet.Range($"A8:K8");
                    range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    IXLRange rrange = worksheet.Range($"A8:E{index + 8}");
                    rrange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rrange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    IXLRange rrrange = worksheet.Range($"K{index + 8}:K{index + 9}");
                    rrrange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rrrange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    IXLRange rrrrange = worksheet.Range($"A8:B{index + 8}");
                    rrrrange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rrrrange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    rrrrange.Style.Font.Bold = true;

                    var columns = worksheet.Columns("B:K");
                    columns.AdjustToContents();


                    IXLRange tableRange = worksheet.Range($"A8:K{index + 7}");
                    tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    tableRange.Style.Border.OutsideBorderColor = XLColor.Black;
                    tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    tableRange.Style.Border.InsideBorderColor = XLColor.Black;

                    workbook.SaveAs(saveFileDialog.FileName);
                    try
                    {
                        dataGridView.Sort(dataGridView.Columns["Số lần nhiều nhất"], ListSortDirection.Descending);
                    }
                    catch
                    {
                    }
                }
            }
        }
    
    }
}
