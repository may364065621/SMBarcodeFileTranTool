using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data.Common;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            cbb_Locate.SelectedIndex = 0;

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.Text = "格式转换  " + version + "  " + producer;
        }
        string version = "v1.0.0";
        string producer = "思谋SmartMore";
        string selectedFile = "";
        private void btn_readcsv_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 设置对话框的属性
            openFileDialog.Title = "选择文件";
            openFileDialog.Filter = "CSV文件 (*.csv)|*.csv";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // 显示对话框并检查用户是否选择了文件
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFile = openFileDialog.FileName;
                // 在这里处理选择的文件，例如将文件路径显示在标签(Label)控件上
                label_Date.Text = selectedFile;
            }
        }
        Dictionary<string, string[]> csvMesg = new Dictionary<string, string[]>();

        string saveSelectedFile = "";
        private void btn_tran_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 设置对话框的属性
            openFileDialog.Title = "请选择保存到哪个文件";
            openFileDialog.Filter = "xlsx文件 (*.xlsx)|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // 显示对话框并检查用户是否选择了文件
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                saveSelectedFile = openFileDialog.FileName;
                // 在这里处理选择的文件，例如将文件路径显示在标签(Label)控件上
                label_savepath.Text = saveSelectedFile;
            }
            csvMesg.Clear();
            parse(selectedFile, saveSelectedFile);
        }
        private void parse(string csvFilePath, string xlsxFilePath)
        {
            string excelFilePath = xlsxFilePath;
            ;
            //string excelFilePath = csvFilePath.Substring(0, csvFilePath.LastIndexOf(".csv")) + ".xlsx";
            string currentSheetName = "";

            Dictionary<string, List<string>> worksheetnames = new Dictionary<string, List<string>>();


            using (ExcelPackage package = new ExcelPackage(xlsxFilePath))
            {
                Dictionary<string, string> date_hashMap = new Dictionary<string, string>();

                // 打开CSV文件并逐行读取内容
                using (StreamReader reader = new StreamReader(csvFilePath))
                {

                    string line = "";
                    while ((line = reader.ReadLine()) != null)
                    {
                        string[] values = line.Split(',');
                        int date_idx = 0;
                        int okng_idx = 1;
                        int barcode_data_idx = 4;
                        if (values[okng_idx].Trim() == "" || values[okng_idx].Trim() == "NG")
                        {
                            continue;
                        }

                        if (csvMesg.ContainsKey(values[date_idx].Substring(2, 21)))
                        {

                            csvMesg[values[date_idx].Substring(2, 21)] = new string[] { values[barcode_data_idx].Split(new char[] { '"' })[1] };
                        }
                        else
                        {
                            csvMesg.Add(values[date_idx].Substring(2, 21), new string[] { values[barcode_data_idx].Split(new char[] { '"' })[1] });
                        }

                    }
                }

                ExcelWorkbook workbook = package.Workbook;
                foreach (ExcelWorksheet item in workbook.Worksheets)
                {
                    if (!worksheetnames.ContainsKey(item.Name))
                    {
                        worksheetnames.Add(item.Name, new List<string>() { "1" });
                        for (int i = 1; i < 2000000; i++)
                        {
                            if (item.Cells[i, 2].Value != null)
                            {
                                if (item.Cells[i, 2].Value.ToString() == "")
                                {
                                    worksheetnames[item.Name][0] = i.ToString();
                                    break;
                                }
                                else
                                {
                                    worksheetnames[item.Name].Add(item.Cells[i, 2].Value.ToString());
                                }
                            }
                            else
                            {
                                worksheetnames[item.Name][0] = i.ToString();
                                break;
                            }
                        }
                    }
                }
                //int row = 1;
                //写入xlsx
                foreach (string item in csvMesg.Keys)
                {
                    bool new_sheet = false;
                    string date = item.Substring(4, item.IndexOf('_') - 4);
                    currentSheetName = date;
                    ExcelWorksheet worksheet;
                    if (!worksheetnames.ContainsKey(date))
                    {
                        new_sheet = true;
                        if (worksheetnames.ContainsKey("Sheet1"))
                        {
                            worksheetnames.Remove("Sheet1");
                            workbook.Worksheets["Sheet1"].Name = currentSheetName;
                        }
                        else if (worksheetnames.ContainsKey("Sheet2"))
                        {
                            worksheetnames.Remove("Sheet2");
                            workbook.Worksheets["Sheet2"].Name = currentSheetName;
                        }
                        else if (worksheetnames.ContainsKey("Sheet3"))
                        {
                            worksheetnames.Remove("Sheet3");
                            workbook.Worksheets["Sheet3"].Name = currentSheetName;
                        }
                        else
                        {
                            workbook.Worksheets.Add(currentSheetName);
                        }
                        worksheetnames.Add(currentSheetName, new List<string>() { "1" });
                    }

                    worksheet = workbook.Worksheets[currentSheetName];
                    int rows = Convert.ToInt32(worksheetnames[currentSheetName][0]);

                    if (new_sheet)
                    {
                        using (ExcelRange rng = worksheet.Cells["A1:A5"])
                        {
                            rng.Merge = true;
                        }
                        setColumnFont(worksheet.Column(1), 11, true, ExcelHorizontalAlignment.CenterContinuous, "Arial");
                        setColumnFont(worksheet.Column(2), 11, false, ExcelHorizontalAlignment.CenterContinuous, "Arial");
                        setColumnFont(worksheet.Column(3), 11, false, ExcelHorizontalAlignment.CenterContinuous, "Arial");
                        setColumnFont(worksheet.Column(4), 11, false, ExcelHorizontalAlignment.CenterContinuous, "Arial");
                        setColumnFont(worksheet.Column(5), 11, true, ExcelHorizontalAlignment.CenterContinuous, "Arial");
                        setColumnFont(worksheet.Column(6), 11, true, ExcelHorizontalAlignment.CenterContinuous, "Arial");



                        setCellFont(worksheet.Cells[rows, 1], 11, true, ExcelHorizontalAlignment.CenterContinuous, "宋体", Color.Lime, false);
                        worksheet.Cells[rows, 1].IsRichText = true;
                        var part1 = worksheet.Cells[rows, 1].RichText.Add("读码成功数据                       ");
                        part1.FontName = "宋体";
                        var part2 = worksheet.Cells[rows, 1].RichText.Add(" ("
                            + item.Substring(4, 2) + "/" + item.Substring(6, 2) + " 00:00-"
                            + item.Substring(4, 2) + "/" + item.Substring(6, 2) + " 24:00) ");
                        part2.Bold = false;
                        part2.FontName = "Arial";

                        worksheet.Cells[rows, 2].Value = "时间";
                        setCellFont(worksheet.Cells[rows, 2], 11, true, ExcelHorizontalAlignment.CenterContinuous, "Arial", Color.Lime, true);
                        worksheet.Cells[rows, 3].Value = "记录";
                        setCellFont(worksheet.Cells[rows, 3], 11, true, ExcelHorizontalAlignment.CenterContinuous, "Arial", Color.Lime, true);
                        worksheet.Cells[rows, 4].Value = "位置";
                        setCellFont(worksheet.Cells[rows, 4], 11, true, ExcelHorizontalAlignment.CenterContinuous, "Arial", Color.Lime, true);
                        worksheet.Cells[rows, 5].Value = "取QR Code前9位";
                        setCellFont(worksheet.Cells[rows, 5], 11, true, ExcelHorizontalAlignment.CenterContinuous, "Arial", Color.Lime, true);
                        worksheet.Cells[rows, 6].Value = "取Barode";
                        setCellFont(worksheet.Cells[rows, 6], 11, true, ExcelHorizontalAlignment.CenterContinuous, "Arial", Color.Lime, true);

                        worksheet.Column(1).Width = 20;
                        worksheet.Column(2).Width = 24;
                        worksheet.Column(3).Width = 20;
                        worksheet.Column(4).Width = 6;
                        worksheet.Column(5).Width = 30;
                        worksheet.Column(6).Width = 12;
                        worksheet.Column(7).Width = 12;
                        worksheet.Column(8).Width = 0.46;
                        rows++; worksheetnames[currentSheetName][0] = rows.ToString();
                    }
                    // 写入每个字段到Excel工作表

                    if (worksheetnames[currentSheetName].Contains(item))
                    {
                        rows = worksheetnames[currentSheetName].IndexOf(item);
                        worksheet.Cells[rows, 2].Value = item;
                        worksheet.Cells[rows, 3].Value = csvMesg[item][0];
                        worksheet.Cells[rows, 4].Value = cbb_Locate.Text.Trim();
                        worksheet.Cells[rows, 5].Value = csvMesg[item][0].Substring(0, 9);
                        worksheet.Cells[rows, 6].Value = csvMesg[item][0].Substring(4, 5);
                    }
                    else
                    {
                        worksheetnames[currentSheetName].Add(item);
                        worksheet.Cells[rows, 2].Value = item;
                        worksheet.Cells[rows, 3].Value = csvMesg[item][0];
                        worksheet.Cells[rows, 4].Value = cbb_Locate.Text.Trim();
                        worksheet.Cells[rows, 5].Value = csvMesg[item][0].Substring(0, 9);
                        worksheet.Cells[rows, 6].Value = csvMesg[item][0].Substring(4, 5);
                        rows++;
                        worksheetnames[currentSheetName][0] = rows.ToString();
                    }


                }
                // 保存Excel文件
                FileInfo excelFile = new FileInfo(excelFilePath);
                package.SaveAs(excelFile);

                label_savepath.Text += "转换完成";
            }


        }

        public static void setCellFont(ExcelRange cell, int size, bool bold, ExcelHorizontalAlignment Horizon, string fontName, Color color, bool setcolor = false)
        {
            cell.Style.Font.Size = size;
            cell.Style.Font.Bold = bold;
            cell.Style.HorizontalAlignment = Horizon;
            cell.Style.Font.Name = fontName;
            if (setcolor)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(color);
            }
            cell.Style.WrapText = true;
        }
        public static void setColumnFont(ExcelColumn column, int size, bool bold, ExcelHorizontalAlignment Horizon, string fontName)
        {
            column.Style.Font.Size = size;
            column.Style.Font.Bold = bold;
            column.Style.HorizontalAlignment = Horizon;
            column.Style.Font.Name = fontName;
            column.Style.WrapText = true;
        }
    }
}