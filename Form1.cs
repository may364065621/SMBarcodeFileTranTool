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
        string version = "v1.0.0";
        string[] producer = { "思谋SmartMore", "X产品部" };
        public Form1()
        {
            InitializeComponent();
            cbb_Locate.SelectedIndex = 0;

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.Text = "思谋读码器输出表格转换工具  " + version + "  " + producer[0];

            minBarcodelist.Add("453");
            minBarcodelist.Add("493");
        }
        string selectedFile = "";
        string saveSelectedFile = "";
        string barcodeContrastSelectedFile = "";
        Dictionary<string, string[]> csvMesg = new Dictionary<string, string[]>();
        private void tb_InTable_DoubleClick(object sender, EventArgs e)
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
                tb_InTable.Text = selectedFile;
            }
        }

        private void tb_OutTable_DoubleClick(object sender, EventArgs e)
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
                tb_OutTable.Text = saveSelectedFile;
            }
        }
        private void tb_BarcodeContrast_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 设置对话框的属性
            openFileDialog.Title = "选择条码对照表文件";
            openFileDialog.Filter = "xlsx文件 (*.xlsx)|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // 显示对话框并检查用户是否选择了文件
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                barcodeContrastSelectedFile = openFileDialog.FileName;
                // 在这里处理选择的文件，例如将文件路径显示在标签(Label)控件上
                tb_BarcodeContrast.Text = barcodeContrastSelectedFile;
            }
        }

        private void btn_tran_Click(object sender, EventArgs e)
        {
            try
            {
                csvMesg.Clear();
                parse(selectedFile, saveSelectedFile);

                MessageBox.Show("转换成功");
            }
            catch (Exception ee)
            {
                MessageBox.Show("转换失败" + ee.ToString());
            }

        }
        /// <summary>
        /// A B F
        /// </summary>
        Dictionary<string, string[]> BarcodeContrast = new Dictionary<string, string[]>();

        private void btn_BarcodeContrast_Click(object sender, EventArgs e)
        {
            try
            {

                using (ExcelPackage package = new ExcelPackage(barcodeContrastSelectedFile))
                {
                    ExcelWorkbook workbook = package.Workbook;
                    bool tempstate = false;
                    foreach (ExcelWorksheet item in workbook.Worksheets)
                    {
                        if (item.Name == "对照表")
                        {
                            tempstate = true;
                            break;
                        }
                    }
                    if (!tempstate)
                    {
                        MessageBox.Show("该xlsx文件没有对照表的sheet");
                        return;
                    }
                    ExcelWorksheet worksheet = workbook.Worksheets["对照表"];

                    for (int i = 2; i < 2000000; i++)
                    {
                        if (worksheet.Cells[i, 4].Value != null &&
                            worksheet.Cells[i, 4].Value.ToString().Trim() != "")
                        {
                            string temp = worksheet.Cells[i, 4].Value.ToString().Trim();
                            if (BarcodeContrast.ContainsKey(temp))
                            {
                                BarcodeContrast[temp] = new string[] {  worksheet.Cells[i, 1].Value == null ? "" : worksheet.Cells[i, 1].Value.ToString().Trim(),
                                                                    worksheet.Cells[i, 2].Value == null ? "" : worksheet.Cells[i, 2].Value.ToString().Trim(),
                                                                    worksheet.Cells[i, 6].Value == null ? "" : worksheet.Cells[i, 6].Value.ToString().Trim()};
                            }
                            else
                            {
                                BarcodeContrast.Add(temp, new string[] {  worksheet.Cells[i, 1].Value == null ? "" : worksheet.Cells[i, 1].Value.ToString().Trim(),
                                                                    worksheet.Cells[i, 2].Value == null ? "" : worksheet.Cells[i, 2].Value.ToString().Trim(),
                                                                    worksheet.Cells[i, 6].Value == null ? "" : worksheet.Cells[i, 6].Value.ToString().Trim()});
                            }
                        }

                        else
                        {
                            break;
                        }
                    }
                }

                MessageBox.Show("读取成功");
            }
            catch (Exception ee)
            {
                MessageBox.Show("读取失败" + ee.ToString());
            }

        }
        List<string> minBarcodelist = new List<string>();

        private void parse(string csvFilePath, string xlsxFilePath)
        {
            string excelFilePath = xlsxFilePath;
            ;
            //string excelFilePath = csvFilePath.Substring(0, csvFilePath.LastIndexOf(".csv")) + ".xlsx";
            string currentSheetName = "";
            List<string> correctedSheetName = new List<string>();//有修改过的sheet表
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
                        if (values[okng_idx].Trim() == "" || values[okng_idx].Trim().Contains("NG"))
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
                //A-F写入xlsx
                foreach (string item in csvMesg.Keys)
                {
                    bool new_sheet = false;
                    string date = item.Substring(4, item.IndexOf('_') - 4);
                    currentSheetName = date;
                    if (!correctedSheetName.Contains(currentSheetName))
                    {
                        correctedSheetName.Add(currentSheetName);
                    }
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

                //清空H-M
                foreach (string item in correctedSheetName)
                {
                    workbook.Worksheets[item].DeleteColumn(8, 12);
                }
                //
                foreach (string item in correctedSheetName)
                {
                    ExcelWorksheet worksheet = workbook.Worksheets[item];

                    Dictionary<string, int> FbarcodeNum = new Dictionary<string, int>();
                    Dictionary<string, string[]> FMaxBarcodeNum = new Dictionary<string, string[]>();
                    Dictionary<string, string[]> FMinBarcodeNum = new Dictionary<string, string[]>();
                    Dictionary<string, string[]> FOtherBarcodeNum = new Dictionary<string, string[]>();
                    for (int i = 2; i < 2000000; i++)
                    {
                        if (worksheet.Cells[i, 6].Value != null && worksheet.Cells[i, 6].Value.ToString().Trim() != "")
                        {
                            string barcode = worksheet.Cells[i, 6].Value.ToString().Trim();
                            if (true)
                            {

                            }
                            if (FbarcodeNum.ContainsKey(barcode))
                            {
                                FbarcodeNum[barcode]++;
                            }
                            else
                            {
                                FbarcodeNum.Add(barcode, 1);
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    //排序 从小到大
                    FbarcodeNum = FbarcodeNum.OrderBy(item => item.Key).ToDictionary(item => item.Key, item => item.Value);
                    //细分到 大码、小码、其他码
                    foreach (string key in FbarcodeNum.Keys)
                    {
                        if (!BarcodeContrast.ContainsKey(key))
                        {
                            if (!FOtherBarcodeNum.ContainsKey(key))
                            {
                                FOtherBarcodeNum.Add(key, new string[] { "", "", "", FbarcodeNum[key].ToString() });
                            }
                        }
                        else
                        {
                            if (minBarcodelist.Contains(BarcodeContrast[key][0]))
                            {
                                FMinBarcodeNum.Add(key, new string[] { BarcodeContrast[key][0] + "-" + BarcodeContrast[key][1], BarcodeContrast[key][2], "6mm X 6mm", FbarcodeNum[key].ToString() });
                            }
                            else
                            {
                                FMaxBarcodeNum.Add(key, new string[] { BarcodeContrast[key][0] + "-" + BarcodeContrast[key][1], BarcodeContrast[key][2], "10mm X 10mm", FbarcodeNum[key].ToString() });

                            }
                        }

                    }
                    //写入到sheet中
                    //标题
                    int rows = 1; int num = 0; int maxnum = 0; int minnum = 0; int othernum = 0;
                    setCell(worksheet.Cells[rows, 8], new string[] { "不同", "Barcode" }, new string[] { "等线", "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                    setCell(worksheet.Cells[rows, 9], new string[] { "轮型&模号" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                    setCell(worksheet.Cells[rows, 10], new string[] { "客户" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                    setCell(worksheet.Cells[rows, 11], new string[] { "码大小" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                    setCell(worksheet.Cells[rows, 12], new string[] { "数量" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                    rows++;
                    //大码数据
                    if (FMaxBarcodeNum.Count > 0)
                    {
                        foreach (string key in FMaxBarcodeNum.Keys)
                        {
                            setCell(worksheet.Cells[rows, 8], new string[] { key }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                            for (int i = 0; i < 4; i++)
                            {
                                setCell(worksheet.Cells[rows, i + 9], new string[] { FMaxBarcodeNum[key][i] }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                            }
                            maxnum += Convert.ToInt32(FMaxBarcodeNum[key][3]);
                            rows++;
                        }
                        //大码小计
                        for (int i = 0; i < 3; i++)
                        {
                            for (int j = 0; j < 5; j++)
                            {
                                setCell(worksheet.Cells[rows + i, j + 8], new string[] { "" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                            }
                        }
                        setCell(worksheet.Cells[rows, 11], new string[] { "已读出数量：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                        setCell(worksheet.Cells[rows, 12], new string[] { maxnum.ToString() }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                        rows++;
                        setCell(worksheet.Cells[rows, 11], new string[] { "未读出数量：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                        setCell(worksheet.Cells[rows, 12], new string[] { "0" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                        rows++;
                        setCell(worksheet.Cells[rows, 11], new string[] { "小计：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                        setCell(worksheet.Cells[rows, 12], new string[] { maxnum.ToString() }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                        setCell(worksheet.Cells[rows, 13], new string[] { "100%" }, new string[] { "Arial" }, 11, false, ExcelHorizontalAlignment.CenterContinuous, Color.Cyan, true);
                        rows++;
                    }

                    //小码数据
                    if (FMinBarcodeNum.Count > 0)
                    {
                        foreach (string key in FMinBarcodeNum.Keys)
                        {
                            setCell(worksheet.Cells[rows, 8], new string[] { key }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                            for (int i = 0; i < 4; i++)
                            {
                                setCell(worksheet.Cells[rows, i + 9], new string[] { FMinBarcodeNum[key][i] }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                            }
                            minnum += Convert.ToInt32(FMinBarcodeNum[key][3]);
                            rows++;
                        }
                        //小码小计
                        for (int i = 0; i < 3; i++)
                        {
                            for (int j = 0; j < 5; j++)
                            {
                                setCell(worksheet.Cells[rows + i, j + 8], new string[] { "" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                            }
                        }
                        setCell(worksheet.Cells[rows, 11], new string[] { "已读出数量：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                        setCell(worksheet.Cells[rows, 12], new string[] { minnum.ToString() }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                        rows++;
                        setCell(worksheet.Cells[rows, 11], new string[] { "未读出数量：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                        setCell(worksheet.Cells[rows, 12], new string[] { "0" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                        rows++;
                        setCell(worksheet.Cells[rows, 11], new string[] { "小计：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                        setCell(worksheet.Cells[rows, 12], new string[] { minnum.ToString() }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                        setCell(worksheet.Cells[rows, 13], new string[] { "100%" }, new string[] { "Arial" }, 11, false, ExcelHorizontalAlignment.CenterContinuous, Color.Yellow, true);
                        rows++;
                    }

                    //其他小计
                    if (FOtherBarcodeNum.Count > 0)
                    {
                        //其他数据
                        foreach (string key in FOtherBarcodeNum.Keys)
                        {
                            setCell(worksheet.Cells[rows, 8], new string[] { key }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                            for (int i = 0; i < 4; i++)
                            {
                                setCell(worksheet.Cells[rows, i + 9], new string[] { FOtherBarcodeNum[key][i] }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                            }
                            othernum += Convert.ToInt32(FOtherBarcodeNum[key][3]);
                            rows++;
                        }
                        //其他小计
                        for (int i = 0; i < 3; i++)
                        {
                            for (int j = 0; j < 5; j++)
                            {
                                setCell(worksheet.Cells[rows + i, j + 8], new string[] { "" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                            }
                        }
                        setCell(worksheet.Cells[rows, 11], new string[] { "已读出数量：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                        setCell(worksheet.Cells[rows, 12], new string[] { othernum.ToString() }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                        rows++;
                        setCell(worksheet.Cells[rows, 11], new string[] { "未读出数量：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                        setCell(worksheet.Cells[rows, 12], new string[] { "0" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                        rows++;
                        setCell(worksheet.Cells[rows, 11], new string[] { "小计：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                        setCell(worksheet.Cells[rows, 12], new string[] { othernum.ToString() }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                        setCell(worksheet.Cells[rows, 13], new string[] { "100%" }, new string[] { "Arial" }, 11, false, ExcelHorizontalAlignment.CenterContinuous, Color.Coral, true);
                        rows++;
                    }
                    //汇总
                    for (int i = 0; i < 5; i++)
                    {
                        setCell(worksheet.Cells[rows, i + 8], new string[] { "" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Lime, true);
                    }
                    setCell(worksheet.Cells[rows, 11], new string[] { "合计：" }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Lime, true);
                    setCell(worksheet.Cells[rows, 12], new string[] { (maxnum + minnum + othernum).ToString() }, new string[] { "Arial" }, 11, true, ExcelHorizontalAlignment.CenterContinuous, Color.Lime, true);

                    int startRow = 2;
                    int endRow = 8;
                    int startCol = 1;
                    int endCol = 4;

                    // 设置边框样式为实线
                    var border = worksheet.Cells[1, 8, rows, 12].Style.Border;

                    border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                    ;
                    border.BorderAround(ExcelBorderStyle.Medium);

                    worksheet.Column(7).Width = 0.46;
                    worksheet.Column(8).Width = 15;
                    worksheet.Column(9).Width = 15;
                    worksheet.Column(10).Width = 15;
                    worksheet.Column(11).Width = 15;
                    worksheet.Column(12).Width = 8.38;

                }




                //sheet表以名字排序
                List<string> sheetNames = new List<string>();
                foreach (var sheet in workbook.Worksheets)
                {
                    sheetNames.Add(sheet.Name);
                }
                sheetNames.Sort();
                foreach (var sheetName in sheetNames)
                {
                    workbook.Worksheets.MoveToStart(sheetName);
                }


                // 保存Excel文件
                FileInfo excelFile = new FileInfo(excelFilePath);
                package.SaveAs(excelFile);
                label_savepath.Text += "转换完成";
            }
        }

        public static void setCell(ExcelRange cell, string[] values, string[] fonts, int size, bool bold, ExcelHorizontalAlignment Horizon, Color color, bool setcolor = false)
        {
            if (values.Length > 1)
            {
                setCellFont(cell, size, bold, Horizon, fonts[0], color, setcolor);
                cell.IsRichText = true;
                for (int i = 0; i < values.Length; i++)
                {
                    var part = cell.RichText.Add(values[i]);
                    part.FontName = fonts[i];
                }
            }
            else
            {
                cell.Value = values[0];
                setCellFont(cell, size, bold, Horizon, fonts[0], color, setcolor);

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

        private void Form1_Load(object sender, EventArgs e)
        {

        }


    }
}