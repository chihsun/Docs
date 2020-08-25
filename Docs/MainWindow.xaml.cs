using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace Docs
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    #region Class
    public class DocType
    {
        public string Index { get; set; }
        public string ID { get; set; }
        public string webID { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }
        public string Depart { get; set; }
        public string doctp { get; set; }
        public DateTime Stime { get; set; }
        public DateTime Rtime { get; set; }
        public DateTime Ntime
        {
            get
            {
                return this.Rtime.AddYears(1);
            }
        }
        public DateTime Etime
        {
            get
            {
                return this.Rtime.AddYears(3);
            }
        }
        public string Own { get; set; }
        public bool Eng { get; set; }
        public string Color { get; set; }

    }
    #endregion
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;
        }
        #region Parameter
        public List<DocType> ADocs = new List<DocType>();
        #endregion
        #region Method
        public void LoadData(string fname)
        {
            if (!System.IO.File.Exists(fname))
                return;
            try
            {
                SLDocument sl = new SLDocument(fname, "20200731更新");
                SLWorksheetStatistics stats = sl.GetWorksheetStatistics();
                if (stats.EndRowIndex <= 0)
                    return;
                for (int i = 0; i < stats.EndRowIndex; i++)
                {
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                        break;
                    DocType docs = new DocType
                    {
                        Index = sl.GetCellValueAsString(i + 2, 1),
                        ID = sl.GetCellValueAsString(i + 2, 2).Trim(),
                        Color = sl.GetCellStyle(i + 2, 2).Font.FontColor.ToString(),
                        webID = sl.GetCellValueAsString(i + 2, 13).Trim(),
                        Name = sl.GetCellValueAsString(i + 2, 4).Trim(),
                        Version = sl.GetCellValueAsString(i + 2, 5).Trim(),
                        Depart = sl.GetCellValueAsString(i + 2, 6).Trim(),
                        doctp = sl.GetCellValueAsString(i + 2, 7).Trim(),
                        Stime = sl.GetCellValueAsDateTime(i + 2, 8),
                        Rtime = sl.GetCellValueAsDateTime(i + 2, 9),
                        Own = sl.GetCellValueAsString(i + 2, 11).Trim()
                    };
                    if (sl.GetCellStyle(i + 2, 2).Font.FontColor == System.Drawing.Color.FromArgb(255, 255, 0, 0))
                        docs.Eng = true;
                    if (docs.webID.Contains("documentId"))
                    {
                        var m = Regex.Match(docs.webID, @"documentId=(\d+)");
                        if (m.Success)
                        {
                            docs.webID = m.Groups[1].ToString();
                        }
                        else
                            docs.webID = "-1";
                    }
                    else if (Int32.TryParse(sl.GetCellValueAsString(i + 2, 3).Trim(), out int id))
                    {
                        docs.webID = id.ToString();
                    }
                    else
                        docs.webID = "-1";

                   ADocs.Add(docs);
                }
                sl.CloseWithoutSaving();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
                /*
                odocs = ADocs.GroupBy(o => o.Depart).ToDictionary(o => o.Key, o => o.ToList());
                try
                {
                    foreach (var x in odocs)
                    {
                        int i = 0;
                        SLDocument sl = new SLDocument();
                        foreach (var y in x.Value)
                        {
                            sl.SetCellValue(i + 1, 1, y.ID);
                            sl.SetCellValue(i + 1, 2, y.webID);
                            sl.SetCellValue(i + 1, 3, y.Name);
                            sl.SetCellValue(i + 1, 4, y.Version);
                            sl.SetCellValue(i + 1, 5, y.Depart);
                            sl.SetCellValue(i + 1, 6, y.Own);
                            i++;
                        }
                        string fpath = Environment.CurrentDirectory + @"\部門";
                        if (!System.IO.Directory.Exists(fpath))
                        {
                            System.IO.Directory.CreateDirectory(fpath);
                        }
                        sl.SaveAs(fpath + @"\" + x.Key + ".xlsx");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                */
        }
        public void LoadFile()
        {
            string fpath = Environment.CurrentDirectory;
            if (!System.IO.Directory.Exists(fpath))
            {
                System.IO.Directory.CreateDirectory(fpath);
            }
            string fname = fpath + @"\合併檔\合併總表.xlsx";
            if (!System.IO.File.Exists(fname))
                return;
            LoadData(fname);
            ExportOwn();
        }
        public void ExportOwn()
        {
            if (ADocs.Count > 0)
            {
                var odocs = ADocs.GroupBy(o => o.Own).ToDictionary(o => o.Key, o => o.ToList());
                try
                {
                    foreach (var x in odocs)
                    {
                        SLDocument sl = ExportExcel(x.Value);
                        string fpath = Environment.CurrentDirectory + @"\管理員";
                        if (!System.IO.Directory.Exists(fpath))
                        {
                            System.IO.Directory.CreateDirectory(fpath);
                        }
                        sl.SaveAs(fpath + @"\" + x.Key + ".xlsx");
                        this.TxtBox1.Text += Environment.NewLine
                            + x.Key + string.Format(" 負責: {0, 5} 份", x.Value.Count.ToString());
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        public SLDocument ExportExcel(List<DocType> dts)
        {
            int i = 1;
            SLDocument sl = new SLDocument();
            SLConditionalFormatting cf = new SLConditionalFormatting("C2", "C1000");
            cf.HighlightCellsWithDuplicates(SLHighlightCellsStyleValues.LightRedFill);
            cf.HighlightCellsEndingWith("0", SLHighlightCellsStyleValues.LightRedFill);
            sl.AddConditionalFormatting(cf);

            SLStyle style = sl.CreateStyle();
            style.Alignment.WrapText = true;
            style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            style.Alignment.Vertical = VerticalAlignmentValues.Center;
            sl.SetCellStyle(1, 1, style);
            sl.SetCellStyle(1, 2, style);
            sl.SetCellStyle(1, 3, style);
            sl.SetCellStyle(1, 4, style);
            sl.SetCellStyle(1, 5, style);
            sl.SetCellStyle(1, 6, style);
            sl.SetCellStyle(1, 7, style);
            sl.SetCellStyle(1, 8, style);
            sl.SetCellStyle(1, 9, style);
            sl.SetCellStyle(1, 10, style);
            sl.SetCellStyle(1, 11, style);
            sl.SetCellStyle(1, 12, style);
            sl.SetCellStyle(1, 13, style);
            sl.SetColumnWidth(1, 10);
            sl.SetColumnWidth(2, 15);
            sl.SetColumnWidth(3, 10);
            sl.SetColumnWidth(4, 60);
            sl.SetColumnWidth(5, 10);
            sl.SetColumnWidth(6, 20);
            sl.SetColumnWidth(7, 15);
            sl.SetColumnWidth(8, 20);
            sl.SetColumnWidth(9, 20);
            sl.SetColumnWidth(10, 20);
            sl.SetColumnWidth(11, 10);
            sl.SetColumnWidth(12, 10);
            sl.SetColumnWidth(13, 50);
            sl.SetCellValue(1, 1, "表單序號");
            sl.SetCellValue(1, 2, "表單代號");
            sl.SetCellValue(1, 3, "網頁代碼");
            sl.SetCellValue(1, 4, "表單名稱");
            sl.SetCellValue(1, 5, "版本");
            sl.SetCellValue(1, 6, "制訂單位");
            sl.SetCellValue(1, 7, "文件類別");
            sl.SetCellValue(1, 8, "首次公佈時間");
            sl.SetCellValue(1, 9, "上次檢視時間");
            sl.SetCellValue(1, 10, "預計檢視時間");
            sl.SetCellValue(1, 11, "負責同仁");
            sl.SetCellValue(1, 12, "備註");
            sl.SetCellValue(1, 13, "備註(2)");
            style.Font.FontColor = System.Drawing.Color.Red;
            foreach (var y in dts)
            {
                sl.SetCellValue(i + 1, 1, y.Index);
                sl.SetCellValue(i + 1, 2, y.ID);
                if (y.Eng)
                    sl.SetCellStyle(i + 1, 2, style);
                sl.SetCellValue(i + 1, 3, y.webID);
                sl.SetCellValue(i + 1, 4, y.Name);
                sl.SetCellValue(i + 1, 5, y.Version);
                sl.SetCellValue(i + 1, 6, y.Depart);
                sl.SetCellValue(i + 1, 7, y.doctp);
                sl.SetCellValue(i + 1, 8, y.Stime.ToString("yyyy-MM-dd"));
                sl.SetCellValue(i + 1, 9, y.Rtime.ToString("yyyy-MM-dd"));
                sl.SetCellValue(i + 1, 10, y.Ntime.ToString("yyyy-MM-dd"));
                if (y.Ntime.AddMonths(1) < DateTime.Now)
                    sl.SetCellStyle(i + 1, 10, style);
                sl.SetCellValue(i + 1, 11, y.Own);
                //sl.SetCellValue(i + 1, 13, y.Color);
                i++;
            }
            return sl;
        }
        public void CombineData()
        {
            string folderName = System.Environment.CurrentDirectory + @"\管理員";
            ADocs.Clear();
            try
            {
                foreach (var finame in System.IO.Directory.GetFileSystemEntries(folderName))
                {
                    if (System.IO.Path.GetExtension(finame) != ".xlsx")
                        continue;
                    LoadData(finame);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        public void ExportHTML()
        {
            string htmlname = Environment.CurrentDirectory + @"\template.htm";
            if (File.Exists(htmlname))
            {
                string sw = File.ReadAllText(htmlname, Encoding.Default);
                sw = sw.Replace("!!!!!!", DateTime.Now.ToString("yyyy-MM-dd"));
                foreach (var x in ADocs)
                {
                    string content = string.Format("" +
                        "	<tr{8}>" + Environment.NewLine +
                        "        <td>{0}</td>" + Environment.NewLine +
                        "        <td>{1}</td>" + Environment.NewLine +
                        "        <td>{2}</td>" + Environment.NewLine +
                        "        <td>{3}</td>" + Environment.NewLine +
                        "        <td>{4}</td>" + Environment.NewLine +
                        "        <td>{5}</td>" + Environment.NewLine +
                        "        <td>{6}</td>" + Environment.NewLine +
                        "        <td>{7}</td>" + Environment.NewLine +
                        "   </tr>" + Environment.NewLine
                        , Int32.TryParse(x.webID, out int id) && id > 0 ? "<a href =\"http://km.sltung.com.tw/km/readdocument.aspx?documentId=" + x.webID + "\">" + x.ID + "</a>" : x.ID
                        , x.Name, x.Version, x.Depart, x.doctp, x.Stime.ToString("yyyy-MM-dd"), x.Rtime.ToString("yyyy-MM-dd"), x.Ntime.ToString("yyyy-MM-dd")
                        , x.Ntime.AddMonths(1) < DateTime.Now ? " bgcolor=\"#FFCCFF\"": "");
                    sw += content;
                }
                string foot = string.Format("" +
                    "</table>" +
                    "<p align=\"center\" />" +
                    "<img alt=\"horizontal rule\" height=\"10\" src=\"poshorsa.gif\" width=\"80%\">" +
                    "<p align=\"center\" />" +
                    "<b> 網頁異動日期：{0} </b>" +
                    "</body>" +
                    "</div>" +
                    "</html>", DateTime.Now.ToString("yyyy-MM-dd"));
                sw += foot;
                string fpath = Environment.CurrentDirectory + @"\合併檔";
                if (!System.IO.Directory.Exists(fpath))
                {
                    System.IO.Directory.CreateDirectory(fpath);
                }
                File.WriteAllText(fpath + @"\new_page_66-1.htm", sw, Encoding.Default);
                if (File.Exists(fpath + @"\new_page_66-1.htm"))
                    File.Copy(fpath + @"\new_page_66-1.htm", fpath + @"\new_page_66-1" + "(" + DateTime.Now.ToString("yyy-MM-dd") + ")" + ".htm", true);
                this.TxtBox1.Text += Environment.NewLine + "~~ 網頁匯出成功 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
            }
            else
                this.TxtBox1.Text += Environment.NewLine + "~~ 找不到網頁範本(template.htm) ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
        }
        #endregion

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ADocs.Clear();
            LoadFile();
            this.TxtBox1.Text += Environment.NewLine + "~~ 讀取結束 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            ADocs.Clear();
            CombineData();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 合併失敗 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            try
            {
                ADocs.Sort((x, y) => { return Convert.ToInt32(x.Index).CompareTo(Convert.ToInt32(y.Index)); });
                SLDocument sl = ExportExcel(ADocs);
                string fpath = Environment.CurrentDirectory + @"\合併檔";
                if (!System.IO.Directory.Exists(fpath))
                {
                    System.IO.Directory.CreateDirectory(fpath);
                }
                sl.SaveAs(fpath + @"\" + "合併總表" + ".xlsx");
                if (File.Exists(fpath + @"\" + "合併總表" + ".xlsx"))
                    File.Copy(fpath + @"\" + "合併總表" + ".xlsx", fpath + @"\" + "合併總表" + "(" + DateTime.Now.ToString("yyy-MM-dd") + ")" + ".xlsx", true);
                this.TxtBox1.Text += Environment.NewLine + "~~ 合併成功 ~~ "  + DateTime.Now.ToLongTimeString() + Environment.NewLine;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            ExportHTML();
        }

        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            if (ADocs.Count <= 0)
            {
                CombineData();
                if (ADocs.Count <= 0)
                {
                    this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取資料 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                    return;
                }
                
                try
                {
                    var exdocs = ADocs.Where(o => o.Ntime.AddMonths(1) < DateTime.Now).ToList();
                    if (exdocs.Count > 0)
                    {
                        SLDocument sl = ExportExcel(exdocs);
                        string fpath = Environment.CurrentDirectory + @"\即將過期";
                        if (!System.IO.Directory.Exists(fpath))
                        {
                            System.IO.Directory.CreateDirectory(fpath);
                        }
                        sl.SaveAs(fpath + @"\" + "即將過期" + "(" + DateTime.Now.ToString("yyy-MM-dd") + ")" + ".xlsx");
                    }
                    this.TxtBox1.Text += Environment.NewLine + string.Format("即將過期份數: {0, 5}份 ", exdocs.Count.ToString()) + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                    this.TxtBox1.Text += Environment.NewLine + "~~ 匯出成功 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
