using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Navigation;
using System.Windows.Threading;
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
            System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            System.Threading.Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture;
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }
        #region Parameter
        public List<DocType> ADocs = new List<DocType>();
        public DispatcherTimer _timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(10)
        };
        #endregion
        #region Method
        public void CleanBackup()
        {
            List<string> lists = new List<string>() { @"\合併檔", @"\管理員\備份" };
            lists.ForEach(o =>
            {
                string fpath = Environment.CurrentDirectory + o;
                if (System.IO.Directory.Exists(fpath))
                {
                    foreach (var x in System.IO.Directory.GetDirectories(fpath))
                    {
                        if (System.IO.Directory.GetCreationTime(x).AddMonths(6) < DateTime.Now)
                            System.IO.Directory.Delete(x, true);
                    }
                }
            });
            this.TxtBox1.Text += Environment.NewLine + "~~ 執行清除六個月前備份 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
        }
        public void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                /*
                if (Cb_tick.IsChecked == false)
                {
                    _timer.Interval = new TimeSpan(1, 0, 0);
                    return;
                }
                */
                if ((22 - DateTime.Now.Hour) > 0)
                {
                    _timer.Interval = DateTime.Now.AddHours(22 - DateTime.Now.Hour) - DateTime.Now + new TimeSpan(0, 0, 10);
                    this.TxtBox1.Text += Environment.NewLine + $"~~ 排程時間尚有 {22 - DateTime.Now.Hour} 小時 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                    return;
                }
                _timer.Interval = new TimeSpan(12, 0, 0);
                LoadFullDocs();
                if (ADocs.Count <= 0)
                {
                    this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取總表 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                    return;
                }
                int b = ADocs.Count;
                CombineData();
                if (ADocs.Count <= 0)
                {
                    this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取資料 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                    return;
                }
                int a = ADocs.Count;
                int c = 0;
                ADocs.ForEach(o =>
                {
                    if (o.Rtime > DateTime.Now)
                        this.TxtBox1.Text += Environment.NewLine + $"~~ 檢視日期錯誤 ~~({o.ID} : {o.Rtime})" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                });
                if (c > 0)
                {
                    this.TxtBox1.Text += Environment.NewLine + "~~ 日期錯誤無法合併 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                    return;
                }
                if (a == b)
                {
                    if (ExportAllExcel())
                        ExportOwn();
                    if (ExportHTML())
                        if (!ExportToWeb())
                            _timer.Interval = new TimeSpan(1, 0, 0);
                    if (DateTime.Now.Day == 1)
                        CleanBackup();
                }
                else
                {
                    this.TxtBox1.Text += Environment.NewLine + "~~ 文件總數不符，請確認確認文件是否有刪減 ~~ ";
                    this.TxtBox1.Text += Environment.NewLine + "~~ 管理員: " + a + " 件 ~~ ";
                    this.TxtBox1.Text += Environment.NewLine + "~~ 總表: " + b + " 件 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        public List<DocType> ImportEXCEL(string fname)
        {
            if (!System.IO.File.Exists(fname))
                return new List<DocType>(); ;
            try
            {
                SLDocument sl = new SLDocument(fname, "工作表1");
                SLWorksheetStatistics stats = sl.GetWorksheetStatistics();
                if (stats.EndRowIndex <= 0)
                    return new List<DocType>();
                List<DocType> ndocs = new List<DocType>();
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
                    if (DateTime.TryParseExact(sl.GetCellValueAsString(i + 2, 8), "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime _date))
                    {
                        docs.Stime = _date;
                    }
                    if (DateTime.TryParseExact(sl.GetCellValueAsString(i + 2, 9), "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _date))
                    {
                        docs.Rtime = _date;
                    }
                    if (sl.GetCellStyle(i + 2, 2).Font.FontColor == System.Drawing.Color.FromArgb(255, 255, 0, 0))
                        docs.Eng = true;
                    if (docs.webID.Contains("documentId"))
                    {
                        var m = Regex.Match(docs.webID, @"documentId=(\d+)");
                        if (m.Success)
                        {
                            docs.webID = m.Groups[1].ToString();
                        }
                    }
                    else if (Int64.TryParse(sl.GetCellValueAsString(i + 2, 3).Trim(), out long id2))
                    {
                        docs.webID = id2.ToString();
                    }
                    else
                        docs.webID = "-1";
                    if (Int64.TryParse(sl.GetCellValueAsString(i + 2, 1).Trim(), out long id))
                    {
                        docs.Index = id.ToString();
                    }
                    else
                        docs.Index = "-1";
                    if (Double.TryParse(sl.GetCellValueAsString(i + 2, 5).Trim(), out double id3))
                    {
                        docs.Version = id3.ToString();
                    }
                    else
                        docs.Version = "-1";
                    ndocs.Add(docs);
                }
                sl.CloseWithoutSaving();
                return ndocs;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new List<DocType>();
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
        public void LoadFullDocs()
        {
            ADocs.Clear();
            string fpath = Environment.CurrentDirectory;
            if (!System.IO.Directory.Exists(fpath))
            {
                System.IO.Directory.CreateDirectory(fpath);
            }
            string fname = fpath + @"\合併檔\合併總表.xlsx";
            if (!System.IO.File.Exists(fname))
                return;
            ADocs = ImportEXCEL(fname);
            this.TxtBox1.Text += Environment.NewLine + "~~ 載入總表 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
        }
        public void ExportOwn()
        {
            if (ADocs.Count > 0)
            {
                var odocs = ADocs.GroupBy(o => o.Own).ToDictionary(o => o.Key, o => o.ToList());
                odocs = odocs.OrderBy(o => o.Key).ToDictionary(o => o.Key, o => o.Value);
                try
                {
                    string fpath = Environment.CurrentDirectory + @"\管理員";
                    if (!System.IO.Directory.Exists(fpath))
                    {
                        System.IO.Directory.CreateDirectory(fpath);
                    }
                    foreach (var finame in System.IO.Directory.GetFileSystemEntries(fpath))
                    {
                        if (System.IO.Path.GetExtension(finame) != ".xlsx")
                            continue;
                        string nfpath = Environment.CurrentDirectory + @"\管理員\備份";
                        if (!System.IO.Directory.Exists(nfpath))
                        {
                            System.IO.Directory.CreateDirectory(nfpath);
                        }
                        nfpath = Environment.CurrentDirectory + @"\管理員\備份\" + DateTime.Now.ToString("yyy-MM-dd-HH-mm-ss");
                        if (!System.IO.Directory.Exists(nfpath))
                        {
                            System.IO.Directory.CreateDirectory(nfpath);
                        }
                        string fname = System.IO.Path.GetFileNameWithoutExtension(finame);
                        File.Copy(fpath + @"\" + fname + ".xlsx", nfpath + @"\" + fname + "(" + DateTime.Now.ToString("yyy-MM-dd-HH-mm-ss") + ")" + ".xlsx", true);
                        File.Delete(fpath + @"\" + fname + ".xlsx");
                    }
                    foreach (var x in odocs)
                    {
                        SLDocument sl = MakeEXCEL(x.Value);
                        sl.SaveAs(fpath + @"\" + x.Key + ".xlsx");
                        this.TxtBox1.Text += Environment.NewLine
                            + x.Key + string.Format(" 負責: {0, 5} 份", x.Value.Count.ToString());
                    }
                    this.TxtBox1.Text += Environment.NewLine
                        + "      " + string.Format(" 總共: {0, 5} 份", ADocs.Count.ToString()) + Environment.NewLine;
                    this.TxtBox1.Text += Environment.NewLine + "~~ 分派檔案結束 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        public SLDocument MakeEXCEL(List<DocType> dts)
        {
            int i = 1;
            SLDocument sl = new SLDocument();
            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "工作表1");

            SLConditionalFormatting cf = new SLConditionalFormatting("C2", "C" + (dts.Count + 1).ToString());
            cf.HighlightCellsWithDuplicates(SLHighlightCellsStyleValues.LightRedFill);
            cf.HighlightCellsEqual(true, "-1", SLHighlightCellsStyleValues.LightRedFill);
            sl.AddConditionalFormatting(cf);
            cf = new SLConditionalFormatting("B2", "B" + (dts.Count + 1).ToString());
            cf.HighlightCellsWithDuplicates(SLHighlightCellsStyleValues.LightRedFill);
            cf.HighlightCellsEqual(true, "0", SLHighlightCellsStyleValues.LightRedFill);
            sl.AddConditionalFormatting(cf);
            cf = new SLConditionalFormatting("A2", "A" + (dts.Count + 1).ToString());
            cf.HighlightCellsWithDuplicates(SLHighlightCellsStyleValues.LightRedFill);
            cf.HighlightCellsEqual(true, "0", SLHighlightCellsStyleValues.LightRedFill);
            sl.AddConditionalFormatting(cf);
            cf = new SLConditionalFormatting("J2", "J" + (dts.Count + 1).ToString());
            cf.HighlightCellsWithFormula("=DATE(YEAR($J2),MONTH($J2)-1,DAY($J2)) <= TODAY()", SLHighlightCellsStyleValues.LightRedFill);
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
            sl.SetCellValue(1, 9, "最近檢視時間");
            sl.SetCellValue(1, 10, "預計檢視時間");
            sl.SetCellValue(1, 11, "負責同仁");
            sl.SetCellValue(1, 12, "備註");
            sl.SetCellValue(1, 13, "備註(2)");
            style.Font.FontColor = System.Drawing.Color.Red;
            foreach (var y in dts)
            {
                sl.SetCellValue(i + 1, 1, Convert.ToInt64(y.Index));
                sl.SetCellValue(i + 1, 2, y.ID);
                if (y.Eng)
                    sl.SetCellStyle(i + 1, 2, style);
                sl.SetCellValue(i + 1, 3, Convert.ToInt64(y.webID));
                sl.SetCellValue(i + 1, 4, y.Name);
                sl.SetCellValue(i + 1, 5, Convert.ToDouble(y.Version));
                sl.SetCellValue(i + 1, 6, y.Depart);
                sl.SetCellValue(i + 1, 7, y.doctp);
                sl.SetCellValue(i + 1, 8, y.Stime);
                sl.SetCellValue(i + 1, 9, y.Rtime);
                //sl.SetCellValue(i + 1, 10, y.Ntime.ToString("yyy-MM-dd"));
                sl.SetCellValue(i + 1, 10, string.Format("=IF(I{0}=\"\",\"\",DATE(YEAR(I{0})+1,MONTH(I{0}),DAY(I{0})))", i + 1));
                if (y.Ntime.AddMonths(-1) < DateTime.Now)
                {
                    sl.SetCellStyle(i + 1, 5, style);
                    sl.SetCellStyle(i + 1, 9, style);
                    sl.SetCellStyle(i + 1, 10, style);
                }
                sl.SetCellValue(i + 1, 11, y.Own);
                //sl.SetCellValue(i + 1, 13, y.Color);
                
                SLStyle st2 = sl.CreateStyle();
                st2.FormatCode = "#,##0.0";
                sl.SetCellStyle(i + 1, 5, st2);
                st2.FormatCode = "yyyy/mm/dd";
                sl.SetCellStyle(i + 1, 8, st2);
                sl.SetCellStyle(i + 1, 9, st2);
                sl.SetCellStyle(i + 1, 10, st2);
                SLStyle stp = sl.CreateStyle();
                stp.Protection.Locked = false;
                stp.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 204, 255, 255), System.Drawing.Color.DarkSalmon);
                stp.SetRightBorder(BorderStyleValues.Thin, System.Drawing.Color.DarkSalmon);
                stp.SetLeftBorder(BorderStyleValues.Thin, System.Drawing.Color.DarkSalmon);
                stp.SetTopBorder(BorderStyleValues.Thin, System.Drawing.Color.DarkSalmon);
                stp.SetBottomBorder(BorderStyleValues.Thin, System.Drawing.Color.DarkSalmon);
                sl.SetCellStyle(i + 1, 5, stp);
                sl.SetCellStyle(i + 1, 9, stp);

                i++;
            }
            SLSheetProtection sp = new SLSheetProtection();
            sp.AllowInsertRows = false;
            sp.AllowInsertColumns = false;
            sp.AllowFormatCells = true;
            sp.AllowDeleteColumns = false;
            sp.AllowDeleteRows = false;
            sp.AllowSelectUnlockedCells = true;
            sp.AllowSelectLockedCells = false;
            sl.ProtectWorksheet(sp);
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
                    ADocs.AddRange(ImportEXCEL(finame));
                }
                if (ADocs.Count > 0)
                {
                    ADocs.Sort((x, y) => {
                        /*
                        Match m = Regex.Match(x.ID, @"(\d+)-");
                        Match n = Regex.Match(y.ID, @"(\d+)-");
                        if (m.Success && n.Success)
                        {
                            return Convert.ToInt32(m.Groups[1].ToString()).CompareTo(Convert.ToInt32(n.Groups[1].ToString()));
                        }
                        else
                        */
                        return x.ID.CompareTo(y.ID); 
                    });
                    for (int i = 0; i < ADocs.Count; i++)
                        ADocs[i].Index = (i + 1).ToString();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        public bool ExportHTML()
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
                        , Int64.TryParse(x.webID, out long id) && id > 0 ? "<a href =\"http://km.sltung.com.tw/km/readdocument.aspx?documentId=" + x.webID + "\" target=\"_blank\">" + x.ID + "</a>" : x.ID
                        , x.Name, string.Format("{0:0.0}", Convert.ToDouble(x.Version)), x.Depart, x.doctp, x.Stime.ToString("yyy-MM-dd"), x.Rtime.ToString("yyy-MM-dd"), x.Ntime.ToString("yyyy-MM-dd")
                        , x.Ntime < DateTime.Now ? " bgcolor=\"#FFCCFF\"" : x.Ntime.AddMonths(-1) < DateTime.Now ? " bgcolor=\"#CCFFFF\"" : "");
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
                if (File.Exists(fpath + @"\new_page_66-1.htm"))
                {
                    string nfpath = Environment.CurrentDirectory + @"\合併檔\備份" + DateTime.Now.ToString("yyy-MM-dd"); ;
                    if (!System.IO.Directory.Exists(nfpath))
                    {
                        System.IO.Directory.CreateDirectory(nfpath);
                    }
                    File.Copy(fpath + @"\new_page_66-1.htm", nfpath + @"\new_page_66-1" + "(" + DateTime.Now.ToString("yyy-MM-dd-HH-mm-ss") + ")" + ".htm", true);
                }
                File.WriteAllText(fpath + @"\new_page_66-1.htm", sw, Encoding.Default);
                this.TxtBox1.Text += Environment.NewLine + "~~ 網頁匯出成功 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return true;
            }
            else
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 找不到網頁範本(template.htm) ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return false;
            }
        }
        public bool ExportToWeb()
        {
            string fpath = Environment.CurrentDirectory + @"\合併檔";
            if (!System.IO.Directory.Exists(fpath) || !File.Exists(fpath + @"\new_page_66-1.htm"))
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 找不到網頁檔案 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return false;
            }
            string dpath = @"P:\d4215.web";
            if (!System.IO.Directory.Exists(dpath) || !File.Exists(dpath + @"\new_page_66-1.htm"))
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 找不到網頁位置 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return false;
            }
            /*
            string ndpath = dpath + @"\backup\備份";
            if (!System.IO.Directory.Exists(ndpath))
            {
                System.IO.Directory.CreateDirectory(ndpath);
            }
            File.Copy(dpath + @"\new_page_66-1.htm", ndpath + @"\new_page_66-1" + "(" + DateTime.Now.ToString("yyy-MM-dd-HH-mm-ss") + ")" + ".htm", true);
            */
            File.Copy(fpath + @"\new_page_66-1.htm", dpath + @"\new_page_66-1.htm", true);
            this.TxtBox1.Text += Environment.NewLine + "~~ 網頁匯至網站成功 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
            return true;
        }
        public bool ExportAllExcel()
        {
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 合併失敗 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return false;
            }
            try
            {
                string fpath = Environment.CurrentDirectory + @"\合併檔";
                if (!System.IO.Directory.Exists(fpath))
                {
                    System.IO.Directory.CreateDirectory(fpath);
                }
                if (File.Exists(fpath + @"\" + "合併總表" + ".xlsx"))
                {
                    string nfpath = Environment.CurrentDirectory + @"\合併檔\備份" + DateTime.Now.ToString("yyy-MM-dd"); ;
                    if (!System.IO.Directory.Exists(nfpath))
                    {
                        System.IO.Directory.CreateDirectory(nfpath);
                    }
                    File.Copy(fpath + @"\" + "合併總表" + ".xlsx", nfpath + @"\" + "合併總表" + "(" + DateTime.Now.ToString("yyy-MM-dd-HH-mm-ss") + ")" + ".xlsx", true);
                }
                SLDocument sl = MakeEXCEL(ADocs);
                sl.SaveAs(fpath + @"\" + "合併總表" + ".xlsx");
                this.TxtBox1.Text += Environment.NewLine + "~~ 合併成功 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
        #endregion

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CombineData();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取資料 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            int a = ADocs.Count;
            LoadFullDocs();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取總表 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            int b = ADocs.Count;
            if (MessageBox.Show("是否確定產生分派檔案?", "分派文件資料", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
            {
                if ((a == b && a > 0) || (a != b && MessageBox.Show("文件總數不符(請確定是否有刪減)，是否確定合併?", "文件總數不符", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes))
                {
                    ExportOwn();
                }
                else
                {
                    this.TxtBox1.Text += Environment.NewLine + "~~ 文件總數不符，請再次確認文件是否有刪減 ~~ ";
                    this.TxtBox1.Text += Environment.NewLine + "~~ 管理員: " + a + " 件 ~~ ";
                    this.TxtBox1.Text += Environment.NewLine + "~~ 總表: " + b + " 件 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                }
            }
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            LoadFullDocs();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取總表 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            int b = ADocs.Count;
            CombineData();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取資料 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            int a = ADocs.Count;
            int c = 0;
            ADocs.ForEach(o =>
            {
                if (o.Rtime > DateTime.Now)
                    this.TxtBox1.Text += Environment.NewLine + $"~~ 檢視日期錯誤 ~~({o.ID} : {o.Rtime})" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
            });
            if (c > 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 日期錯誤無法合併 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            if (MessageBox.Show("是否確定合併?", "合併各負責人文件資料", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
            {
                if ((a == b && a > 0) || (a != b && MessageBox.Show(string.Format("文件總數不符(請確定是否有刪減)，是否確定合併?\n管理員: {0} 件\n總表: {1} 件", a, b), "文件總數不符", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes))
                {
                    if (ExportAllExcel())
                    {
                        ExportOwn();
                    }
                    if (ExportHTML())
                    {
                        ExportToWeb();
                    }
                    if (DateTime.Now.Day == 1)
                        CleanBackup();
                }
                else
                {
                    this.TxtBox1.Text += Environment.NewLine + "~~ 文件總數不符，請再次確認文件是否有刪減 ~~ ";
                    this.TxtBox1.Text += Environment.NewLine + "~~ 管理員: " + a + " 件 ~~ ";
                    this.TxtBox1.Text += Environment.NewLine + "~~ 總表: " + b + " 件 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                }
            }
        }

        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            LoadFullDocs();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取總表 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }

            try
            {
                var exdocs = ADocs.Where(o => o.Ntime.AddMonths(-1) < DateTime.Now).ToList();
                if (exdocs.Count > 0)
                {
                    SLDocument sl = MakeEXCEL(exdocs);
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

        private void Button4_Click(object sender, RoutedEventArgs e)
        {
            CombineData();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取資料 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            int a = ADocs.Count;
            LoadFullDocs();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取總表 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            int b = ADocs.Count;
            if (a == b && a > 0)
            {
                try
                {
                    string fpath = Environment.CurrentDirectory + @"\部門";
                    if (!System.IO.Directory.Exists(fpath))
                    {
                        System.IO.Directory.CreateDirectory(fpath);
                    }
                    var odocs = ADocs.GroupBy(o => o.Own).ToDictionary(o => o.Key, o => o.ToList());
                    foreach (var x in odocs)
                    {
                        fpath = Environment.CurrentDirectory + @"\部門\" + x.Key;
                        if (!System.IO.Directory.Exists(fpath))
                        {
                            System.IO.Directory.CreateDirectory(fpath);
                        }
                        foreach (var finame in System.IO.Directory.GetFileSystemEntries(fpath))
                        {
                            if (System.IO.Path.GetExtension(finame) != ".xlsx")
                                continue;
                            string nfpath = fpath + @"\備份";
                            if (!System.IO.Directory.Exists(nfpath))
                            {
                                System.IO.Directory.CreateDirectory(nfpath);
                            }
                            string fname = System.IO.Path.GetFileNameWithoutExtension(finame);
                            File.Copy(fpath + @"\" + fname + ".xlsx", nfpath + @"\" + fname + ".xlsx", true);
                            File.Delete(fpath + @"\" + fname + ".xlsx");
                        }
                        var ddocs = x.Value.GroupBy(o => o.Depart).ToDictionary(o => o.Key, o => o.ToList());
                        foreach (var y in ddocs)
                        {
                            SLDocument sl = MakeEXCEL(y.Value);
                            sl.SaveAs(fpath + @"\" + y.Key + "(原-" + x.Key + ")" + ".xlsx");
                        }
                    }
                    this.TxtBox1.Text += Environment.NewLine + "~~ 匯出部門成功 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 文件總數不符，請再次確認文件是否有刪減 ~~ ";
                this.TxtBox1.Text += Environment.NewLine + "~~ 管理員: " + a + " 件 ~~ ";
                this.TxtBox1.Text += Environment.NewLine + "~~ 總表: " + b + " 件 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
            }
        }

        private void Button5_Click(object sender, RoutedEventArgs e)
        {
            CombineData();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取資料 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            int a = ADocs.Count;
            LoadFullDocs();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取總表 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            int b = ADocs.Count;
            try
            {
                string fpath = Environment.CurrentDirectory + @"\部門";
                if (!System.IO.Directory.Exists(fpath))
                {
                    this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取部門資料 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                    return;
                }
                List<DocType> cdocs = new List<DocType>();
                foreach (var fp in System.IO.Directory.GetDirectories(fpath))
                {
                    var npath = fpath + @"\" + Path.GetFileName(fp);
                    int count = 0;
                    foreach (var finame in System.IO.Directory.GetFileSystemEntries(npath))
                    {
                        if (System.IO.Path.GetExtension(finame) != ".xlsx")
                            continue;
                        var ndocs = ImportEXCEL(finame);
                        ndocs.ForEach(o => o.Own = Path.GetFileName(fp));
                        cdocs.AddRange(ndocs);
                        count += ndocs.Count;
                    }
                    this.TxtBox1.Text += Environment.NewLine
                            + Path.GetFileName(fp) + string.Format(" 負責: {0, 5} 份", count.ToString());
                }
                this.TxtBox1.Text += Environment.NewLine
                        + "      " + string.Format(" 總共: {0, 5} 份", cdocs.Count.ToString()) + Environment.NewLine;
                int c = cdocs.Count;
                if (c > 0 && a == b && b == c && c == a)
                {
                    ADocs.Clear();
                    ADocs = cdocs;
                    ADocs.Sort((x, y) =>
                    {
                        /*
                        Match m = Regex.Match(x.ID, @"(\d+)-");
                        Match n = Regex.Match(y.ID, @"(\d+)-");
                        if (m.Success && n.Success)
                        {
                            return Convert.ToInt32(m.Groups[1].ToString()).CompareTo(Convert.ToInt32(n.Groups[1].ToString()));
                        }
                        else
                        */
                        return x.ID.CompareTo(y.ID);
                    });
                    for (int i = 0; i < ADocs.Count; i++)
                        ADocs[i].Index = (i + 1).ToString();
                    if (MessageBox.Show("是否確定合併?", "合併重分配文件資料", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
                    {
                        if (ExportAllExcel())
                            this.TxtBox1.Text += Environment.NewLine + "~~ 文件重新分派成功 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                    }
                    else
                        this.TxtBox1.Text += Environment.NewLine;
                }
                else
                {
                    this.TxtBox1.Text += Environment.NewLine + "~~ 文件總數不符，請重新分派 ~~ ";
                    this.TxtBox1.Text += Environment.NewLine + "~~ 管理員: " + a + " 件 ~~ ";
                    this.TxtBox1.Text += Environment.NewLine + "~~ 總表: " + b + " 件 ~~ ";
                    this.TxtBox1.Text += Environment.NewLine + "~~ 合併後: " + c + " 件 ~~ " + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Button6_Click(object sender, RoutedEventArgs e)
        {
            LoadFullDocs();
            if (ADocs.Count <= 0)
            {
                this.TxtBox1.Text += Environment.NewLine + "~~ 無法讀取總表 ~~" + DateTime.Now.ToLongTimeString() + Environment.NewLine;
                return;
            }
            var odocs = ADocs.GroupBy(o => o.Own).ToDictionary(o => o.Key, o => o.ToList());
            odocs = odocs.OrderBy(o => o.Key).ToDictionary(o => o.Key, o => o.Value);
            try
            {
                foreach (var x in odocs)
                {
                    this.TxtBox1.Text += Environment.NewLine
                        + x.Key + string.Format(" 負責: {0, 5} 份", x.Value.Count.ToString());
                }
                this.TxtBox1.Text += Environment.NewLine
                        + "      " + string.Format(" 總共: {0, 5} 份", ADocs.Count.ToString()) + Environment.NewLine;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TxtBox1_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            TxtBox1.CaretIndex = TxtBox1.Text.Length;
            TxtBox1.ScrollToEnd();
        }

        private void Cb_tick_Checked(object sender, RoutedEventArgs e)
        {
            _timer.Start();
        }

        private void Cb_tick_Unchecked(object sender, RoutedEventArgs e)
        {
            _timer.Stop();
        }
    }
}
