using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Threading;
using DocumentFormat.OpenXml.Packaging;
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
        //public string Index { get; set; }
        public string ID { get; set; }
        public string webID { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }
        public string Depart { get; set; }
        public string doctp { get; set; }
        public DateTime Stime { get; set; }
        public DateTime Rtime { get; set; }
        public DateTime Ntime { get; set; }
        public DateTime Etime { get; set; }
        public string Own { get; set; }
        public bool Eng { get; set; }
        public string Color { get; set; }
        public bool Invalid { get; set; }
        public List<DocHistory> History { get; set; }
        public DocType()
        {
            History = new List<DocHistory>();
        }
    }
    public class DocHistory
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public DateTime Rtime { get; set; }
    }
    public class DocNum
    {
        public string Title { get; set; }
        public int Docnumber { get; set; }
        public int ntime { get; set; }
        public int etime { get; set; }
        public string hTitle { get; set; }
        public string Renew { get; set; }
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
            //_timer.Start();
            List<string> dtype = new List<string>() { "跨部門文件", "部門內文件" };
            CB_Doctype.ItemsSource = dtype;
            CB_Doctype.SelectedIndex = 0;
            LB_build.Content = "編譯時間: " + System.IO.File.GetLastWriteTime(this.GetType().Assembly.Location).ToString("yyyy-MM-dd");
        }
        #region Parameter
        public List<DocType> ADocs = new List<DocType>();
        public DispatcherTimer _timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(10)
        };
        public DocNum Dnum = new DocNum();
        public bool DataChange = false;
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
            ShowMessage("執行清除六個月前備份");
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
                    _timer.Interval = DateTime.Today.AddHours(22) - DateTime.Now + new TimeSpan(0, 0, 10);
                    ShowMessage($"排程時間尚有 {22 - DateTime.Now.Hour} 小時");
                    return;
                }
                _timer.Interval = new TimeSpan(12, 0, 0);
                ADocs.Clear();
                ADocs = LoadFullDocs();
                if (ADocs.Count <= 0)
                {
                    ShowMessage($"無法讀取總表({Dnum.Title})");
                    return;
                }
                var combine = LoadOwn();
                if (combine.Count <= 0)
                {
                    ShowMessage("無法讀取資料");
                    return;
                }
                if (!DataChange)
                {
                    ShowMessage("檔案似無異動，暫停更新一次 ");
                    return;
                }
                ADocs.ForEach(o =>
                {
                    if (o.Rtime > DateTime.Now || o.Rtime == DateTime.MinValue)
                    {
                        ShowMessage($"檢視日期錯誤({o.ID} : {o.Rtime})");
                        ShowMessage("日期錯誤無法合併");
                        return;
                    }
                });
                /*
                 * 合併新舊資料庫
                 */
                if (ADocs.Where(o => o.Invalid == false).ToList().Count <= combine.Count)
                {
                    if (ExportAllExcel())
                        ExportOwn();
                    if (ExportHTML())
                        if (!ExportToWeb())
                            _timer.Interval = new TimeSpan(1, 0, 0);
                    if (DateTime.Now.Day == 1)
                        CleanBackup();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        public List<DocType> ImportEXCEL(string fname, int ntime, int etime)
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
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 2)))
                        break;
                    DocType docs = new DocType
                    {
                        //Index = sl.GetCellValueAsString(i + 2, 1),
                        ID = sl.GetCellValueAsString(i + 2, 2).Trim(),
                        Color = sl.GetCellStyle(i + 2, 2).Font.FontColor.ToString(),
                        webID = sl.GetCellValueAsString(i + 2, 13).Trim(),
                        Name = sl.GetCellValueAsString(i + 2, 4).Trim(),
                        Version = sl.GetCellValueAsString(i + 2, 5).Trim(),
                        Depart = sl.GetCellValueAsString(i + 2, 6).Trim(),
                        doctp = sl.GetCellValueAsString(i + 2, 7).Trim(),
                        Stime = sl.GetCellValueAsDateTime(i + 2, 8),
                        Rtime = sl.GetCellValueAsDateTime(i + 2, 9),
                        Own = sl.GetCellValueAsString(i + 2, 11).Trim(),
                        Invalid = sl.GetCellValueAsString(i + 2, 12).Trim() == "廢止"
                    };
                    int h = 14;
                    while (!String.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, h).Trim()))
                    {
                        string[] hisdocs = sl.GetCellValueAsString(i + 2, h).Trim().Split(';');
                        if (hisdocs.Length == 3)
                        {
                            List<DocHistory> historys = new List<DocHistory>();
                            if (Double.TryParse(hisdocs[1].Trim(), out double hid)
                                && DateTime.TryParseExact(hisdocs[2].Trim(), "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime hrtime))
                            {
                                historys.Add(new DocHistory()
                                {
                                    Name = hisdocs[0].Trim(),
                                    Version = hid.ToString(),
                                    Rtime = hrtime
                                });
                            }
                        }
                        h++;
                    }
                    docs.Ntime = docs.Rtime.AddYears(ntime);
                    docs.Etime = docs.Rtime.AddYears(etime);
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
                    /*if (Int64.TryParse(sl.GetCellValueAsString(i + 2, 1).Trim(), out long id))
                    {
                        docs.Index = id.ToString();
                    }
                    else
                        docs.Index = "-1";*/
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
        public bool CheckRule(List<DocType> result)
        {
            Dictionary<string, DocType> CheckDocs = new Dictionary<string, DocType>();
            Dictionary<string, DocType> CheckWebID = new Dictionary<string, DocType>();
            Dictionary<string, DocType> CheckName = new Dictionary<string, DocType>();
            try
            {
                bool check = false;
                result.ForEach(o => {
                    if (!CheckDocs.ContainsKey(o.ID))
                    {
                        CheckDocs.Add(o.ID, o);
                    }
                    else
                    {
                        ShowMessage($"重複文件編號: {o.ID}-{CheckDocs[o.ID].Name} / {o.Name} ({o.Own})");
                        check = true;
                    }
                    if (o.Rtime > DateTime.Now)
                    {
                        ShowMessage($"檢視時間錯誤: {o.ID}-{o.Name} {o.Rtime:yyyy-MM-dd} ({o.Own})");
                        check = true;
                    }
                    if (o.webID == "-1")
                    {
                        ShowMessage($"網頁代碼錯誤: {o.ID}-{o.Name} ({o.Own})");
                        check = true;
                    }
                    else if (!CheckWebID.ContainsKey(o.webID))
                    {
                        CheckWebID.Add(o.webID, o);
                    }
                    else
                    {
                        ShowMessage($"網頁代碼重覆: {o.webID}-{CheckWebID[o.webID].ID}-{CheckWebID[o.webID].Name}({CheckWebID[o.webID].Own}) / {o.ID}-{o.Name}({o.Own})");
                        check = true;
                    }
                    if (!CheckName.ContainsKey(o.ID))
                    {
                        CheckName.Add(o.ID, o);
                    }
                    else
                    {
                        ShowMessage($"重複文件名稱: {o.Name}-{CheckName[o.ID].ID} / {o.ID} ({o.Own})");
                        check = true;
                    }
                });
                if (check)
                {
                    ShowMessage("!! 資料格式轉換錯誤 !!");
                    return false;
                }
                else
                    return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                ShowMessage("資料格式轉換錯誤(可能有相同文件編號)");
                return false;
            }
        }
        public List<DocType> LoadFullDocs()
        {
            string fpath = Environment.CurrentDirectory;
            if (!System.IO.Directory.Exists(fpath))
            {
                System.IO.Directory.CreateDirectory(fpath);
            }

            string fname = fpath + @"\合併檔\合併總表-" + Dnum.Title + ".xlsx";
            if (!System.IO.File.Exists(fname))
                return new List<DocType>();
            var result = new List<DocType>();
            result = ImportEXCEL(fname, Dnum.ntime, Dnum.etime);
            result.Sort((x, y) => {
                return x.ID.CompareTo(y.ID);
            });
            var numericList = result.Where(i => int.TryParse(i.ID.Split('-')[0], out _)).OrderBy(j => int.Parse(j.ID.Split('-')[0])).ToList();
            var nonNumericList = result.Where(i => !int.TryParse(i.ID.Split('-')[0], out _)).OrderBy(j => j.ID).ToList();
            if (result.Count > 0 && numericList.Count + nonNumericList.Count == result.Count)
            {
                result.Clear();
                result.AddRange(numericList);
                result.AddRange(nonNumericList);
            }
            if (CheckRule(result))
            {
                ShowMessage($"載入總表({Dnum.Title})，生效中總份數: {result.Count} 份，廢止總份數: {result.Where(o => o.Invalid).ToList().Count} 份");
                return result;
            }
            else
            {
                ShowMessage($"載入總表({Dnum.Title})失敗");
                return new List<DocType>();
            }
        }
        public List<DocType> LoadOwn()
        {
            DataChange = false;
            string fpath = System.Environment.CurrentDirectory + @"\管理員\" + Dnum.Title;
            if (!System.IO.Directory.Exists(fpath))
            {
                System.IO.Directory.CreateDirectory(fpath);
            }
            var result = new List<DocType>();
            try
            {
                DateTime dts = DateTime.MinValue;
                foreach (var finame in System.IO.Directory.GetFileSystemEntries(fpath))
                {
                    if (System.IO.Path.GetExtension(finame) != ".xlsx" || System.IO.Path.GetFileNameWithoutExtension(finame).Contains("~$"))
                        continue;
                    if (dts == DateTime.MinValue)
                        dts = System.IO.File.GetLastWriteTime(finame);
                    if (System.IO.File.GetLastWriteTime(finame) > dts.AddHours(1) || System.IO.File.GetLastWriteTime(finame).AddHours(1) < dts)
                        DataChange = true;

                    dts = System.IO.File.GetLastWriteTime(finame);
                    result.AddRange(ImportEXCEL(finame, Dnum.ntime, Dnum.etime));
                    result.Sort((x, y) => {
                        return x.ID.CompareTo(y.ID);
                    });
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            if (!CheckRule(result))
            {
                return new List<DocType>();
            }
            return result;
        }
        public SLDocument MakeEXCEL(List<DocType> dts)
        {
            return MakeEXCEL(dts, false);
        }
        public SLDocument MakeEXCEL(List<DocType> dts, bool IncludeInvalid)
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
            cf = new SLConditionalFormatting("L2", "L" + (dts.Count + 1).ToString());
            cf.HighlightCellsEqual(true, "廢止", SLHighlightCellsStyleValues.LightRedFill);
            sl.AddConditionalFormatting(cf);
            /*
            cf = new SLConditionalFormatting("A2", "A" + (dts.Count + 1).ToString());
            cf.HighlightCellsWithDuplicates(SLHighlightCellsStyleValues.LightRedFill);
            cf.HighlightCellsEqual(true, "0", SLHighlightCellsStyleValues.LightRedFill);
            sl.AddConditionalFormatting(cf);
            */
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
            sl.SetCellStyle(1, 14, style);
            sl.SetCellStyle("L2", "L" + (dts.Count + 1).ToString(), style);
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
            sl.SetColumnWidth(14, 50);
            sl.SetCellValue(1, 1, "表單分類");
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
            sl.SetCellValue(1, 12, "廢止");
            sl.SetCellValue(1, 13, "備註");
            sl.SetCellValue(1, 14, "歷程記錄");
            style.Font.FontColor = System.Drawing.Color.Red;
            SLStyle st1 = sl.CreateStyle();
            st1.Alignment.WrapText = true;
            st1.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            st1.Alignment.Vertical = VerticalAlignmentValues.Center;
            st1.Font.FontColor = System.Drawing.Color.DarkBlue;
            foreach (var y in dts)
            {
                if (IncludeInvalid == false && y.Invalid == true)
                    continue;
                sl.SetCellValue(i + 1, 1, Dnum.Title);
                if (Dnum.Docnumber == 0)
                {
                    st1.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.LightPink, System.Drawing.Color.LightPink);
                }
                else
                {
                    st1.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.LightGreen, System.Drawing.Color.LightGreen);
                }
                sl.SetCellStyle(i + 1, 1, st1);
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
                sl.SetCellValue(i + 1, 10, string.Format("=IF(I{0}=\"\",\"\",DATE(YEAR(I{0})+{1},MONTH(I{0}),DAY(I{0})))", i + 1, y.Ntime.Year - y.Rtime.Year));
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

                SLDataValidation dv;
                dv = sl.CreateDataValidation(i + 1, 12);
                dv.AllowList("\"廢止\"", true, true);
                sl.AddDataValidation(dv);
                if (y.Invalid)
                {
                    sl.SetCellValue(i + 1, 12, "廢止");
                }
                /*
                 *寫入歷程記錄
                 */
                if (y.History?.Count > 0)
                {
                    int hnum = 14;
                    foreach (var h in y.History)
                    {
                        sl.SetCellValue(i + 1, hnum, $"{h.Name};{h.Version};{h.Rtime:yyyy-MM-dd}");
                        hnum++;
                    }
                }

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
        public bool ExportHTML()
        {
            string hname = @"\new_page_" + Dnum.hTitle;
            string tname = Environment.CurrentDirectory + @"\template.htm";

            if (File.Exists(tname))
            {
                string sw = File.ReadAllText(tname, Encoding.Default);
                sw = sw.Replace("!!!!!!", DateTime.Now.ToString("yyyy-MM-dd")).Replace("@@@@@@", Dnum.Renew).Replace("######", Dnum.Title);

                string recentdocs = string.Empty;

                foreach (var x in ADocs)
                {
                    /*
                     * 排除廢止文件
                     */
                    if (x.Invalid == true)
                        continue;

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
                        , x.Stime.AddMonths(1) > DateTime.Now ? " bgcolor=\"#F1C232\"" : x.Ntime < DateTime.Now ? " bgcolor=\"#FFCCFF\"" : x.Ntime.AddMonths(-1) < DateTime.Now ? " bgcolor=\"#CCFFFF\"" : "");
                    sw += content;
                    if (x.Rtime.AddMonths(1) > DateTime.Now)
                    {
                        recentdocs += content;
                    }
                }
                sw = sw.Replace("%%%%%%", recentdocs);
                string foot = string.Format("" +
                    "</table>" +
                    "<p align=\"center\" />" +
                    "<img alt=\"horizontal rule\" height=\"10\" src=\"poshorsa.gif\" width=\"85%\">" +
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
                if (File.Exists(fpath + hname + ".htm"))
                {
                    string nfpath = Environment.CurrentDirectory + @"\合併檔\備份" + DateTime.Now.ToString("yyy-MM-dd"); ;
                    if (!System.IO.Directory.Exists(nfpath))
                    {
                        System.IO.Directory.CreateDirectory(nfpath);
                    }
                    File.Copy(fpath + hname + ".htm", nfpath + hname + "(" + DateTime.Now.ToString("yyy-MM-dd-HH-mm-ss") + ")" + ".htm", true);
                }
                File.WriteAllText(fpath + hname + ".htm", sw, Encoding.Default);
                ShowMessage("網頁匯出成功 ");
                return true;
            }
            else
            {
                ShowMessage("找不到網頁範本(template.htm) ");
                return false;
            }
        }
        public bool ExportToWeb()
        {
            string hname = @"\new_page_" + Dnum.hTitle;

            string fpath = Environment.CurrentDirectory + @"\合併檔";
            if (!System.IO.Directory.Exists(fpath) || !File.Exists(fpath + hname + ".htm"))
            {
                ShowMessage("找不到網頁檔案 ");
                return false;
            }
            string dpath = @"P:\d4215.web";
            if (!System.IO.Directory.Exists(dpath) || !File.Exists(dpath + hname + ".htm"))
            {
                ShowMessage("找不到網頁位置 ");
                return false;
            }
            File.Copy(fpath + hname + ".htm", dpath + hname + ".htm", true);
            ShowMessage("網頁匯至網站成功 ");
            return true;
        }
        public void ExportOwn()
        {
            if (ADocs.Count > 0)
            {
                var odocs = ADocs.GroupBy(o => o.Own).ToDictionary(o => o.Key, o => o.ToList());
                odocs = odocs.OrderBy(o => o.Key).ToDictionary(o => o.Key, o => o.Value);
                try
                {
                    string fpath = Environment.CurrentDirectory + @"\管理員\" + Dnum.Title;
                    if (!System.IO.Directory.Exists(fpath))
                    {
                        System.IO.Directory.CreateDirectory(fpath);
                    }
                    foreach (var finame in System.IO.Directory.GetFileSystemEntries(fpath))
                    {
                        if (System.IO.Path.GetExtension(finame) != ".xlsx")
                            continue;
                        string nfpath = Environment.CurrentDirectory + @"\管理員\" + Dnum.Title + @"\備份" + DateTime.Now.ToString("yyy-MM-dd-HH-mm-ss");
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
                        sl.SaveAs(fpath + @"\" + Dnum.Title + "-" + x.Key + ".xlsx");
                        this.TxtBox1.Text += Environment.NewLine
                            + x.Key + string.Format(" 負責: {0, 5} 份", x.Value.Count.ToString());
                    }
                    this.TxtBox1.Text += Environment.NewLine
                        + "      " + string.Format(" 總共: {0, 5} 份", ADocs.Count.ToString()) + Environment.NewLine;
                    ShowMessage("分派檔案結束");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        public bool ExportAllExcel()
        {
            if (ADocs.Count <= 0)
            {
                ShowMessage($"無法讀取總表({Dnum.Title})");
                return false;
            }
            try
            {
                string fpath = Environment.CurrentDirectory + @"\合併檔";
                if (!System.IO.Directory.Exists(fpath))
                {
                    System.IO.Directory.CreateDirectory(fpath);
                }
                if (File.Exists(fpath + @"\合併總表-" + Dnum.Title + ".xlsx"))
                {
                    string nfpath = Environment.CurrentDirectory + @"\合併檔\備份" + DateTime.Now.ToString("yyy-MM-dd"); ;
                    if (!System.IO.Directory.Exists(nfpath))
                    {
                        System.IO.Directory.CreateDirectory(nfpath);
                    }
                    File.Copy(fpath + @"\合併總表-" + Dnum.Title + ".xlsx", nfpath + @"\合併總表-" + Dnum.Title + "(" + DateTime.Now.ToString("yyy-MM-dd-HH-mm-ss") + ")" + ".xlsx", true);
                }
                SLDocument sl = MakeEXCEL(ADocs, true);
                sl.SaveAs(fpath + @"\合併總表-" + Dnum.Title + ".xlsx");
                ShowMessage("總表合併成功(" + Dnum.Title + ") ");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
        public List<DocType> MergeData()
        {
            if (ADocs.Count <= 0)
            {
                ShowMessage($"無法讀取總表({Dnum.Title})");
                return new List<DocType>();
            }
            var CDocs = LoadOwn();
            if (CDocs.Count <= 0)
            {
                ShowMessage("無法合併分派檔");
                return new List<DocType>();
            }
            /*
             * 檢核是否有檔案被誤刪
             */
            if (CDocs.Count < ADocs.Where(o => o.Invalid == false).ToList().Count)
            {
                ShowMessage("檔案數量錯誤，請檢查是否有資料被誤刪");
                return new List<DocType>();
            }

            List<DocType> docchange = new List<DocType>();
            List<DocType> docadd = new List<DocType>();
            List<DocType> docdelete = new List<DocType>();
            List<DocType> docinvalid = new List<DocType>();
            List<DocType> docrepeated = new List<DocType>();
            List<DocType> Doc_Combine = new List<DocType>();

            var OldDocs = ADocs.Where(o => o.Invalid == false).ToDictionary(o => o.ID, o => o);
            var NewDocs = CDocs.ToDictionary(o => o.ID, o => o);
            
            /*
             * 檢核是否有檔案被誤刪
             */
            foreach (var x in OldDocs.Values)
            {
                if (!NewDocs.ContainsKey(x.ID))
                {
                    docdelete.Add(x);
                }
            }
            if (docdelete.Count > 0)
            {
                ShowMessage($"可能誤刪的檔案({string.Join(",", docdelete.Select(o => o.ID + o.Name).ToList())})");
                return new List<DocType>();
            }
            OldDocs = ADocs.ToDictionary(o => o.ID, o => o);

            foreach (var ndoc in CDocs)
            {
                if (OldDocs.ContainsKey(ndoc.ID))
                {
                    var odoc = OldDocs[ndoc.ID];
                    if (odoc.Invalid == true)
                    {
                        docrepeated.Add(ndoc);
                        continue;
                    }
                    else if (odoc.Invalid == false && ndoc.Invalid == true)
                    {
                        docinvalid.Add(ndoc);
                    }
                    /*
                     * 檢查資料無有變動
                     */
                    if (ndoc.Rtime == odoc.Rtime)
                    {
                        Doc_Combine.Add(ndoc);
                    }
                    else
                    {
                        var history = odoc.History;
                        history.Add(new DocHistory()
                        {
                            Name = odoc.Name,
                            Version = odoc.Version,
                            Rtime = odoc.Rtime
                        });
                        DocType dt = new DocType();
                        var doc_new = ndoc;
                        doc_new.History = history;

                        Doc_Combine.Add(doc_new);
                        docchange.Add(doc_new);
                    }
                }
                else
                {
                    Doc_Combine.Add(ndoc);
                    docadd.Add(ndoc);
                }
            }
            if (Doc_Combine.Count <= 0)
            {
                ShowMessage($"資料合併錯誤");
                return new List<DocType>();
            }
            if (docrepeated.Count > 0)
            {
                ShowMessage($"已廢止的重複檔案({string.Join(",", docrepeated.Select(o => o.ID + o.Name).ToList())})");
            }
            if (docchange.Count > 0 || docadd.Count > 0 || docinvalid.Count > 0)
            {
                ShowMessage($"變更的檔案({string.Join(",", docchange.Select(o => o.ID + o.Name).ToList())})");
                ShowMessage($"新增的檔案({string.Join(",", docadd.Select(o => o.ID + o.Name).ToList())})");
                ShowMessage($"廢止的檔案({string.Join(",", docinvalid.Select(o => o.ID + o.Name).ToList())})");
            }
            else
            {
                ShowMessage("檔案似乎未有任何變動，不需合併，程式自動略過合併");
                return new List<DocType>();
            }
            /*
             * 需把資料庫中已廢止的資料加回，是否會重複加入已廢止的文件？？
             */
            Doc_Combine.AddRange(ADocs.Where(o => o.Invalid == true).ToList());

            Doc_Combine.Sort((x, y) => {
                return x.ID.CompareTo(y.ID);
            });
            var numericList = Doc_Combine.Where(i => int.TryParse(i.ID.Split('-')[0], out _)).OrderBy(j => int.Parse(j.ID.Split('-')[0])).ToList();
            if (numericList.Count > 0 && numericList.Count == Doc_Combine.Count)
            {
                Doc_Combine.Clear();
                Doc_Combine.AddRange(numericList);
            }
            return Doc_Combine;
        }
        public void ShowMessage(string message)
        {
            this.TxtBox1.Text += Environment.NewLine + $"{DateTime.Now.ToLongTimeString()}: ~~ {message}" + Environment.NewLine;
        }
        #endregion

        private void BT_ExportOwn(object sender, RoutedEventArgs e)
        {
            if (ADocs.Count <= 0)
            {
                ShowMessage($"無法讀取總表({Dnum.Title})");
                return;
            }
            ExportOwn();
        }

        private void BT_ALL(object sender, RoutedEventArgs e)
        {
            if (ADocs.Count <= 0)
            {
                ShowMessage($"無法讀取總表({Dnum.Title})");
                return;
            }
            List<DocType> doc_merge = new List<DocType>();
            try
            {
                doc_merge = MergeData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                ShowMessage("資料合併檢核錯誤");
            }
            if (doc_merge.Count <= 0)
            {
                ShowMessage("檔案合併檢核失敗");
                return;
            }

            if (MessageBox.Show("是否確定合併?", "合併各負責人文件資料", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
            {
                /*
                 * 合併至主資料庫，並更新合併檔
                 */
                ShowMessage("檔案合併開始");
                ADocs.Clear();
                ADocs.AddRange(doc_merge);

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
            /*

            int b = ADocs.Count;
            var olddocs = ADocs.Select(o => o.ID + o.Name).ToList();
            CombineData();
            if (ADocs.Count <= 0)
            {
                ShowMessage("無法讀取資料");
                if (MessageBox.Show("無原始分派資料，是否先產生初始分派檔案?", "分派文件資料", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                {
                    LoadFullDocs();
                    ExportOwn();
                    CombineData();
                }
                if (ADocs.Count <= 0)
                {
                    ShowMessage("無法讀取初始資料");
                    return;
                }
            }
            int a = ADocs.Count;
            var newdocs = ADocs.Select(o => o.ID + o.Name).ToList();
            int c = 0;
            ADocs.ForEach(o =>
            {
                if (o.Rtime > DateTime.Now)
                    ShowMessage($"檢視日期錯誤({o.ID} : {o.Rtime})");
            });
            if (c > 0)
            {
                ShowMessage("日期錯誤無法合併");
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
                    this.TxtBox1.Text += Environment.NewLine + "文件總數不符，請再次確認文件是否有刪減 ";
                    this.TxtBox1.Text += Environment.NewLine + "管理員: " + a + " 件 ";
                    ShowMessage("總表: " + b + " 件 ");
                }
                if (a != b)
                {
                    var diffdocs = newdocs.Except(olddocs).ToList();
                    if (diffdocs?.Count > 0)
                        ShowMessage($"新增的檔案({string.Join(",", diffdocs)})");
                    diffdocs = olddocs.Except(newdocs).ToList();
                    if (diffdocs?.Count > 0)
                        ShowMessage($"刪減的檔案({string.Join(",", diffdocs)})");
                }
            }
            */
        }

        private void BT_ExportExpired(object sender, RoutedEventArgs e)
        {
            if (ADocs.Count <= 0)
            {
                ShowMessage($"無法讀取總表({Dnum.Title})");
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
                    sl.SaveAs(fpath + @"\" + Dnum.Title + "-即將過期(" + DateTime.Now.ToString("yyy-MM-dd") + ")" + ".xlsx");
                }
                ShowMessage(string.Format("即將過期份數: {0, 5}份 ", exdocs.Count.ToString()));
                ShowMessage("匯出成功(" + Dnum.Title + ") ");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BT_ExportDepart(object sender, RoutedEventArgs e)
        {
            if (ADocs.Count <= 0)
            {
                ShowMessage($"無法讀取總表({Dnum.Title})");
                return;
            }

            try
            {
                string fpath = Environment.CurrentDirectory + @"\部門\" + Dnum.Title;
                if (!System.IO.Directory.Exists(fpath))
                {
                    System.IO.Directory.CreateDirectory(fpath);
                }
                var odocs = ADocs.GroupBy(o => o.Own).ToDictionary(o => o.Key, o => o.ToList());
                foreach (var x in odocs)
                {
                    fpath = Environment.CurrentDirectory + @"\部門\" + Dnum.Title + @"\" + x.Key;
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
                        SLDocument sl = MakeEXCEL(y.Value, true);
                        sl.SaveAs(fpath + @"\" + Dnum.Title + "-" + y.Key + "(原-" + x.Key + ")" + ".xlsx");
                    }
                }
                ShowMessage("匯出部門成功 ");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BT_ImportDepart(object sender, RoutedEventArgs e)
        {
            if (ADocs.Count <= 0)
            {
                ShowMessage($"無法讀取總表({Dnum.Title})");
                return;
            }
            try
            {
                string fpath = Environment.CurrentDirectory + @"\部門\" + Dnum.Title;
                if (!System.IO.Directory.Exists(fpath))
                {
                    ShowMessage("無法讀取部門資料");
                    return;
                }
                List<DocType> cdocs = new List<DocType>();

                foreach (var fp in System.IO.Directory.GetDirectories(fpath))
                {
                    var npath = fpath + @"\" + Path.GetFileName(fp);
                    int count = 0;
                    foreach (var finame in System.IO.Directory.GetFileSystemEntries(npath))
                    {
                        if (System.IO.Path.GetExtension(finame) != ".xlsx" || System.IO.Path.GetFileNameWithoutExtension(finame).Contains("~$"))
                            continue;
                        var ndocs = ImportEXCEL(finame, Dnum.ntime, Dnum.etime);
                        ndocs.ForEach(o => o.Own = Path.GetFileName(fp));
                        cdocs.AddRange(ndocs);
                        count += ndocs.Count;
                    }
                    this.TxtBox1.Text += Environment.NewLine
                            + Path.GetFileName(fp) + string.Format(" 負責: {0, 5} 份", count.ToString());
                }
                this.TxtBox1.Text += Environment.NewLine
                        + "      " + string.Format(" 總共: {0, 5} 份", cdocs.Count.ToString()) + Environment.NewLine;
                if (cdocs.Count == ADocs.Count && cdocs.Count > 0 && CheckRule(cdocs))
                {
                    var d_cdoc = cdocs.Select(o => o.ID).ToList();
                    var d_adoc = ADocs.Select(o => o.ID).ToList();
                    if (!d_cdoc.OrderBy(o => o).SequenceEqual(d_adoc.OrderBy(o => o)))
                    {
                        ShowMessage("文件編號不一致，請重新分派");
                    }
                    else if (MessageBox.Show("是否確定合併新分配部門?", "合併新分配部門", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
                    {
                        ADocs.Clear();
                        ADocs.AddRange(cdocs);
                        ADocs.Sort((x, y) => {
                            return x.ID.CompareTo(y.ID);
                        });
                        if (ExportAllExcel())
                            ShowMessage("新分配部門合併成功 ");
                    }
                }
                else
                {
                    ShowMessage("文件總數不符，請重新分派");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BT_CalCount(object sender, RoutedEventArgs e)
        {
            if (ADocs.Count <= 0)
            {
                ShowMessage($"無法讀取總表({Dnum.Title})");
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

        private void BT_ExportSimple(object sender, RoutedEventArgs e)
        {
            if (ADocs.Count <= 0)
            {
                ShowMessage($"無法讀取總表({Dnum.Title})");
                return;
            }
            ExportAllExcel();
            ExportHTML();
        }

        private void CB_Doctype_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (CB_Doctype.SelectedIndex != 0)
            {
                this.GD_main.Background = System.Windows.Media.Brushes.Wheat;
                this.TxtBox1.Background = System.Windows.Media.Brushes.AliceBlue;
                Dnum.Title = "部門內";
                Dnum.Docnumber = CB_Doctype.SelectedIndex;
                Dnum.ntime = 2;
                Dnum.etime = 2;
                Dnum.hTitle = "66-2";
                Dnum.Renew = "每兩年";
            }
            else
            {
                this.GD_main.Background = System.Windows.Media.Brushes.LightPink ;
                this.TxtBox1.Background = System.Windows.Media.Brushes.LavenderBlush;
                Dnum.Title = "跨部門";
                Dnum.Docnumber = CB_Doctype.SelectedIndex;
                Dnum.ntime = 1;
                Dnum.etime = 3;
                Dnum.hTitle = "66-1";
                Dnum.Renew = "每年";
            }
            this.TxtBox1.Text = String.Empty;
            ADocs.Clear();
            try
            {
                ADocs = LoadFullDocs();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
