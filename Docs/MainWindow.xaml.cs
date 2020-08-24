using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml.Math;
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
                        webID = sl.GetCellValueAsString(i + 2, 13).Trim(),
                        Name = sl.GetCellValueAsString(i + 2, 4).Trim(),
                        Version = sl.GetCellValueAsString(i + 2, 5).Trim(),
                        Depart = sl.GetCellValueAsString(i + 2, 6).Trim(),
                        doctp = sl.GetCellValueAsString(i + 2, 7).Trim(),
                        Stime = sl.GetCellValueAsDateTime(i + 2, 8),
                        Rtime = sl.GetCellValueAsDateTime(i + 2, 9),
                        Own = sl.GetCellValueAsString(i + 2, 11).Trim()
                    };
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

                   ADocs.Add(docs);
                }
                sl.CloseWithoutSaving();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            if (ADocs.Count > 0)
            {
                var odocs = ADocs.GroupBy(o => o.Own).ToDictionary(o => o.Key, o => o.ToList());
                MessageBox.Show(string.Join(";", odocs.Select(o => o.Key)));
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
                            sl.SetCellValue(i + 1, 7, y.Stime.ToString("yyyy-MM-dd"));
                            i++;
                        }
                        string fpath = Environment.CurrentDirectory + @"\管理員";
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
            }
        }
        public void LoadFile()
        {
            string fpath = Environment.CurrentDirectory;
            if (!System.IO.Directory.Exists(fpath))
            {
                System.IO.Directory.CreateDirectory(fpath);
            }
            string fname = fpath + @"\文件總表.xlsx";
            if (!System.IO.File.Exists(fname))
                return;
            LoadData(fname);
            ExportOwn();
        }
        public void ExportOwn()
        {

        }
        public void ExportDepart()
        {

        }
        public void ExportHTML()
        {

        }
        #endregion

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ADocs.Clear();
            LoadFile();
        }
    }
}
