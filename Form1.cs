using ClosedXML.Excel;
using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.LookAndFeel;
using DevExpress.Xpo.DB;
using DevExpress.XtraEditors;
using DocumentFormat.OpenXml.Office2013.Word;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Keys = OpenQA.Selenium.Keys;
using System.IO.Compression;
using System.Text.RegularExpressions;
using DevExpress.ChartRangeControlClient.Core;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.Internal.WinApi.Windows.UI.Notifications;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Toolkit.Uwp.Notifications;
using static SaovietTax.frmMain;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using System.Diagnostics;
using System.Globalization;
using DevExpress.Utils;
using Windows.UI.Xaml.Controls;
using DocumentFormat.OpenXml.Spreadsheet;
using static DevExpress.Data.Helpers.ExpressiveSortInfo;
using DocumentFormat.OpenXml.Bibliography;
using DevExpress.Xpo.DB.Helpers;
using Tesseract;
using Svg;
using System.Drawing.Imaging;
using DevExpress.XtraLayout.Customization;
using System.Collections;
using System.Web.UI.WebControls;
using DevExpress.Data.Filtering;
using DevExpress.Utils.Extensions;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Views.Printing;
using DevExpress.Utils.VisualEffects;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using GridView = DevExpress.XtraGrid.Views.Grid.GridView;
using static iText.IO.Image.Jpeg2000ImageData;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Threading.Tasks;
using Windows.Media.Ocr;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Runtime.Remoting.Contexts;
using AForge.Video.DirectShow;
using AForge.Video;
using static iText.IO.Codec.TiffWriter;
using SaovietTax.DTO;
using DocumentFormat.OpenXml.VariantTypes;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Net.Http;
using DevExpress.XtraWaitForm;
using DevExpress.XtraVerticalGrid;
using DataTable = System.Data.DataTable;
using DevExpress.Xpo;

namespace SaovietTax
{
    public partial class frmMain : DevExpress.XtraEditors.XtraForm
    {
        #region  Khai báo
        public string tknh { get; set; }
        public string savedPath = "";
        string dbPath = "";
        public int type { get; set; }
        public string username { get; set; }
        public string pasword { get; set; }
        public DateTime dtFrom { get; set; }
        public DateTime dtTo { get; set; }
        public static int Id = 1;
        bool isSetuppath = false;
        bool isetupDbpath = false;
        string password, connectionString;
        public string MasterMST = "3502264379";
        public System.Data.DataTable result { get; set; }
        private BindingList<FileImport> people = new BindingList<FileImport>();
        private BindingList<FileImport> people2 = new BindingList<FileImport>();
        private BindingList<FileImport> lstImportVao = new BindingList<FileImport>();
        private BindingList<FileImport> lstImportRa = new BindingList<FileImport>();
        System.Windows.Forms.BindingSource bindingSource = new System.Windows.Forms.BindingSource();
        System.Windows.Forms.BindingSource bindingSource2 = new System.Windows.Forms.BindingSource();
        public class FileImportDetail
        {
            public int ID { get; set; }
            public string Ten { get; set; }
            public int ParentId { get; set; }
            public string SoHieu { get; set; }
            public double Soluong { get; set; }
            public double Dongia { get; set; }
            public string DVT { get; set; }
            public string MaCT { get; set; }
            public string TKNo { get; set; }
            public string TKCo { get; set; }
            public double TTien { get; set; }
            public double TgTThue { get; set; }
          
            public FileImportDetail(string ten, int parentId, string soHieu, double soluong, double dongia, string dVT, string maCT, string tkNo, string tkCo,double ttien)
            {
                Ten = ten;
                ParentId = parentId;
                SoHieu = soHieu;
                Soluong = soluong;
                Dongia = dongia;
                DVT = dVT;
                MaCT = maCT;
                TKNo = tkNo;
                TKCo = tkCo;
                TTien = ttien;
            }
        }
        public class FileImport
        { 
            public string Path { get; set; }
            public bool Checked { get; set; }
            public int ID { get; set; }
            public string SHDon { get; set; }
            public string KHHDon { get; set; }
            public DateTime NLap { get; set; }
            public string Ten { get; set; }
            public string Noidung { get; set; }
            public string TKCo { get; set; }
            public string TKNo { get; set; }
            public int TkThue { get; set; }
            public string Mst { get; set; }
            public double TongTien { get; set; }
            public double TPhi { get; set; }
            public double TgTCThue { get; set; }
            public double TgTThue { get; set; }
            public int Vat { get; set; }
            public int Type { get; set; }
            public bool isAcess { get; set; }
            public bool isHaschild { get; set; }

            public string SoHieuTP { get; set; }
            public List<FileImportDetail> fileImportDetails;
            public FileImport(string path, string shdon, string khhdon, DateTime nlap, string ten, string noidung, string tkno, string tkco, int tkthue, string mst, double tongTien, int vat, int type, string tenTP,bool isacess, double tPhi, double tgTCThue, double tgTThue, bool _isHaschild)
            {
                ID = Id;
                SHDon = shdon;
                KHHDon = khhdon;
                NLap = nlap;
                Ten = ten;
                Noidung = noidung;
                TKCo = tkco;
                TKNo = tkno;
                TkThue = tkthue;
                Mst = mst;
                TongTien = tongTien;
                Vat = vat;
                Id += 1;
                fileImportDetails = new List<FileImportDetail>();
                Type = type;
                Checked = noidung.Contains("(*)") ? false : true;
                Path = path;
                SoHieuTP = tenTP;
                isAcess = isacess;
                TPhi = tPhi;
                TgTCThue = tgTCThue;
                TgTThue = tgTThue;
                isHaschild = _isHaschild;
            }

        }
        #endregion# 
        #region loadData
        public frmMain()
        {
            InitializeComponent(); 

        }
        
        private List<People> GetListFileImport()
        {

            return null;
        }
        public string tokken { get; set; } = "";
        private void LoadDataGridview(int type)
        {
            string mstcty = "";
            string query = "SELECT * FROM tbRegister";
            var kq2 = ExecuteQuery(query, null);
            if (string.IsNullOrEmpty(tokken))
            {
                 
                tokken = kq2.Rows[0]["tokken"].ToString();
                 
            }
            mstcty = kq2.Rows[0]["Username"].ToString();
            System.Data.DataTable tbDinhDanhtaikhoan = LoadDinhDanhTaiKhoan();
            System.Data.DataTable tbDinhDanhtaikhoanUuTien = LoadDinhDanhTaiKhoanUuTien(type);
            if (type == 1)
            {
                lstImportVao = new BindingList<FileImport>();
                // For individual GridControl 
                //Load data từ database 
                string queryCheckVatTu = @"SELECT * FROM tbimport   WHERE  Status <> 1 and Status <>2 AND Type = ?";

                var parameterss = new OleDbParameter[]
                { 
                 new OleDbParameter("?", type)
                };
                var kq = ExecuteQuery(queryCheckVatTu, parameterss);
                if (kq.Rows.Count > 0)
                {

                    var filteredRows = kq.AsEnumerable()
                       .Where(row => DateTime.TryParse(row["NLap"].ToString(), out DateTime nl)
                                     && nl.Date >= dtTungay.DateTime.Date && nl.Date <= dtDenngay.DateTime.Date);

                    // Kiểm tra xem có hàng nào sau khi lọc không
                    if (filteredRows.Any())
                    {
                        kq = filteredRows.CopyToDataTable(); 
                    }
                    else
                    {
                        // Xử lý trường hợp không có hàng nào
                        // Ví dụ: tạo DataTable rỗng hoặc thông báo cho người dùng
                        kq = kq.Clone(); // Tạo một DataTable rỗng với cấu trúc giống như kq
                    }
                }

                if (kq.Rows.Count > 0)
                {
                    lblSofiles.Text = kq.Rows.Count.ToString();
                    foreach (DataRow item in kq.Rows)
                    {
                        bool haschild = true;
                        if (item["IsHaschild"].ToString() == "1")
                        {
                            haschild = true;
                        }
                        else
                        {
                            haschild = false;

                        }
                            FileImport fileImport = new FileImport(item["Path"].ToString(), item["SHDon"].ToString(), item["KHHDon"].ToString(), DateTime.Parse(item["NLap"].ToString()), Helpers.ConvertVniToUnicode(item["Ten"].ToString()), Helpers.ConvertVniToUnicode(item["Noidung"].ToString()), item["TKNo"].ToString(), item["TKCo"].ToString(), int.Parse(item["TKThue"].ToString()), item["Mst"].ToString(), double.Parse(item["TongTien"].ToString()), int.Parse(item["Vat"].ToString()), int.Parse(item["Type"].ToString()), item["SohieuTP"].ToString(), true, double.Parse(item["TPHi"].ToString()), double.Parse(item["TgTCThue"].ToString()), double.Parse(item["TgTThue"].ToString()), haschild);
                        //add detail
                        fileImport.ID = int.Parse(item["ID"].ToString());
                        
                         queryCheckVatTu = @"SELECT * FROM tbimportdetail WHERE   ParentId= ?";
                         parameterss = new OleDbParameter[]
                         {
                            new OleDbParameter("?",int.Parse(item["ID"].ToString())) 
                         };
                          kq2 = ExecuteQuery(queryCheckVatTu, parameterss); 
                        if (kq2.Rows.Count > 0)
                        {
                            foreach (DataRow itemDetail in kq2.Rows)
                            {
                                FileImportDetail fileImportDetail = new FileImportDetail(Helpers.ConvertVniToUnicode(itemDetail["Ten"].ToString()), int.Parse(itemDetail["ParentId"].ToString()), itemDetail["SoHieu"].ToString(), double.Parse(itemDetail["SoLuong"].ToString(), CultureInfo.InvariantCulture), double.Parse(itemDetail["DonGia"].ToString()),Helpers.ConvertVniToUnicode(itemDetail["DVT"].ToString()), itemDetail["MaCT"].ToString(), itemDetail["TKNo"].ToString(), itemDetail["TKCo"].ToString(), double.Parse(itemDetail["TTien"].ToString()));
                                fileImportDetail.ID = int.Parse(itemDetail["ID"].ToString()); 
                                fileImport.fileImportDetails.Add(fileImportDetail);
                            }
                        }
                        lstImportVao.Add(fileImport); 
                    }
                }
                //load matdinhtaikhoan
                foreach (DataRow it in kq.Rows)
                {
                    FileImport item = lstImportVao.Where(m => m.ID == int.Parse(it["ID"].ToString())).FirstOrDefault();
                    if (item != null)
                    {
                        ApplyDefaultAndRuleBasedAccounts(item, tbDinhDanhtaikhoan, tbDinhDanhtaikhoanUuTien);
                        if (item.fileImportDetails.Count > 0 && string.IsNullOrEmpty(item.Noidung))
                        {
                            item.Noidung = item.fileImportDetails.FirstOrDefault().Ten;
                        }
                    }
                   
                }
                bindingSource.DataSource = lstImportVao;
                gridControl1.DataSource = bindingSource;
                //gridView1.Columns.Add(new DevExpress.XtraGrid.Columns.GridColumn
                //{
                //    Caption = "STT",
                //    FieldName = "Index",
                //    Visible = true,
                //    UnboundType = DevExpress.Data.UnboundColumnType.Integer,
                //    VisibleIndex = 0
                //});
                gridView1.OptionsDetail.EnableMasterViewMode = true;
                progressPanel1.Visible = false; // Ẩn progressPanel
            }
            else if (type == 2)
            {
                lstImportRa = new BindingList<FileImport>();
                string queryCheckVatTu = @"SELECT * FROM tbimport    WHERE  Status <> 1  AND Type = ?";

                var parameterss = new OleDbParameter[]
                {
                 new OleDbParameter("?", type)
                };
                var kq = ExecuteQuery(queryCheckVatTu, parameterss);
                if (kq.Rows.Count > 0)
                {
                    var filteredRows = kq.AsEnumerable()
                       .Where(row => DateTime.TryParse(row["NLap"].ToString(), out DateTime nl)
                                     && nl.Date >= dtTungay.DateTime.Date && nl.Date <= dtDenngay.DateTime.Date);
                    if (filteredRows.Any())
                    {
                        kq = filteredRows.CopyToDataTable();
                    }
                    else
                    {
                        // Xử lý trường hợp không có hàng nào
                        // Ví dụ: tạo DataTable rỗng hoặc thông báo cho người dùng
                        kq = kq.Clone(); // Tạo một DataTable rỗng với cấu trúc giống như kq
                    }
                }
                int i = 1;
                if (kq.Rows.Count > 0)
                {
                    lblSofiles2.Text = kq.Rows.Count.ToString();
                    foreach (DataRow item in kq.Rows)
                    {
                        bool haschild = true;
                        if (item["IsHaschild"].ToString() == "1")
                        {
                            haschild = true;
                        }
                        else
                        {
                            haschild = false;

                        }
                        FileImport fileImport = new FileImport(item["Path"].ToString(), item["SHDon"].ToString(), item["KHHDon"].ToString(), DateTime.Parse(item["NLap"].ToString()), Helpers.ConvertVniToUnicode(item["Ten"].ToString()), Helpers.ConvertVniToUnicode(item["Noidung"].ToString()), item["TKNo"].ToString(), item["TKCo"].ToString(), int.Parse(item["TKThue"].ToString()), item["Mst"].ToString(), double.Parse(item["TongTien"].ToString()), int.Parse(item["Vat"].ToString()), int.Parse(item["Type"].ToString()), item["SohieuTP"].ToString(), true, double.Parse(item["TPHi"].ToString()), double.Parse(item["TgTCThue"].ToString()), double.Parse(item["TgTThue"].ToString()), haschild);
                        //add detail
                         fileImport.ID = int.Parse(item["ID"].ToString());

                        progressPanel1.Caption = "Đang xử lý hóa đơn thứ " + i + "/ " + kq.Rows.Count.ToString();
                        Application.DoEvents();
                        queryCheckVatTu = @"SELECT * FROM tbimportdetail WHERE   ParentId= ?";
                        parameterss = new OleDbParameter[]
                        {
                            new OleDbParameter("?",int.Parse(item["ID"].ToString()))
                        };
                          kq2 = ExecuteQuery(queryCheckVatTu, parameterss);
                        i++;
                        if (kq2.Rows.Count > 0)
                        {
                            foreach (DataRow itemDetail in kq2.Rows)
                            {
                                FileImportDetail fileImportDetail = new FileImportDetail(Helpers.ConvertVniToUnicode(itemDetail["Ten"].ToString()), int.Parse(itemDetail["ParentId"].ToString()), itemDetail["SoHieu"].ToString(), double.Parse(itemDetail["SoLuong"].ToString()), double.Parse(itemDetail["DonGia"].ToString()), Helpers.ConvertVniToUnicode(itemDetail["DVT"].ToString()), itemDetail["MaCT"].ToString(), itemDetail["TKNo"].ToString(), itemDetail["TKCo"].ToString(), double.Parse(itemDetail["TTien"].ToString()));
                                fileImportDetail.ID = int.Parse(itemDetail["ID"].ToString());
                                fileImport.fileImportDetails.Add(fileImportDetail);
                            }
                        }
                        else
                        {
                            GetdetailXML2(mstcty, item["KHHDon"].ToString(), item["SHDon"].ToString(), tokken, int.Parse(item["InvoiceType"].ToString()));
                            var parame = new OleDbParameter[]
                        {
                            new OleDbParameter("?",int.Parse(item["ID"].ToString()))
                        };
                            var kq3 = ExecuteQuery(queryCheckVatTu, parame);
                            foreach (DataRow itemDetail in kq3.Rows)
                            {
                                FileImportDetail fileImportDetail = new FileImportDetail(Helpers.ConvertVniToUnicode(itemDetail["Ten"].ToString()), int.Parse(itemDetail["ParentId"].ToString()), itemDetail["SoHieu"].ToString(), double.Parse(itemDetail["SoLuong"].ToString()), double.Parse(itemDetail["DonGia"].ToString()), Helpers.ConvertVniToUnicode(itemDetail["DVT"].ToString()), itemDetail["MaCT"].ToString(), itemDetail["TKNo"].ToString(), itemDetail["TKCo"].ToString(), double.Parse(itemDetail["TTien"].ToString()));
                                fileImportDetail.ID = int.Parse(itemDetail["ID"].ToString());
                                fileImport.fileImportDetails.Add(fileImportDetail);
                            }
                        }
                            lstImportRa.Add(fileImport);
                    }
                }
                bindingSource.DataSource = lstImportRa;
                gridControl2.DataSource = bindingSource;
                gridControl2.RefreshDataSource();
                gridView3.OptionsDetail.EnableMasterViewMode = true;
                progressPanel1.Visible = false; // Ẩn progressPanel
            }

            SetGridViewOptions(gridView1);
            SetGridViewOptions(gridView3);
        }

        private void SetGridViewOptions(DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            view.OptionsSelection.MultiSelect = true;
            view.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect;
            GridStripRow(view);
        }


        public void GridStripRow(DevExpress.XtraGrid.Views.Grid.GridView gridView)
        {
            if (gridView != null)
            {
                // Kích hoạt kiểu dáng hàng chẵn và lẻ
                gridView.OptionsView.EnableAppearanceEvenRow = true;
                gridView.OptionsView.EnableAppearanceOddRow = true;

                // Thiết lập màu sắc cho hàng chẵn
                gridView.Appearance.EvenRow.BackColor = System.Drawing.Color.FromArgb(168, 255, 253);

                gridView.Appearance.EvenRow.ForeColor = System.Drawing.Color.Black; // Màu chữ cho hàng chẵn

                // Thiết lập màu sắc cho hàng lẻ
                gridView.Appearance.OddRow.BackColor = System.Drawing.Color.White; // Màu nền cho hàng lẻ
                gridView.Appearance.OddRow.ForeColor = System.Drawing.Color.Black; // Màu chữ cho hàng lẻ

                gridView.CellValueChanged += GridView_CellValueChanged;

            }
        }

        bool getMessage = true;
        private void GridView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            // Lấy thông tin về hàng và cột của ô đã thay đổi
            int rowHandle = e.RowHandle;
            string columnName = e.Column.FieldName; // Tên cột
            if (columnName != "TKNo" && columnName != "TKCo")
                return;
            object newValue = e.Value; // Giá trị mới
            
            string query = "SELECT * FROM HeThongTK WHERE SoHieu = ?";
            if (newValue == null)
                return;
            if (!string.IsNullOrEmpty(newValue.ToString()))
            {
                // Tạo mảng tham số với giá trị cho câu lệnh SQL
                OleDbParameter[] parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", newValue),
                };
                var kq = ExecuteQuery(query, parameters);
                if (kq.Rows.Count == 0 && !newValue.ToString().Contains("|") && getMessage)
                {
                    getMessage = false;
                    DevExpress.XtraGrid.Views.Grid.GridView gridView = gridControl1.MainView as DevExpress.XtraGrid.Views.Grid.GridView;
                    //gridView.SetRowCellValue(rowHandle, e.Column, "");
                    XtraMessageBox.Show("Số tài khoản không tồn tại trong hệ thống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error); 

                }
            }
        }

        static bool TableExists(OleDbConnection connection, string tableName)
        {
            try
            {
                // Kiểm tra sự tồn tại của bảng
                System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow row in schemaTable.Rows)
                {
                    if (row["TABLE_NAME"].ToString().Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi kiểm tra bảng: {ex.Message}");
            }
            return false;
        }
        static void CreateTableNganhang(OleDbConnection connection, string tableName)
        {
            string createTableQuery = $@"
        CREATE TABLE {tableName} (
            ID AUTOINCREMENT PRIMARY KEY,
            SHDon TEXT,
            NgayGD TEXT,
            DienGiai TEXT,
            TongTien NUMBER,
            TKNo TEXT,  
            TKCo TEXT 
        );";

            using (OleDbCommand command = new OleDbCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
        static void CreateTableDinhDanh(OleDbConnection connection, string tableName)
        {
            string createTableQuery = $@"
        CREATE TABLE {tableName} (
            ID AUTOINCREMENT PRIMARY KEY,
            Type TEXT,
            KeyValue TEXT,
            TKNo TEXT,  
            TKCo TEXT,
            TKThue TEXT,
            Uutien TEXT
        );";

            using (OleDbCommand command = new OleDbCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
        static void CreateTableDinhDanhNganhang(OleDbConnection connection, string tableName)
        {
            string createTableQuery = $@"
        CREATE TABLE {tableName} (
            ID AUTOINCREMENT PRIMARY KEY,  
            Noidung TEXT,
            TK TEXT 
        );";

            using (OleDbCommand command = new OleDbCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
        static void CreateTable(OleDbConnection connection, string tableName)
        {
            string createTableQuery = $@"
        CREATE TABLE {tableName} (
            ID AUTOINCREMENT PRIMARY KEY,
            SHDon TEXT,
            KHHDon TEXT,
            NLap TEXT,
            Ten TEXT,
            Noidung TEXT,
            TKCo TEXT,
            TKNo TEXT,
            TkThue TEXT,
            Mst TEXT,
            Status NUMBER,
            Ngaytao TEXT,
            TongTien NUMBER,
            Vat NUMBER,
            SohieuTP TEXT
        );";

            using (OleDbCommand command = new OleDbCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
        static void CreateTableDetail(OleDbConnection connection, string tableName)
        {
            string createTableQuery = $@"
        CREATE TABLE {tableName} (
            ID AUTOINCREMENT PRIMARY KEY,
            ParentId TEXT,
            SoHieu TEXT,
            SoLuong TEXT,
            DonGia TEXT,
            DVT TEXT,
            Ten TEXT ,
            MaCT TEXT,
            TKNo TEXT,
            TKCo TEXT
        );";

            using (OleDbCommand command = new OleDbCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
        string mstcongty = "";
        private void InitDB()
        {
            string appPath = Assembly.GetExecutingAssembly().Location;

            // Lấy thư mục chứa ứng dụng
            string directoryPath = Path.GetDirectoryName(appPath);

            // Xóa phần \bin\Debug để lấy đường dẫn gốc
            string rootDirectory = Path.GetFullPath(Path.Combine(directoryPath, @"..\.."));

            // Tạo đường dẫn đến file dpPath.txt trong thư mục hoadon
            string filePaths = Path.Combine(rootDirectory, "hoadon", "dpPath.txt");
            try
            {
                string content = File.ReadAllText(filePaths);
                dbPath = content;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi đọc file: " + ex.Message);
            }

            
            // Đọc toàn bộ nội dung tệp
            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            //connectionString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            // connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database";
            //connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\S.T.E 25\S.T.E 25\DATA\importData.accdb;Persist Security Info=False";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = "SELECT * FROM License";

                    // Tạo mảng tham số với giá trị cho câu lệnh SQL

                    var kq = ExecuteQuery(query, null);
                    if (kq.Rows.Count > 0)
                    {
                        string tencongty = kq.Rows[0]["TenCty"].ToString();
                        string fileName = Path.GetFileName(dbPath.Trim());
                        mstcongty = kq.Rows[0]["MaSoThue"].ToString();
                        lblDpPath.Text = Helpers.ConvertVniToUnicode(tencongty) +"|"+ mstcongty + "|"+ fileName+" | "+"Version 3.62";
                    }
                   
                }
                catch (Exception ex)
                {
                    throw ex;
                }


                string tableName = "tbimport";
                string tableNamedetail = "tbimportdetail";
                string tableDinhdanh = "tbDinhdanhtaikhoan";
                string tableDinhdanhNganhang = "tbDinhdanhNganhang";
                string tableNganhang = "tbNganhang";
                // Kiểm tra xem bảng đã tồn tại hay không
                if (!TableExists(connection, tableNganhang))
                {
                    // Tạo bảng nếu chưa tồn tại
                    CreateTableNganhang(connection, tableNganhang);
                    Console.WriteLine($"Bảng '{tableNganhang}' đã được tạo thành công.");
                }
                if (!TableExists(connection, tableDinhdanhNganhang))
                {
                    // Tạo bảng nếu chưa tồn tại
                    CreateTableDinhDanhNganhang(connection, tableDinhdanhNganhang);
                    Console.WriteLine($"Bảng '{tableDinhdanhNganhang}' đã được tạo thành công.");
                }
                if (!TableExists(connection, tableDinhdanh))
                {
                    // Tạo bảng nếu chưa tồn tại
                    CreateTableDinhDanh(connection, tableDinhdanh);
                    Console.WriteLine($"Bảng '{tableDinhdanh}' đã được tạo thành công.");
                }
                else
                {
                    if (!ColumnExists(connection, "tbNganhang", "Status"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbNganhang", "Status", "NUMBER"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    if (!ColumnExists(connection, "tbNganhang", "TongTien2"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbNganhang", "TongTien2", "NUMBER"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    // Kiểm tra xem cột tkoco đã tồn tại hay chưa
                    if (!ColumnExists(connection, "tbRegister", "tokken"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbRegister", "tokken", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    // Kiểm tra xem cột tkoco đã tồn tại hay chưa
                    if (!ColumnExists(connection, "tbRegister", "col1"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbRegister", "col1", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    if (!ColumnExists(connection, "tbRegister", "col2"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbRegister", "col2", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    if (!ColumnExists(connection, "tbimport", "InvoiceType"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbimport", "InvoiceType", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    if (!ColumnExists(connection, "tbimport", "IsHaschild"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbimport", "IsHaschild", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    // Kiểm tra xem cột tkoco đã tồn tại hay chưa
                    if (!ColumnExists(connection, "tbimport", "Path"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbimport", "Path", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    if (!ColumnExists(connection, "tbimport", "Type"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbimport", "Type", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    if (!ColumnExists(connection, "tbimport", "TPhi"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbimport", "TPhi", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    if (!ColumnExists(connection, "tbimport", "TgTCThue"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbimport", "TgTCThue", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                    if (!ColumnExists(connection, "tbimport", "TgTThue"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbimport", "TgTThue", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần 
                    }
                }
                if (!TableExists(connection, tableName))
                {
                    // Tạo bảng nếu chưa tồn tại
                    CreateTable(connection, tableName);
                    Console.WriteLine($"Bảng '{tableName}' đã được tạo thành công.");
                }
                else
                {
                    Console.WriteLine($"Bảng '{tableName}' đã tồn tại.");
                }
                // Kiểm tra xem bảng đã tồn tại hay không
                if (!TableExists(connection, tableNamedetail))
                {
                    // Tạo bảng nếu chưa tồn tại
                    CreateTableDetail(connection, tableNamedetail);
                    Console.WriteLine($"Bảng '{tableNamedetail}' đã được tạo thành công.");
                }
                else
                {
                    // Kiểm tra xem cột tkoco đã tồn tại hay chưa
                    if (!ColumnExists(connection, "tbimportdetail", "TKCo"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbimportdetail", "TKCo", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần
                        Console.WriteLine("Cột 'tkoco' đã được thêm vào bảng 'tbimportdetail'.");
                    }
                    else
                    {
                        Console.WriteLine("Cột 'tkoco' đã tồn tại trong bảng 'tbimportdetail'.");
                    }
                    //
                    if (!ColumnExists(connection, "tbimportdetail", "TTien"))
                    {
                        // Nếu không tồn tại, thêm cột tkoco
                        AddColumn(connection, "tbimportdetail", "TTien", "TEXT"); // Bạn có thể thay đổi kiểu dữ liệu nếu cần
                        Console.WriteLine("Cột 'TTien' đã được thêm vào bảng 'tbimportdetail'.");
                    }
                    else
                    {
                        Console.WriteLine("Cột 'TTien' đã tồn tại trong bảng 'tbimportdetail'.");
                    }

                    Console.WriteLine($"Bảng '{tableNamedetail}' đã tồn tại.");
                }
            }

        }
        static void AddColumn(OleDbConnection connection, string tableName, string columnName, string dataType)
        {
            string sql = $"ALTER TABLE [{tableName}] ADD COLUMN [{columnName}] {dataType};";
            using (OleDbCommand command = new OleDbCommand(sql, connection))
            {
                command.ExecuteNonQuery();
            }
        }
        static bool ColumnExists(OleDbConnection connection, string tableName, string columnName)
        {
            using (OleDbCommand command = new OleDbCommand($"SELECT TOP 1 * FROM [{tableName}]", connection))
            {
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if (reader.GetName(i).Equals(columnName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        private void ControlsSetup()
        {

            //Thiết lập cho cbb Từ ngày, đến ngày

            comboBoxEdit1.Properties.Buttons[0].Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Combo;

            comboBoxEdit1.Properties.Items.AddRange(new string[]
         {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12"
         });

            comboBoxEdit1.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;

            comboBoxEdit2.Properties.Items.AddRange(new string[]
        {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12"
        });

            comboBoxEdit1.SelectedIndex = 0;
            comboBoxEdit2.SelectedItem = DateTime.Now.Month.ToString();

            comboBoxEdit2.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            // progressBarControl1.EditValue = 10;
        }
        private void InitData()
        {
            savedPath = ConfigurationManager.AppSettings["LastFilePath"];
            string query = "SELECT * FROM tbRegister";

            // Tạo mảng tham số với giá trị cho câu lệnh SQL

            var kq = ExecuteQuery(query, null);
            if (kq.Rows.Count > 0)
            {
                savedPath = kq.Rows[0]["Hoadonpath"].ToString();
                txtuser.Text = kq.Rows[0]["Username"].ToString();
                txtpass.Text = kq.Rows[0]["Password"].ToString();
            }
          
        }
        private BackgroundWorker worker;
        private static int GetColumnLength(OleDbConnection connection, string tableName, string columnName)
        {
            int length = 0;

            string sql = $"SELECT TOP 1 [{columnName}] FROM [{tableName}]";
            using (OleDbCommand command = new OleDbCommand(sql, connection))
            {
                using (OleDbDataReader reader = command.ExecuteReader(CommandBehavior.SchemaOnly))
                {
                    System.Data.DataTable schemaTable = reader.GetSchemaTable();
                    if (schemaTable != null)
                    {
                        foreach (DataRow row in schemaTable.Rows)
                        {
                            if (row["ColumnName"].ToString() == columnName)
                            {
                                length = Convert.ToInt32(row["ColumnSize"]);
                                break;
                            }
                        }
                    }
                }
            }

            return length;
        }
        private void CheckDB()
        {
            OleDbConnection connection = new OleDbConnection(connectionString);
            if (connectionString == null)
                return;
            connection.Open();
            int checkLenghtTen = GetColumnLength(connection, "KhachHang", "Ten");
            if (checkLenghtTen < 255)
            {
                // Lệnh SQL để thay đổi kích thước cột Ten từ 100 sang 255
                string sql = "ALTER TABLE KhachHang ALTER COLUMN Ten TEXT(255)";

                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    // Thực thi lệnh SQL
                    command.ExecuteNonQuery();
                    Console.WriteLine("Kích thước cột Ten đã được thay đổi thành 255.");

                }
            }
            int checkLenghtDiachi = GetColumnLength(connection, "KhachHang", "DiaChi");
            if (checkLenghtDiachi < 255)
            {
                // Lệnh SQL để thay đổi kích thước cột Ten từ 100 sang 255
                string sql = "ALTER TABLE KhachHang ALTER COLUMN DiaChi TEXT(255)";

                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    // Thực thi lệnh SQL
                    command.ExecuteNonQuery();
                    Console.WriteLine("Kích thước cột Ten đã được thay đổi thành 255.");
                }
            }
        }
        private void SetVietnameseCulture()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("vi-VN");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("vi-VN");
           // var files = Directory.EnumerateFiles(savedPath + @"\HDVao", "*.xml", SearchOption.AllDirectories).ToList();

            dtTungay.DateTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            dtDenngay.DateTime = DateTime.Now;

            progressPanel1.Caption = "Đang xử lý...";
            progressPanel1.Description = "Vui lòng chờ...";
        }
      
        public class SvgConverter
        {
            public void ConvertBase64ToSvg(string base64Data, string outputPath)
            {
                // Tách phần đầu để lấy dữ liệu Base64
                var base64 = base64Data.Substring(base64Data.IndexOf(",") + 1);

                // Giải mã dữ liệu Base64
                byte[] svgBytes = Convert.FromBase64String(base64);

                // Lưu vào tệp SVG
                File.WriteAllBytes(outputPath, svgBytes);
            }
        }

        private void Testimg2(string base64data)
        {
            string base64Data = base64data;
            string outputPath = AppDomain.CurrentDomain.BaseDirectory + "output.svg";

            SvgConverter converter = new SvgConverter();
            converter.ConvertBase64ToSvg(base64Data, outputPath);
                
            Console.WriteLine("Tệp SVG đã được lưu tại: " + outputPath); 
            RunMain();
            var readcapcha = Readcapcha();
        }
        private void RunMain()
        {
            string exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "main.exe");

            try
            {
                // Kiểm tra xem tệp có tồn tại không
                if (!File.Exists(exePath))
                {
                    Console.WriteLine("Tệp main.exe không tồn tại.");
                    return;
                }

                // Tạo một Process để chạy tệp .exe
                Process process = new Process();
                process.StartInfo.FileName = exePath;
                process.StartInfo.UseShellExecute = false; // Không sử dụng shell để chạy
                process.StartInfo.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory; // Đặt thư mục làm việc

                process.Start(); // Bắt đầu tiến trình
                Thread.Sleep(2000); // Đợi 2 giây 

                // Đóng tiến trình
                if (!process.HasExited)
                {
                    process.Kill();
                }
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("Tệp không tìm thấy: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show("Không có quyền truy cập: " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khác xảy ra: " + ex.Message);
            }
        }
        private string Readcapcha()
        
        {
            string filePath = AppDomain.CurrentDomain.BaseDirectory+ "captcha.txt"; // Đảm bảo tệp ở cùng thư mục với chương trình
          
            try
            {
                // Đọc nội dung từ tệp
                string content = File.ReadAllText(filePath);
                Console.WriteLine("Nội dung của captcha.txt:");
                Console.WriteLine(content); 
                return content; // Trả về nội dung đã đọc
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("Tệp không tồn tại.");
                return null; // Hoặc trả về một giá trị mặc định nếu tệp không tồn tại
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message);
                return null; // Hoặc trả về một giá trị mặc định
            }
        }

        private Bitmap ProcessImageForOCR(Bitmap original)
        {
            // Tăng kích thước ảnh để cải thiện nhận diện
            var newWidth = original.Width * 2;
            var newHeight = original.Height * 2;

            Bitmap resized = new Bitmap(newWidth, newHeight);
            using (Graphics g = Graphics.FromImage(resized))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(original, 0, 0, newWidth, newHeight);
            }

            // Chuyển sang ảnh đen trắng
            Bitmap bw = new Bitmap(resized.Width, resized.Height);
            using (Graphics g = Graphics.FromImage(bw))
            {
                ColorMatrix cm = new ColorMatrix(new float[][] {
            new float[] {0.3f, 0.3f, 0.3f, 0, 0},
            new float[] {0.59f, 0.59f, 0.59f, 0, 0},
            new float[] {0.11f, 0.11f, 0.11f, 0, 0},
            new float[] {0, 0, 0, 1, 0},
            new float[] {0, 0, 0, 0, 1}
        });

                ImageAttributes ia = new ImageAttributes();
                ia.SetColorMatrix(cm);

                g.DrawImage(resized,
                            new Rectangle(0, 0, resized.Width, resized.Height),
                            0, 0, resized.Width, resized.Height,
                            GraphicsUnit.Pixel, ia);
            }

            // Tăng độ tương phản
            Bitmap highContrast = new Bitmap(bw.Width, bw.Height);
            float contrast = 1.5f; // Tăng độ tương phản
            float adjusted = -(0.5f * contrast) + 0.5f;

            using (Graphics g = Graphics.FromImage(highContrast))
            {
                ImageAttributes ia = new ImageAttributes();
                ia.SetColorMatrix(new ColorMatrix(new float[][] {
            new float[] {contrast, 0, 0, 0, 0},
            new float[] {0, contrast, 0, 0, 0},
            new float[] {0, 0, contrast, 0, 0},
            new float[] {0, 0, 0, 1, 0},
            new float[] {adjusted, adjusted, adjusted, 0, 1}
        }));

                g.DrawImage(bw,
                            new Rectangle(0, 0, bw.Width, bw.Height),
                            0, 0, bw.Width, bw.Height,
                            GraphicsUnit.Pixel, ia);
            }

            return highContrast;
        }
        static Bitmap PreprocessImage(Bitmap original)
        {
            // Tạo ảnh mới với cùng kích thước
            var newImage = new Bitmap(original.Width, original.Height);

            // Chuyển sang grayscale và tăng độ tương phản
            for (int x = 0; x < original.Width; x++)
            {
                for (int y = 0; y < original.Height; y++)
                {
                    System.Drawing.Color pixel = original.GetPixel(x, y);
                    int grayValue = (int)(pixel.R * 0.3 + pixel.G * 0.59 + pixel.B * 0.11);
                    newImage.SetPixel(x, y, grayValue < 128 ? System.Drawing.Color.Black : System.Drawing.Color.White);
                }
            }

            return newImage;
        }
        private string MSTCongTY = "";
        private void GetMST()
        {
            string query = "SELECT * FROM License";

            // Tạo mảng tham số với giá trị cho câu lệnh SQL

            var kq = ExecuteQuery(query, null);
            if (kq.Rows.Count > 0)
            {
                MSTCongTY = kq.Rows[0]["MaSoThue"].ToString();
            }
        }
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            // Thực hiện công việc khởi tạo ở đây
        }
       
        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressPanel2.Visible = false; // Ẩn ProgressPanel khi hoàn thành
        }
        public List<VatTu> lstvt = new List<VatTu>();
        private async void frmMain_Load(object sender, EventArgs e)
        {
           
            InitDB();
            InitData(); 
            SetVietnameseCulture();
            GetMST();
            string fileName = Path.GetFileName(dbPath.Trim());
            CheckDB();
            ControlsSetup();
            lstvt = await  LoadDataVattuAsync();
            var query = "SELECT * FROM KhachHang"; // Giả sử bạn muốn lấy tất cả dữ liệu từ bảng KhachHang
            tbKhachhang = ExecuteQuery(query);
        }
        #endregion
        #region Xử lý xml
        int progressPercentage = 0;
        int filesLoaded = 1;
        int totalCount = 0;
        private static bool IsFileInMonthRange(string filePath, string baseDirectory, int fromMonth, int toMonth)
        {
            // Lấy tên thư mục từ đường dẫn file
            string directoryName = System.IO.Path.GetDirectoryName(filePath)?.Split(System.IO.Path.DirectorySeparatorChar).Last();

            // Kiểm tra xem tên thư mục có thuộc khoảng tháng không
            if (int.TryParse(directoryName, out int month))
            {
                return month >= fromMonth && month <= toMonth;
            }

            return false; // Không phải thư mục tháng hợp lệ
        }
        private static bool IsFileInDateRange(string filePath, string baseDirectory, DateTime fromDate, DateTime toDate)
{
    // Lấy tên thư mục từ đường dẫn file
    string directoryName = System.IO.Path.GetDirectoryName(filePath)?.Split(System.IO.Path.DirectorySeparatorChar).Last();

    // Kiểm tra xem tên thư mục có thể chuyển đổi thành ngày không
    if (DateTime.TryParse(directoryName, out DateTime directoryDate))
    {
        return directoryDate >= fromDate && directoryDate <= toDate;
    }

    return false; // Không phải thư mục ngày hợp lệ
}
        private void LoadXmlFiles(string path,int type)
        {
            progressPanel1.Visible = true;
            lblThongbao.Text = "Bắt đầu chạy";
            progressPanel1.Caption = "Bắt đầu chạy...";
            if (type == 1)
                people = new BindingList<FileImport>();
            if (type == 2)
                people2 = new BindingList<FileImport>();
            BindingList<FileImport> peopleTemp = new BindingList<FileImport>();
            progressBarControl1.EditValue = 0;
            Application.DoEvents(); 
            path += (type == 1 ? "\\HDVao" : "\\HDRa");

            int fromMonth = int.Parse(dtTungay.DateTime.Month.ToString()); // Thay đổi theo tháng bắt đầu (ví dụ: 3 cho tháng 3)
            int toMonth = int.Parse(dtDenngay.DateTime.Month.ToString());   // Thay đổi theo tháng kết thúc (ví dụ: 7 cho tháng 7)
            // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
            var files = Directory.EnumerateFiles(path, "*.xml", SearchOption.AllDirectories)
                .Where(file => IsFileInMonthRange(file, path, dtTungay.DateTime.Month, dtDenngay.DateTime.Month)).ToList();
            //Lọc thêm điều kiện theo ngày
            lblThongbao.Text = "Đếm file xml";
            progressPanel1.Caption = "Đếm file xml...";
            List<string> lstdelete = new List<string>();
            foreach (var item in files)
            {
                XmlDocument xmlDoc = new XmlDocument();
                string fullPath = item;
                using (StreamReader reader = new StreamReader(fullPath, Encoding.UTF8))
                {
                    try
                    {
                        xmlDoc.Load(reader); // Tải file XML
                    }
                    catch (XmlException ex)
                    {
                        Console.WriteLine($"Lỗi khi tải file XML: {ex.Message}");
                        return;
                    }
                    XmlNode root = xmlDoc.DocumentElement;
                    XmlNode nTTChungNode = root.SelectSingleNode("//TTChung");
                    DateTime NLap = DateTime.Parse(nTTChungNode.SelectSingleNode("NLap")?.InnerText);
                    if (NLap >= dtTungay.DateTime && NLap <= dtDenngay.DateTime)
                        continue;
                    else
                        lstdelete.Add(item);

                }
            }
            foreach(var item in lstdelete)
            {
                files.Remove(item);
            }

            int countXml = files.Count();
            lblThongbao.Text = "Đếm file excel";
            progressPanel1.Caption = "Đếm file excel...";
            Dictionary<string, string> lstHodpn = new Dictionary<string, string>();
            //Lấy danh sách hóa đơn để kiểm tra cho excel

            // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
            var excelFiles = Directory.EnumerateFiles(path, "*.xlsx", SearchOption.AllDirectories)
                .Where(file => IsFileInMonthRange(file, path, dtTungay.DateTime.Month, dtDenngay.DateTime.Month)).ToList(); // Kiểm tra xem file có nằm trong khoảng tháng
            int rowCount = 0;

            //Kiểm tra xem có bao nhieu dòng dữ liệu trong Excel
            for (int j = 0; j < excelFiles.Count; j++)
            {
                int demdong = 0;
                using (var workbook = new XLWorkbook(excelFiles[j]))
                {
                    // Lấy worksheet đầu tiên
                    var worksheet = workbook.Worksheet(1); // Hoặc bạn có thể dùng tên worksheet như worksheet = workbook.Worksheet("Sheet1");
                                                           // Lấy giá trị của ô A6

                    var currentCell = worksheet.Cell("A7"); // Bắt đầu từ ô A6

                    // Kiểm tra các ô bắt đầu từ A6 cho đến khi gặp ô trống
                    while (!currentCell.IsEmpty())
                    {
                        var getid = worksheet.Cell(demdong + 7, 1).Value.ToString().Trim();
                        DateTime getNgayLap= DateTime.Parse(worksheet.Cell(demdong + 7, 5).Value.ToString().Trim());
                        int id = 0;
                        if (int.TryParse(getid, out id))
                        {
                            if(getNgayLap>=dtTungay.DateTime && getNgayLap <= dtDenngay.DateTime)
                            {
                                rowCount++; // Tăng số dòng
                            }

                            demdong++;
                        }
                        currentCell = currentCell.Worksheet.Row(currentCell.Address.RowNumber + 1).Cell("A"); // Chuyển xuống ô bên dưới

                    }

                }
            }

            int countExcel = 0;
            if (rowCount > 0)
                countExcel = rowCount;

            totalCount = countXml + countExcel; 
            if (type == 1)
                lblSofiles.Text = totalCount.ToString();
            if (type == 2)
                lblSofiles2.Text = totalCount.ToString();
            //foreach (string file in files)
            //{
            //    progressPercentage = (filesLoaded * 100) / totalCount;
            //    filesLoaded += 1;
            //    progressBarControl1.EditValue = progressPercentage;
            //} 
            foreach (string file in files)
            {
                lblThongbao.Text = "Đọc file thứ "+(files.IndexOf(file) + 1);
                //progressPanel1.Caption = "Đọc file thứ " + (files.IndexOf(file) + 1) +"/ "+ totalCount;
                progressPanel1.Caption = "Đang load files...";
                progressPercentage = (filesLoaded * 100) / totalCount;
                filesLoaded += 1;
                progressBarControl1.EditValue = progressPercentage;
                Application.DoEvents();

                // Lấy tên tệp từ đường dẫn
                string fileName = System.IO.Path.GetFileName(file);
                //people.Add(new FileImport(file,10,"asdsa"));

                //Đọc từ XML
                XmlDocument xmlDoc = new XmlDocument();
                string fullPath = file;
                using (StreamReader reader = new StreamReader(fullPath, Encoding.UTF8))
                {
                    try
                    {
                        xmlDoc.Load(reader); // Tải file XML
                    }
                    catch (XmlException ex)
                    {
                        Console.WriteLine($"Lỗi khi tải file XML: {ex.Message}");
                        return;
                    }
                }

                // Lấy phần tử gốc
                XmlNode root = xmlDoc.DocumentElement;

                // Lấy phần tử <NDHDon>
                XmlNode ndhDonNode = root.SelectSingleNode("//NDHDon");
                XmlNode nTTChungNode = root.SelectSingleNode("//TTChung");
                XmlNode nTTThanToan = root.SelectSingleNode("//LTSuat");
                var nThTien = root.SelectNodes("//LTSuat//ThTien");
                var nTSuat = root.SelectNodes("//LTSuat//TSuat");
                var nTPhi= root.SelectSingleNode("//TToan//TPhi");
                var nTgTCThue = root.SelectSingleNode("//TToan//TgTCThue")!=null?  root.SelectSingleNode("//TToan//TgTCThue").InnerText.Replace('.', ','):"";
                var nTgTThue = root.SelectSingleNode("//TToan//TgTThue")!=null? root.SelectSingleNode("//TToan//TgTThue").InnerText.Replace('.', ','):"";
                bool isAcess = true;
                if(type==1)
                {
                    string getmst =  root.SelectSingleNode("//NMua//MST").InnerText;
                    isAcess =mstcongty == getmst ? true : false;
                }
                if (type == 2)
                {
                    string getmst = root.SelectSingleNode("//NBan//MST").InnerText;
                    isAcess = mstcongty == getmst ? true : false;
                }
                string SHDon = "";
                string KHHDon = "";
                string ten = "";
                string mst = "";
                string ten2 = "";
                string mst2 = "";
                string SoHD = "";
                int TkCo = 0;
                int TkNo = 0;
                int TkThue = 0;
                int Vat = 0;
                double Thanhtien = 0;
                double TPhi = 0;
                double TgTCThue = 0;
                double TgTThue = 0;
                string diengiai = "";
                DateTime NLap = new DateTime();
                if (nTTChungNode != null)
                {
                    SHDon = nTTChungNode.SelectSingleNode("SHDon")?.InnerText;
                    KHHDon = nTTChungNode.SelectSingleNode("KHHDon")?.InnerText;
                    NLap = DateTime.Parse(nTTChungNode.SelectSingleNode("NLap")?.InnerText);
                }
                //Kiểm tra trong database có hoa do nay chưa
                string query = "SELECT * FROM HoaDon WHERE KyHieu = ? AND SoHD LIKE ?";

                // Tạo mảng tham số với giá trị cho câu lệnh SQL
                OleDbParameter[] parameters = new OleDbParameter[]
                {
            new OleDbParameter("KyHieu", KHHDon),          // Sử dụng chỉ số mà không cần tên
            new OleDbParameter("SoHD", "%" + SHDon + "%")  // Thêm ký tự % cho LIKE
                };
                lblThongbao.Text = "Kiểm tra hóa đơn ";
                Application.DoEvents();
                var kq = ExecuteQuery(query, parameters);
                if (kq.Rows.Count > 0)
                {
                    continue;
                }
                if (people.Any(m => m.SHDon.Contains(SHDon) && m.KHHDon == KHHDon))
                {
                    continue;
                }
                lblThongbao.Text = "Kiểm tra table import ";
                Application.DoEvents();
                query = "SELECT * FROM tbimport WHERE KHHDon = ? AND SHDon LIKE ?";
                parameters = new OleDbParameter[]
              {
            new OleDbParameter("KHHDon", KHHDon),          // Sử dụng chỉ số mà không cần tên
            new OleDbParameter("SHDon", "%" + SHDon + "%")  // Thêm ký tự % cho LIKE
              };
                kq = ExecuteQuery(query, parameters);
                if (kq.Rows.Count > 0)
                {
                    continue;
                }
                XmlNode nBanNode=null;
                if (type == 1)
                    nBanNode = ndhDonNode.SelectSingleNode("NBan");
                if (type == 2)
                    nBanNode = ndhDonNode.SelectSingleNode("NMua");
                //  XmlNode nMuaNode = ndhDonNode.SelectSingleNode("NMua");
                if (nBanNode != null)
                {
                    ten = nBanNode.SelectSingleNode("Ten")?.InnerText;
                    if (string.IsNullOrEmpty(ten))
                    {
                        ten = nBanNode.SelectSingleNode("HVTNMHang")?.InnerText;
                    }
                    mst = nBanNode.SelectSingleNode("MST")?.InnerText;
                   if(string.IsNullOrEmpty(mst))
                    {
                        if (root.SelectSingleNode("//NMua//DChi")!= null)
                        {
                            string getdc = root.SelectSingleNode("//NMua//DChi").InnerText;
                            if (string.IsNullOrEmpty(getdc))
                                InitCustomer(type == 1 ? 2 : 3, "", ten, "", mst);
                            else
                            {
                                InitCustomer(type == 1 ? 2 : 3, "", ten, getdc, mst);
                            }
                        }
                        else
                        {
                            InitCustomer(type == 1 ? 2 : 3, "", ten, "", mst);
                        }
                       
                    }
                }

                if (nTSuat != null)
                {
                    for (int i = 0; i < nTSuat.Count; i++)
                    {
                        XmlNode item = nTSuat[i];
                        if (item.InnerText != "KKKNT" && item.InnerText != "KCT")
                        {
                            if (Vat==null || Vat == 0)
                                Vat = int.Parse(item.InnerText.Replace("%", ""));
                        }
                        else
                            Vat = 0;
                    }

                }
                else
                {
                    Vat = 0;
                }
                if (nThTien != null)
                {
                    if (nThTien.Count > 0)
                    {
                        for (int i = 0; i < nThTien.Count; i++)
                        {
                            if (nThTien[i].InnerText != "0")
                                Thanhtien = double.Parse(nThTien[i].InnerText.Replace('.', ',')); 
                        }
                    }
                    else
                    {
                        XmlNode TgTTTBSo = root.SelectSingleNode("//TToan//TgTTTBSo");
                        Thanhtien = double.Parse(TgTTTBSo.InnerText.Replace('.', ','));
                    }

                }
                else
                {
                    //Kiểm tra tiếp
                    XmlNode TgTTTBSo = root.SelectSingleNode("//TToan//TgTTTBSo");
                    Thanhtien = double.Parse(TgTTTBSo.InnerText.Replace('.', ','));
                }
                if (string.IsNullOrEmpty(nTgTCThue))
                {
                    nTgTCThue = Thanhtien.ToString();
                } 
                //Tìm tổng tiền
                if (nTPhi!=null && !string.IsNullOrEmpty(nTPhi.InnerText))
                {
                    TPhi = double.Parse(nTPhi.InnerText);
                }
                if (!string.IsNullOrEmpty(nTgTCThue))
                {
                    TgTCThue = double.Parse(nTgTCThue);
                    if (TPhi != 0)
                    {
                        TgTCThue += TPhi;
                    }
                }
                if (!string.IsNullOrEmpty(nTgTThue))
                {
                    TgTThue = double.Parse(nTgTThue);
                }
                else
                {
                    TgTThue = 0;
                }
                //Kiểm tra thêm mới khách hàng
                if (mst == null)
                    mst = "";
                string querykh = @" SELECT TOP 1 *  FROM KhachHang As kh
WHERE kh.MST = ?"; // Sử dụng ? thay cho @mst trong OleDb

                lblThongbao.Text = "Kiểm tra khách hàng";
                Application.DoEvents();
                System.Data.DataTable result = ExecuteQuery(querykh, new OleDbParameter("?", mst));
                if (result.Rows.Count == 0 && !string.IsNullOrEmpty(mst))
                {
                    string diachi = nBanNode.SelectSingleNode("DChi")?.InnerText;
                    var Sohieu = GetLastFourDigits(mst.Replace("-",""));
                    ten = Helpers.ConvertUnicodeToVni(ten);
                    if (diachi != null)
                        diachi = Helpers.ConvertUnicodeToVni(diachi);
                    query = @" SELECT TOP 1 *  FROM KhachHang As kh
WHERE kh.SoHieu = ?";
                    System.Data.DataTable result2 = ExecuteQuery(query, new OleDbParameter("?", Sohieu));
                    if (result2.Rows.Count > 0)
                        Sohieu = "0" + Sohieu;
                    if (string.IsNullOrEmpty(diachi))
                    {
                        diachi = Helpers.ConvertUnicodeToVni("Bô sung địa chỉ");
                    }
                    InitCustomer(type == 1 ? 2 : 3, Sohieu, ten, diachi, mst);
                }
                 
                //Add detail
                var hhdVuList = xmlDoc.SelectNodes("//HHDVu");
                //Mật định tài khoản 
                //Kiểm tra Đã tồn tại số hóa đơn và số hiệu
                //if (!peopleTemp.Any(m => m.SHDon.Contains(SHDon) && m.KHHDon == KHHDon))
                if (!peopleTemp.Any(m => m.SHDon== SHDon && m.KHHDon == KHHDon))
                {
                    string diachi = nBanNode.SelectSingleNode("DChi")?.InnerText;
                    if (string.IsNullOrEmpty(mst) && string.IsNullOrEmpty(diachi))
                    {
                        ten= "Khách vãng lai";
                    }
                    if (string.IsNullOrEmpty(mst) && !string.IsNullOrEmpty(diachi))
                    {
                        string aa = RemoveVietnameseDiacritics(ten.Split(' ').LastOrDefault());
                        aa = CapitalizeFirstLetter(aa);
                        mst = ConvertToTenDigitNumber(aa).ToString();
                    }
                        peopleTemp.Add(new FileImport(file, SHDon, KHHDon, NLap, ten, diengiai, TkNo.ToString(), TkCo.ToString(), TkThue, mst, Thanhtien, Vat, 1, "",isAcess, TPhi, TgTCThue, TgTThue,true));
                }

                lblThongbao.Text = "Thêm danh sách sản phẩm con";
                Application.DoEvents();
                for (int i = 0; i < hhdVuList.Count; i++)
                {
                   
                    try
                    {
                        string tkno = peopleTemp.LastOrDefault().TKNo;
                        string tkco = peopleTemp.LastOrDefault().TKCo;
                        string mct = "";
                        if (hhdVuList[i].SelectSingleNode("DVTinh") != null && !string.IsNullOrEmpty(hhdVuList[i].SelectSingleNode("DVTinh").ToString()))
                        {
                            var THHDVu = hhdVuList[i].SelectSingleNode("THHDVu").InnerText;
                            var DVTinh = hhdVuList[i].SelectSingleNode("DVTinh").InnerText; 
                            if (!string.IsNullOrEmpty(DVTinh))
                            {
                                var SLuong = hhdVuList[i].SelectSingleNode("SLuong").InnerText.Replace('.', ',');
                                var DGia = "";
                                if (hhdVuList[i].SelectSingleNode("DGia") != null)
                                    DGia = hhdVuList[i].SelectSingleNode("DGia").InnerText.Replace('.', ',');
                                else
                                    DGia = "0";
                                //string newName = Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(THHDVu.Trim()));
                                string newName = Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(THHDVu.Trim()));
                                //Kiểm tra trong database xem có sản phẩm chưa, nếu chưa có thì thêm mới
                                query = @"SELECT * FROM Vattu 
WHERE LCase(TenVattu) = LCase(?) AND LCase(DonVi) = LCase(?)";

                                //int rs = (int)ExecuteQuery(query, new OleDbParameter("?", "SAdsd")).Rows[0][0];
                                var getdata = ExecuteQuery(query, new OleDbParameter("?", newName.ToLower()), new OleDbParameter("?", Helpers.ConvertUnicodeToVni(DVTinh).ToLower()));
                                //Kiểm tra thêm trong list
                                string newdvt = Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(DVTinh)).ToLower();
                                var checkold = peopleTemp.ToList().Where(n=>n.fileImportDetails.Any(m => m.Ten.ToLower() == newName.ToLower() && Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(m.DVT.ToLower())) == newdvt)).FirstOrDefault();
                                 
                                string sohieu = "";
                                if (getdata.Rows.Count == 0)
                                {
                                    if (checkold == null)
                                        sohieu = GenerateResultString(NormalizeVietnameseString(THHDVu.Trim()));
                                    else
                                        sohieu = checkold.fileImportDetails.Where(m => m.Ten.ToLower() == newName.ToLower() && Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(m.DVT.ToLower())) == newdvt).FirstOrDefault().SoHieu;
                                }
                                else
                                    sohieu = getdata.Rows[0]["SoHieu"].ToString();

                                //Gán giá trị cho các giá trị ""
                                DGia = !string.IsNullOrEmpty(DGia) ? DGia : "0";
                                SLuong = !string.IsNullOrEmpty(SLuong) ? SLuong : "0";
                                //Thiết lập MÃ ctrinh2 và tkno cho detail 
                                FileImportDetail fileImportDetail = new FileImportDetail(newName, peopleTemp.LastOrDefault().ID, sohieu.ToUpper(), double.Parse(SLuong), double.Parse(DGia), DVTinh, mct, tkno, tkco,0);
                                peopleTemp.LastOrDefault().fileImportDetails.Add(fileImportDetail);
                            }
                           
                        }
                        else
                        {
                            var THHDVu = hhdVuList[i].SelectSingleNode("THHDVu").InnerText;
                            if (THHDVu.ToLower().Contains("chiết khấu"))
                            {
                                var ThTien = hhdVuList[i].SelectSingleNode("ThTien")?.InnerText;
                                if (ThTien == null)
                                    ThTien = hhdVuList[i].SelectSingleNode("THTien")?.InnerText;
                                if (hhdVuList.Count == 1)
                                {
                                    FileImportDetail fileImportDetail = new FileImportDetail(THHDVu, peopleTemp.LastOrDefault().ID, "711", 1, double.Parse(ThTien), "Exception", "", "","",0);
                                    peopleTemp.LastOrDefault().TKNo = "331";
                                    peopleTemp.LastOrDefault().TKCo = "711";
                                    peopleTemp.LastOrDefault().TkThue = 1331;
                                    peopleTemp.LastOrDefault().Noidung = "Chiếc khấu thương mại";
                                    peopleTemp.LastOrDefault().fileImportDetails.Add(fileImportDetail);
                                }
                                else
                                {
                                    FileImportDetail fileImportDetail = new FileImportDetail(THHDVu, peopleTemp.LastOrDefault().ID, "711", 0, double.Parse(ThTien), "Exception", "", "","",0);
                                    peopleTemp.LastOrDefault().fileImportDetails.Add(fileImportDetail);
                                }

                            }
                            else
                            {
                                var ThTien = hhdVuList[i].SelectSingleNode("ThTien")?.InnerText;
                                if (ThTien == null)
                                    ThTien = hhdVuList[i].SelectSingleNode("THTien")?.InnerText;
                                if (ThTien != null && double.Parse(ThTien) > 0)
                                {
                                    FileImportDetail fileImportDetail = new FileImportDetail(THHDVu, peopleTemp.LastOrDefault().ID, "6422", 0, double.Parse(ThTien), "Exception", "", "","", 0);
                                    peopleTemp.LastOrDefault().fileImportDetails.Add(fileImportDetail);
                                }

                            }

                        }
                    }
                    catch (Exception ex)
                    {

                    }
                    //Kiểm tra nếu ko có con thì tk cha sẽ là 6240
                   
                }
                if (peopleTemp.LastOrDefault().fileImportDetails.Count == 0)
                {
                    if (type == 1)
                        peopleTemp.LastOrDefault().TKNo = "6422";
                }
            }
            //Trường hợp không đủ info
            //Th1 có 1 sản phẩm và ko có đơn vị tính
            string querydinhdanh = @" SELECT *  FROM tbDinhdanhtaikhoan"; // Sử dụng ? thay cho @mst trong OleDb

            result = ExecuteQuery(querydinhdanh, new OleDbParameter("?", ""));
            if (type == 1)
                querydinhdanh = @" SELECT *  FROM tbDinhdanhtaikhoan where KeyValue like '%Ưu tiên vào%'"; // Sử dụng ? thay cho @mst trong OleDb
            if (type == 2)
                querydinhdanh = @" SELECT *  FROM tbDinhdanhtaikhoan where KeyValue like '%Ưu tiên ra%'"; // Sử dụng ? thay cho @mst trong OleDb

           var result3 = ExecuteQuery(querydinhdanh, new OleDbParameter("?", ""));
            foreach (var item in peopleTemp)
            {
                if (item.fileImportDetails.Count == 1 && string.IsNullOrEmpty(item.fileImportDetails[0].DVT))
                {
                    item.TKNo = "6422";
                    item.TKCo = "1111";
                    item.TkThue = 1331;
                }
                if (item.TKNo == "0")
                {
                    //Nếu có con
                    if (item.fileImportDetails.Count > 0)
                    {
                        foreach (DataRow row in result.Rows)
                        {
                            string[] conditions = row["KeyValue"].ToString().Split('&');
                            string name = Helpers.ConvertUnicodeToVni((string)row["KeyValue"]);
                            int hasdata = 0;
                            string key = "";
                            foreach (string condition in conditions)
                            {
                                string[] parts = Regex.Split(condition, @"([><=%]+)"); // Vẫn giữ % để linh hoạt nếu cần
                                if (parts.Length == 3)
                                {
                                    key = parts[0];
                                    string operatorStr = parts[1];
                                    string valueStr = parts[2];
                                    if (key == "Ten")
                                    {
                                        //var newname = Helpers.ConvertUnicodeToVni(valueStr);
                                        var newname = valueStr;
                                        string chuoiBinhThuong = newname.Replace("\\\"", "\"").Trim('"');
                                        chuoiBinhThuong = Helpers.ConvertUnicodeToVni(chuoiBinhThuong);
                                        var check = item.fileImportDetails.Where(m => m.Ten.Contains(chuoiBinhThuong)).FirstOrDefault();
                                        if (check != null)
                                        {
                                            hasdata += 1;
                                        }
                                    }
                                    if (key == "TongTien")
                                    {
                                        var check = item.TongTien > double.Parse(valueStr);
                                        if (check)
                                            hasdata += 1;
                                    }
                                    if (key == "MST")
                                    {
                                        if (item.Mst == valueStr)
                                            hasdata += 1;
                                    }
                                }

                            }
                            if (hasdata == conditions.Count() && item.TKNo == "0")
                            {
                                if (key != "TongTien")
                                    item.TKNo = row["TKNo"].ToString();
                                item.TKCo = row["TKCo"].ToString();
                                item.TkThue = int.Parse(row["TkThue"].ToString());
                                if (string.IsNullOrEmpty(item.Noidung) && key == "MST" && item.TKNo == "6422")
                                    item.Noidung = row["Type"].ToString();
                            }
                            //if (item.fileImportDetails.Any(m => m.Ten.Contains(name)))
                            //{
                            //    item.Noidung = row["Type"].ToString();
                            //    item.TKNo = row["TKNo"].ToString();
                            //    item.TKCo = row["TKCo"].ToString();
                            //    item.TkThue = int.Parse(row["TKThue"].ToString());
                            //}
                        }
                    }
                    else
                    {
                        item.TKCo = "1111";
                        item.TkThue = 1331;
                    }
                }
                if (result3.Rows.Count > 0)
                {
                    foreach (DataRow row in result3.Rows)
                    {
                        if (string.IsNullOrEmpty(item.TKNo) || item.TKNo == "0")
                            item.TKNo = row["TKNo"].ToString();
                        if (string.IsNullOrEmpty(item.TKCo) || item.TKCo == "0")
                            item.TKCo = row["TKCo"].ToString();
                        if (item.TkThue == 0)
                            item.TkThue = int.Parse(row["TkThue"].ToString());
                        //Cho truong hop 331 711
                        if (item.TKNo == "331")
                        {
                            //item.TKNo = "711";
                            //item.TKCo = "3311";
                        }
                    }
                }
                if (item.fileImportDetails.Count > 0)
                {
                    if (string.IsNullOrEmpty(item.Noidung))
                        item.Noidung = Helpers.ConvertVniToUnicode(item.fileImportDetails.FirstOrDefault().Ten);
                }
                foreach (var t2 in item.fileImportDetails)
                {
                    t2.TKNo = item.TKNo;
                    t2.TKCo = item.TKCo;
                }
                if (item.isAcess == false)
                {
                    //item.Noidung = "Mã số thuế không hợp lệ";
                    item.Checked = false;
                }
                if (item.TPhi == null)
                {
                    
                }
            }
             
            progressBarControl1.EditValue = 100;
            //Fill cho people
            if (type == 1)
                people = peopleTemp;
            if (type == 2)
                people2 = peopleTemp;
             
        }
        private int CountExcelRows(List<string> excelFiles, DateTime startDate, DateTime endDate)
        {
            int rowCount = 0;
            foreach (var file in excelFiles)
            {
                try
                {
                    using (var workbook = new XLWorkbook(file))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var currentRow = worksheet.Cell("A7");
                        while (!currentRow.IsEmpty())
                        {
                            // Sử dụng GetFormattedString để lấy giá trị từ ô
                            string ngayLapStr = worksheet.Cell(currentRow.Address.RowNumber, 5).GetFormattedString()?.Trim();

                            if (DateTime.TryParse(ngayLapStr, out DateTime ngayLap) && ngayLap >= startDate && ngayLap <= endDate)
                            {
                                rowCount++;
                            }

                            currentRow = currentRow.Worksheet.Row(currentRow.Address.RowNumber + 1).Cell("A");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi khi đếm dòng Excel: {ex.Message}");
                }
            }
            return rowCount;
        }
        private System.Data.DataTable LoadExistingData(string tableName, string keyColumn1, string keyColumn2)
        {
            string query = $"SELECT {keyColumn1}, {keyColumn2} FROM {tableName}";
            return ExecuteQuery(query, null);
        }
        private System.Data.DataTable LoadExistingData2(string tableName, string keyColumn1, string keyColumn2, string keyColumn3)
        {
            string query = $"SELECT {keyColumn1}, {keyColumn2}, {keyColumn3} FROM {tableName}";
            return ExecuteQuery(query, null);
        }

        private async Task<List<VatTu>> LoadDataVattuAsync()
        {
            var query = @" SELECT * FROM Vattu ";
            var ListVattu = await Task.Run(() => ExecuteQuery(query, null));
            var lstVattu = new List<VatTu>();

            foreach (DataRow item in ListVattu.Rows)
            {
                item["TenVattu"] = Helpers.ConvertVniToUnicode(item["TenVattu"].ToString());
                item["DonVi"] = Helpers.ConvertVniToUnicode(item["DonVi"].ToString());
            }

            List<Task<VatTu>> vatTuTasks = new List<Task<VatTu>>();

            foreach (DataRow item in ListVattu.Rows)
            {
                vatTuTasks.Add(Task.Run(() =>
                {
                    var VatTu = new VatTu
                    {
                        MaSo = int.Parse(item["MaSo"].ToString()),
                        MaPhanLoai = int.Parse(item["MaPhanLoai"].ToString()),
                        TenVattu = item["TenVattu"].ToString(),
                        SoHieu = item["SoHieu"].ToString(),
                        DonVi = item["DonVi"].ToString(),
                        GhiChu = item["GhiChu"].ToString()
                    };

                    // Truy vấn số lượng và thành tiền
                    int cnt = 12;
                    var queryTonKho = @" SELECT * FROM TonKho WHERE MaVatTu= ? ";
                    var parameters = new OleDbParameter[] { new OleDbParameter("?", VatTu.MaSo) };

                    var kq = ExecuteQuery(queryTonKho, parameters);
                    if (kq.Rows.Count > 0)
                    {
                        while (cnt > 0 && kq.Rows[0]["Luong_" + cnt].ToString() == "0")
                        {
                            cnt--;
                        }

                        var soluong = kq.Rows[0]["Luong_" + cnt] != DBNull.Value ? double.Parse(kq.Rows[0]["Luong_" + cnt].ToString()) : 0;
                        VatTu.SoLuong = soluong;
                        var thanhtien = kq.Rows[0]["Tien_" + cnt] != DBNull.Value ? double.Parse(kq.Rows[0]["Tien_" + cnt].ToString()) : 0;
                        VatTu.ThanhTien = thanhtien;

                        if (soluong != 0 && thanhtien != 0)
                        {
                            VatTu.Dongia = thanhtien / soluong;
                        }
                    }

                    return VatTu;
                }));
            }

            // Chờ cho tất cả các tác vụ hoàn thành
            var vatTus = await Task.WhenAll(vatTuTasks);
            lstVattu.AddRange(vatTus);

            return lstVattu;
        }
        private Dictionary<string, DataRow> LoadExistingKhachHang(string tableName, string keyColumn)
        {
            string query = $"SELECT * FROM {tableName}";
            System.Data.DataTable dt = ExecuteQuery(query, null);
            Dictionary<string, DataRow> khachHangDictionary = new Dictionary<string, DataRow>();
            foreach (DataRow row in dt.Rows)
            {
                string key = row[keyColumn]?.ToString().ToLower();
                if (!string.IsNullOrEmpty(key) && !khachHangDictionary.ContainsKey(key))
                {
                    khachHangDictionary.Add(key, row);
                }
                // Bạn có thể thêm một else ở đây để xử lý các trường hợp trùng lặp nếu cần
                // Ví dụ: ghi log, chọn hàng đầu tiên, v.v.
            }
            return khachHangDictionary;
        }


        private Dictionary<string, DataRow> LoadExistingVatTu(string tableName, string keyColumn1, string keyColumn2)
        {
            string query = $"SELECT * FROM {tableName}";
            System.Data.DataTable dt = ExecuteQuery(query, null);
            Dictionary<string, DataRow> vatTuDictionary = new Dictionary<string, DataRow>();
            foreach (DataRow row in dt.Rows)
            {
                string key1 = row[keyColumn1]?.ToString().Trim().ToLower();
                string key2 = row[keyColumn2]?.ToString().Trim().ToLower();
                string key = $"{key1}-{key2}";
                if (!string.IsNullOrEmpty(key1) && !string.IsNullOrEmpty(key2) && !vatTuDictionary.ContainsKey(key))
                {
                    vatTuDictionary.Add(key, row);
                }
                // Bạn có thể thêm một else ở đây để xử lý các trường hợp trùng lặp nếu cần
                // Ví dụ: ghi log, chọn hàng đầu tiên, v.v.
            }
            return vatTuDictionary;
        }


        private System.Data.DataTable LoadDinhDanhTaiKhoan()
        {
            string query = @"SELECT * FROM tbDinhdanhtaikhoan";
            return ExecuteQuery(query, null);
        }

        private System.Data.DataTable LoadDinhDanhTaiKhoanUuTien(int type)
        {
            string query = (type == 1)
                ? @"SELECT * FROM tbDinhdanhtaikhoan WHERE KeyValue LIKE '%Ưu tiên vào%'"
                : @"SELECT * FROM tbDinhdanhtaikhoan WHERE KeyValue LIKE '%Ưu tiên ra%'";
            return ExecuteQuery(query, null);
        }
        private DataRow LoadKhachHangByMST(string mst)
        {
            string query = "SELECT TOP 1 * FROM KhachHang WHERE MST = ?";
            System.Data.DataTable result = ExecuteQuery(query, new OleDbParameter("?", mst));
            return result.Rows.Count > 0 ? result.Rows[0] : null;
        }
        private string GenerateKhachVangLaiMST(string ten, string diachi)
        {
            // Tạo MST tạm thời cho khách vãng lai dựa trên tên hoặc địa chỉ
            string baseString = string.IsNullOrEmpty(ten) ? diachi : ten;
            string normalized = RemoveVietnameseDiacritics(baseString.ToLower().Replace(" ", ""));
            long numericValue = 0;
            foreach (char c in normalized)
            {
                numericValue = numericValue * 31 + c; // Sử dụng một hàm hash đơn giản
            }
            return "KVL" + Math.Abs(numericValue % 10000000000).ToString("D10"); // Ví dụ: KVL followed by 10 digits
        }
        private bool EvaluateCondition(double value1, string operatorStr, double value2)
        {
            if (operatorStr == ">")
                return value1 > value2;
            else if (operatorStr == ">=")
                return value1 >= value2;
            else if (operatorStr == "<")
                return value1 < value2;
            else if (operatorStr == "<=")
                return value1 <= value2;
            else if (operatorStr == "==")
                return Math.Abs(value1 - value2) < 1e-9; // So sánh số thực
            else if (operatorStr == "!=")
                return Math.Abs(value1 - value2) >= 1e-9;
            else
                return false;
        }
        private void ApplyDefaultAndRuleBasedAccounts(FileImport item, System.Data.DataTable dinhDanhChung, System.Data.DataTable dinhDanhUuTien)
        {
            if (item.fileImportDetails.Count == 1 && string.IsNullOrEmpty(item.fileImportDetails[0].DVT))
            {
                item.TKNo = "6422";
                item.TKCo = "1111";
                item.TkThue = 1331;
            }

            //if (string.IsNullOrEmpty(item.TKNo) || item.TKNo == "0")
            //{
            //    if (item.fileImportDetails.Count > 0)
            //    {

            //    }
            //    else
            //    {
            //        item.TKCo = "1111";
            //        item.TkThue = 1331;
            //    }
            //}
            foreach (DataRow row in dinhDanhChung.Rows)
            {
                string[] conditions = row["KeyValue"].ToString().Split('&');
                int hasMatch = 0;
                foreach (string condition in conditions)
                {
                    string[] parts = Regex.Split(condition, @"([><=%]+)");
                    if (parts.Length == 3)
                    {
                        string key = parts[0];
                        string operatorStr = parts[1];
                        string valueStr = parts[2].Trim('"');
                        string chuoiBinhThuong = Helpers.ConvertUnicodeToVni(valueStr);

                        foreach (var detail in item.fileImportDetails)
                        {
                            if (key == "Ten" && !string.IsNullOrEmpty(detail.Ten) && detail.Ten.IndexOf(chuoiBinhThuong, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                hasMatch++;
                                break; // Chỉ cần một chi tiết thỏa mãn
                            }
                        }
                        if (key == "TongTien" && double.TryParse(valueStr, out var val) && EvaluateCondition(item.TongTien, operatorStr, val))
                        {
                            hasMatch++;
                        }
                        if (key == "MST" && item.Mst?.Equals(valueStr, StringComparison.OrdinalIgnoreCase) == true)
                        {
                            hasMatch++;
                        }
                    }
                }

                if (hasMatch == conditions.Length)
                {
                    item.TKNo = row["TKNo"].ToString();
                    item.TKCo = row["TKCo"].ToString();
                    item.TkThue = int.Parse(row["TkThue"].ToString());
                    if ( conditions.Any(c => c.StartsWith("MST")) && item.TKNo == "6422")
                    {
                        item.Noidung = row["Type"].ToString();
                    }
                    return; // Tìm thấy quy tắc phù hợp, thoát
                }
            }
            // Áp dụng quy tắc ưu tiên
            if (dinhDanhUuTien.Rows.Count > 0)
            {
                foreach (DataRow row in dinhDanhUuTien.Rows)
                {
                    if (string.IsNullOrEmpty(item.TKNo) || item.TKNo == "0") item.TKNo = row["TKNo"].ToString();
                    if (string.IsNullOrEmpty(item.TKCo) || item.TKCo == "0") item.TKCo = row["TKCo"].ToString();
                    if (item.TkThue == 0) item.TkThue = int.Parse(row["TkThue"].ToString());
                    if (item.TKNo == "331")
                    {
                        // Ví dụ về logic đặc biệt cho TKNo = 331
                        // item.TKNo = "711";
                        // item.TKCo = "3311";
                    }
                }
            }
        }
        static string RemoveSpecialCharacters(string input)
        {
            // Biểu thức chính quy để xóa ký tự đặc biệt
            return Regex.Replace(input, @"[^\w\s]", string.Empty);
        }

        public int currentselectId = 0;
        protected override bool ProcessCmdKey(ref Message msg, System.Windows.Forms.Keys keyData)
        {
            // Kiểm tra phím tắt (ví dụ: Ctrl + N)
            if (keyData == (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N))
            {
                AddNewChildRow();
                return true; // Đã xử lý phím
            }
            if (keyData == (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.D))
            {
                GridView parentView = gridControl1.MainView as GridView;

                // Lấy chỉ số dòng đang được chọn trong GridView cha
                int focusedRowHandle = parentView.FocusedRowHandle;

                 string noidung=  parentView.GetRowCellValue(focusedRowHandle, "Noidung").ToString();
                string tenkh= Helpers.ConvertUnicodeToVni(parentView.GetRowCellValue(focusedRowHandle, "Ten").ToString());
                string mst = Helpers.ConvertUnicodeToVni(parentView.GetRowCellValue(focusedRowHandle, "Mst").ToString());
                string query = @"select * FROM KhachHang 
                                    WHERE MST = ? ";
                var parameterss = new OleDbParameter[]
                {
                new OleDbParameter("?",mst)
                   };
                var kq = ExecuteQuery(query, parameterss);
                if (kq.Rows.Count > 0)
                {
                    if (xtraTabControl2.SelectedTabPageIndex == 0)
                    {
                        foreach (var item in lstImportVao)
                        {
                            if (item.Mst == mst)
                            {
                                item.Noidung = noidung;
                                query = @"UPDATE tbimport SET Noidung=?  WHERE ID=?";
                                var parameters = new OleDbParameter[]
                         {
                                new OleDbParameter("?", Helpers.ConvertUnicodeToVni(noidung)),
                                   new OleDbParameter("?",item.ID),
                         };
                                int rowsAffected = ExecuteQueryResult(query, parameters);
                            }
                        }
                        gridControl1.DataSource = lstImportVao;
                        gridControl1.RefreshDataSource();
                    } 
                        //Mật định tài khoản
                        string getMST = kq.Rows[0]["MST"].ToString();
                      query = @"select * FROM tbDinhdanhtaikhoan 
                                    WHERE KeyValue LIKE ? ";
                     parameterss = new OleDbParameter[]
               {
                new OleDbParameter("?","MST="+getMST)
                  };
                    kq = ExecuteQuery(query, parameterss);
                    if (kq.Rows.Count == 0)
                    {
                        query = @"INSERT INTO tbDinhdanhtaikhoan (Type,KeyValue,TKNo,TKCo, TKThue) VALUES (?, ?, ?, ?, ?)";
                       var parameters = new OleDbParameter[]
                        {
            new OleDbParameter("?", noidung),
             new OleDbParameter("?","MST="+getMST),
              new OleDbParameter("?", 6422),
               new OleDbParameter("?", 1111),
                 new OleDbParameter("?", 1331),
                        };
                        int rowsAffected = ExecuteQueryResult(query, parameters);
                    }
                    else
                    {
                        query = @"UPDATE tbDinhdanhtaikhoan SET Type=?  WHERE ID=?";
                        var parameters = new OleDbParameter[]
                         {
            new OleDbParameter("?", noidung),
             new OleDbParameter("?",kq.Rows[0]["ID"].ToString()), 
                         };
                        int rowsAffected = ExecuteQueryResult(query, parameters);
                    }
                }

                return true; // Đã xử lý phím
            }
            return base.ProcessCmdKey(ref msg, keyData); // Chuyển tiếp cho xử lý tiếp
        }
        private void AddNewChildRow()
        {
            if (xtraTabControl2.SelectedTabPageIndex == 0)
            {
                GridView parentView = gridControl1.MainView as GridView;

                // Lấy chỉ số dòng đang được chọn trong GridView cha
                int focusedRowHandle = parentView.FocusedRowHandle;

                // Lấy GridView con
                GridView childView = parentView.GetDetailView(focusedRowHandle, 0) as GridView;
                if (childView != null)
                {
                    // Lấy chỉ số dòng đang được chọn trong GridView con
                    int focusedChildRowHandle = childView.FocusedRowHandle;
                    var parentId = childView.GetRowCellValue(focusedChildRowHandle, "ParentId");
                    int getID = (int)childView.GetRowCellValue(focusedChildRowHandle, "ID");
                    var getparent = lstImportVao.Where(m => m.ID == (int)parentId).FirstOrDefault();
                    if (getparent != null)
                    {
                        var getcurrent = getparent.fileImportDetails.Where(m => m.ID == getID).FirstOrDefault();
                        if (getcurrent != null)
                        {
                            getparent.fileImportDetails.Add(new FileImportDetail(getcurrent.Ten, getcurrent.ParentId, getcurrent.SoHieu, getcurrent.Soluong, getcurrent.Dongia, getcurrent.DVT, getcurrent.MaCT, getcurrent.TKNo, getcurrent.TKCo, getcurrent.TTien));
                            gridControl1.DataSource = lstImportVao;
                        }

                    }
                }
            }
            else
            {
                GridView parentView = gridControl2.MainView as GridView;

                // Lấy chỉ số dòng đang được chọn trong GridView cha
                int focusedRowHandle = parentView.FocusedRowHandle;

                // Lấy GridView con
                GridView childView = parentView.GetDetailView(focusedRowHandle, 0) as GridView;
                if (childView != null)
                {
                    // Lấy chỉ số dòng đang được chọn trong GridView con
                    int focusedChildRowHandle = childView.FocusedRowHandle;
                    var parentId = childView.GetRowCellValue(focusedChildRowHandle, "ParentId");
                    int getID = (int)childView.GetRowCellValue(focusedChildRowHandle, "ID");
                    var getparent = lstImportRa.Where(m => m.ID == (int)parentId).FirstOrDefault();
                    if (getparent != null)
                    {
                        var getcurrent = getparent.fileImportDetails.Where(m => m.ID == getID).FirstOrDefault();
                        if (getcurrent != null)
                        {
                            getparent.fileImportDetails.Add(new FileImportDetail(getcurrent.Ten, getcurrent.ParentId, getcurrent.SoHieu, getcurrent.Soluong, getcurrent.Dongia, getcurrent.DVT, getcurrent.MaCT, getcurrent.TKNo, getcurrent.TKCo, getcurrent.TTien));
                            gridControl2.DataSource = lstImportRa;
                        }

                    }
                }
            }
           
        } 
        public static bool AreNamesSimilar(string name1, string name2)
        {
            string cleanName1 = CleanUpName(name1);
            string cleanName2 = CleanUpName(name2);

            // Kiểm tra nếu một tên chứa tên còn lại
            return cleanName1.Contains(cleanName2) || cleanName2.Contains(cleanName1);
        }

        private static string CleanUpName(string name)
        {
            // Loại bỏ ký tự không phải chữ cái và số
            string cleaned = Regex.Replace(name, @"[^\w\s]", "");

            // Thay thế khoảng trắng đứng sau ký tự đặc biệt
            cleaned = Regex.Replace(cleaned, @"\s+", " "); // Đảm bảo chỉ có 1 khoảng trắng
            return cleaned.ToLower().Trim(); // Chuyển thành chữ thường và loại bỏ khoảng trắng thừa

        }
        public class Vattucheck
        {
            public string SoHieu { get; set; }
            public string Namecore { get; set; }
            public string Namecheck { get; set; }
            public double Percent { get; set; }
            public double Price { get; set; }
        }
        public static string FormatNumber(double number)
        {
            // Chuyển số thành chuỗi và loại bỏ các số 0 không cần thiết ở cuối
            string formatted = number.ToString("0.###########", CultureInfo.InvariantCulture);

            // Nếu có phần thập phân toàn số 0 thì bỏ phần thập phân
            if (formatted.Contains("."))
            {
                string decimalPart = formatted.Split('.')[1];
                if (decimalPart.TrimEnd('0').Length == 0)
                {
                    formatted = formatted.Split('.')[0];
                }
                else
                {
                    formatted = formatted.TrimEnd('0').TrimEnd('.');
                }
            }

            return formatted;
        }

        public static string FormatNumber(string numberString)
        {
            if (double.TryParse(numberString, NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
            {
                return FormatNumber(number);
            }
            return numberString; // Trả về nguyên bản nếu không phải số
        }
        private void LoadXmlFilesOptimized(string path, int type)
            { 

            progressPanel1.Visible = true;
            lblThongbao.Text = "Bắt đầu chạy";
            progressPanel1.Caption = "Bắt đầu chạy...";

            BindingList<FileImport> currentPeopleList = (type == 1) ? people = new BindingList<FileImport>() : people2 = new BindingList<FileImport>();
            BindingList<FileImport> peopleTemp = new BindingList<FileImport>();
            progressBarControl1.EditValue = 0;
            Application.DoEvents();

            string dataPath = path + (type == 1 ? "\\HDVao" : "\\HDRa");
            int startMonth = dtTungay.DateTime.Month;
            int endMonth = dtDenngay.DateTime.Month;
            DateTime startDate = dtTungay.DateTime.Date;
            DateTime endDate = dtDenngay.DateTime.Date.AddDays(1).AddSeconds(-1); // Để bao gồm cả ngày cuối

            //lblThongbao.Text = "Đếm và lọc file";
            //progressPanel1.Caption = "Đếm và lọc file...";
            //Application.DoEvents();
            var allFiles = Directory.EnumerateFiles(dataPath, "*.*", SearchOption.AllDirectories)
                .Where(file => IsFileInMonthRange(file, dataPath, startMonth, endMonth)).ToList();

            var xmlFiles = allFiles.Where(f => f.ToLower().EndsWith(".xml")).ToList();
            var excelFiles = allFiles.Where(f => f.ToLower().EndsWith(".xlsx")).ToList();
            //List<string> lstdelete = new List<string>();
            //foreach (var item in xmlFiles)
            //{
            //    XmlDocument xmlDoc = new XmlDocument();
            //    string fullPath = item;
            //    using (StreamReader reader = new StreamReader(fullPath, Encoding.UTF8))
            //    {
            //        try
            //        {
            //            xmlDoc.Load(reader); // Tải file XML
            //        }
            //        catch (XmlException ex)
            //        {
            //            Console.WriteLine($"Lỗi khi tải file XML: {ex.Message}");
            //            return;
            //        }
            //        XmlNode root = xmlDoc.DocumentElement;
            //        XmlNode nTTChungNode = root.SelectSingleNode("//TTChung");
            //        DateTime NLap = DateTime.Parse(nTTChungNode.SelectSingleNode("NLap")?.InnerText);
            //        if (NLap >= dtTungay.DateTime && NLap <= dtDenngay.DateTime)
            //            continue;
            //        else
            //            lstdelete.Add(item);

            //    }
            //}
            //foreach (var item in lstdelete)
            //{
            //    xmlFiles.Remove(item);
            //}
             totalCount = xmlFiles.Count + CountExcelRows(excelFiles, startDate, endDate);
            //lblSofiles.Text = totalCount.ToString();
            Application.DoEvents();
            if (type == 2) lblSofiles2.Text = totalCount.ToString();


            // Tải dữ liệu hóa đơn và tbimport vào bộ nhớ để kiểm tra nhanh hơn
            System.Data.DataTable existingHoaDon = LoadExistingData2("HoaDon", "KyHieu", "SoHD","NgayPH");
            System.Data.DataTable existingTbImport = LoadExistingData("tbimport", "KHHDon", "SHDon");
            System.Data.DataTable existingTbImportVattu = LoadExistingData2("Vattu", "TenVattu", "DonVi", "SoHieu");
            Dictionary<string, DataRow> existingKhachHang = LoadExistingKhachHang("KhachHang", "MST");
            Dictionary<string, DataRow> existingKhachHang2 = LoadExistingKhachHang("KhachHang", "Ten");
            Dictionary<string, DataRow> existingVatTu = LoadExistingVatTu("Vattu", "TenVattu", "DonVi");
            System.Data.DataTable tbDinhDanhtaikhoan = LoadDinhDanhTaiKhoan();
            System.Data.DataTable tbDinhDanhtaikhoanUuTien = LoadDinhDanhTaiKhoanUuTien(type);

            foreach (string xmlFile in xmlFiles)
            {
               // lblThongbao.Text = $"Đọc file XML thứ {filesLoaded + 1}/{totalCount}";
                progressPanel1.Caption = $"Đọc file XML thứ {filesLoaded + 1}/{totalCount}"; 
               // progressBarControl1.EditValue = progressPercentage;
                Application.DoEvents();
                filesLoaded++;

                XmlDocument xmlDoc = new XmlDocument();
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.XmlResolver = null;
                settings.DtdProcessing = DtdProcessing.Ignore;
                settings.CheckCharacters = false;
                try
                {
                    using (StreamReader sr = new StreamReader(xmlFile, Encoding.UTF8, true))
                    {
                        xmlDoc.Load(sr);
                    }
                }
                catch (XmlException ex)
                {
                    Console.WriteLine($"Lỗi khi tải file XML: {ex.Message}");
                    continue; // Bỏ qua file lỗi và tiếp tục
                }

                XmlNode root = xmlDoc.DocumentElement;
                XmlNode nTTChungNode = root?.SelectSingleNode("//TTChung");
                XmlNode ndhDonNode = root?.SelectSingleNode("//NDHDon");

                if (nTTChungNode == null || ndhDonNode == null) continue;

                string sHDon = nTTChungNode.SelectSingleNode("SHDon")?.InnerText;
                string kHHDon = nTTChungNode.SelectSingleNode("KHHDon")?.InnerText;
                DateTime nLap = DateTime.TryParse(nTTChungNode.SelectSingleNode("NLap")?.InnerText, out var date) ? date : DateTime.MinValue;
                var getmonth = nLap.Month;
                // Kiểm tra trùng lặp trong database
                if (existingHoaDon.Rows.Cast<DataRow>().Any(row => row["KyHieu"]?.ToString() == kHHDon && row["SoHD"]?.ToString().Contains(sHDon) == true && DateTime.Parse(row["NgayPH"]?.ToString()).Month== getmonth)) continue;

                // Kiểm tra trùng lặp trong BindingList tạm thời
                if (peopleTemp.Any(m => m.SHDon == sHDon && m.KHHDon == kHHDon)) continue;

                // Kiểm tra trùng lặp trong tbimport
                if (existingTbImport.Rows.Cast<DataRow>().Any(row => row["KHHDon"]?.ToString() == kHHDon && row["SHDon"]?.ToString().Contains(sHDon) == true)) continue;

                XmlNode nBanNode = (type == 1) ? ndhDonNode.SelectSingleNode("NBan") : ndhDonNode.SelectSingleNode("NMua");
                var getten = nBanNode?.SelectSingleNode("Ten");
                string ten = "";
                if (getten == null )
                {
                    ten = nBanNode?.SelectSingleNode("HVTNMHang").InnerText;
                    //HVTNMHang 
                }
                else 
                {
                    if(string.IsNullOrEmpty(getten.InnerText))
                    {
                        ten = nBanNode?.SelectSingleNode("HVTNMHang").InnerText;
                    }
                    else
                    {
                        ten = getten.InnerText;
                    }
                }
                    //ten = RemoveSpecialCharacters(ten).Trim();
                    string mst = nBanNode?.SelectSingleNode("MST")?.InnerText;
                string diachiBan = nBanNode?.SelectSingleNode("DChi")?.InnerText;
                //Thêm khách hàng
                //InitCustomer(type == 1 ? 2 : 3, "", ten, diachiBan, mst);
                //Trường hợp không có MST

                //if (string.IsNullOrEmpty(mst) && root.SelectSingleNode("//NMua//DChi") != null)
                //{
                //    string getdc = root.SelectSingleNode("//NMua//DChi").InnerText;
                //    InitCustomer(type == 1 ? 2 : 3, "", ten, getdc, mst);
                //}
                //else if (string.IsNullOrEmpty(mst) && !string.IsNullOrEmpty(ten))
                //{
                //    InitCustomer(type == 1 ? 2 : 3, "", ten, "", mst);
                //}

                // bool isAcess = (type == 1 && root.SelectSingleNode("//NMua//MST")?.InnerText == mstcongty) || (type == 2 && root.SelectSingleNode("//NBan//MST")?.InnerText == mstcongty);
                bool isAcess = true;
                XmlNodeList nTSuat = root.SelectNodes("//LTSuat//TSuat");
                int vat = 0;
                foreach (XmlNode item in nTSuat)
                {
                    if (item.InnerText != "KKKNT" && item.InnerText != "KCT")
                    {
                        vat = int.TryParse(item.InnerText.Replace("%", ""), out var v) ? v : vat;
                        break; // Lấy giá trị VAT đầu tiên hợp lệ
                    }
                }
                double thanhtien = 0;

                XmlNodeList nThTien = root.SelectNodes("//LTSuat//ThTien");
                if (nThTien.Count>0)
                {
                    nThTien[0].InnerText = FormatNumber(nThTien[0].InnerText?.ToString());
                    if (nThTien.Count > 0 && double.TryParse(nThTien[0].InnerText.Replace('.', ','), out var tt))
                    {
                        thanhtien = tt;
                    }
                }
                else 
                {
                    var ttso = root.SelectSingleNode("//TToan//TgTTTBSo")?.InnerText;
                    ttso = FormatNumber(ttso);
                    if (double.TryParse(ttso.Replace('.', ','), out var tt2))
                        thanhtien = tt2;
                }
                thanhtien= Math.Round(thanhtien, MidpointRounding.ToEven);
                XmlNode nTPhi = root.SelectSingleNode("//TToan//TPhi");
                double tPhi = double.TryParse(nTPhi?.InnerText, out var phi) ? phi : 0;

                string nTgTCThueStr = root.SelectSingleNode("//TToan//TgTCThue")?.InnerText;
                nTgTCThueStr = FormatNumber(nTgTCThueStr);
               // nTgTCThueStr = nTgTCThueStr.Replace('.', ',');
                double tgTCThue = double.TryParse(nTgTCThueStr, out var ttc) ? ttc : thanhtien;
                tgTCThue= Math.Round(tgTCThue, MidpointRounding.ToEven);
                if (tPhi != 0) tgTCThue += tPhi;

                string nTgTThueStr = root.SelectSingleNode("//TToan//TgTThue")?.InnerText;
                nTgTThueStr = FormatNumber(nTgTThueStr);
               // nTgTThueStr = nTgTThueStr?.Replace('.', ',');
                double tgTThue = double.TryParse(nTgTThueStr, out var tth) ? tth : 0;
                tgTThue = Math.Round(tgTThue, MidpointRounding.ToEven);
                string diengiai = "";
                int tkCo = 0;
                int tkNo = 0;
                int tkThue = 0;
                string sohieuKH = "";
                // Kiểm tra và thêm mới khách hàng (tối ưu hóa bằng cách sử dụng Dictionary)
                if (!string.IsNullOrEmpty(mst))
                {
                    if (!existingKhachHang.ContainsKey(mst))
                    {
                         sohieuKH = GetLastFourDigits(mst.Replace("-", ""));
                        string tenKHVni = Helpers.ConvertUnicodeToVni(ten);
                        string diachiKHVni = !string.IsNullOrEmpty(diachiBan) ? Helpers.ConvertUnicodeToVni(diachiBan) : Helpers.ConvertUnicodeToVni("Bổ sung địa chỉ");
                        if (existingKhachHang.Values.Any(kh => kh["SoHieu"]?.ToString() == sohieuKH))
                        {
                            sohieuKH = "0" + sohieuKH;
                        }
                        if (existingKhachHang.Values.Any(kh => kh["SoHieu"]?.ToString() == sohieuKH))
                        {
                            sohieuKH = "00" + sohieuKH;
                        }
                        InitCustomer(type == 1 ? 2 : 3, sohieuKH, tenKHVni, diachiKHVni, mst);
                        // Cập nhật Dictionary sau khi thêm mới (nếu cần cho các file sau)
                        DataRow newKhachHangRow = LoadKhachHangByMST(mst);
                        if (newKhachHangRow != null)
                        {
                            existingKhachHang[mst] = newKhachHangRow;
                        }
                    }
                    
                }
                else
                {
                    if (!existingKhachHang2.ContainsKey(Helpers.ConvertUnicodeToVni(ten.ToLower())))
                    {
                         sohieuKH = RemoveVietnameseDiacritics(ten.Split(' ').LastOrDefault());
                        //Sohieu = CapitalizeFirstLetter(Sohieu);
                        if(sohieuKH.Length>=3)
                        sohieuKH = sohieuKH.ToUpper().Substring(0,3);
                        int randNumber = random.Next(101, 999);
                        sohieuKH = sohieuKH + randNumber.ToString();
                        mst = "00";
                        ten = Helpers.ConvertUnicodeToVni(ten);
                        string diachiKHVni = !string.IsNullOrEmpty(diachiBan) ? Helpers.ConvertUnicodeToVni(diachiBan) : Helpers.ConvertUnicodeToVni("Bổ sung địa chỉ");
                        InitCustomer(type == 1 ? 2 : 3, sohieuKH, ten, diachiKHVni, mst);
                    }
                    else
                    {
                        mst = "00";
                        sohieuKH = existingKhachHang2.Where(m => m.Key == Helpers.ConvertUnicodeToVni(ten.ToLower())).FirstOrDefault().Value[2].ToString();
                    }
                }
                string newxmlFile = "";
                if (type == 1)
                    newxmlFile = xmlFile.Replace("HDVao", "HDVaoChonLoc").ToString();
                else
                    newxmlFile = xmlFile.Replace("HDRa", "HDRaChonLoc").ToString();

                FileImport newFileImport = new FileImport(newxmlFile, sHDon, kHHDon, nLap, ten ?? "Khách vãng lai", diengiai, tkNo.ToString(), tkCo.ToString(), tkThue, mst == "00" ? sohieuKH : mst, thanhtien, vat, type, "", isAcess, tPhi, tgTCThue, tgTThue,true);
                peopleTemp.Add(newFileImport);

                // Thêm chi tiết hóa đơn
                XmlNodeList hhdVuList = root.SelectNodes("//HHDVu");
                foreach (XmlNode hhdVu in hhdVuList)
                {
                    try
                    {
                        string thhdVu = hhdVu.SelectSingleNode("THHDVu")?.InnerText;
                        //thhdVu = RemoveSpecialCharacters(thhdVu);
                        string dvTinh = hhdVu.SelectSingleNode("DVTinh")?.InnerText;
                        string sLuongStr = hhdVu.SelectSingleNode("SLuong")?.InnerText;
                        sLuongStr = FormatNumber(sLuongStr);
                        string dGiaStr = hhdVu.SelectSingleNode("DGia")?.InnerText;
                        dGiaStr =  FormatNumber(dGiaStr);
                        //dGiaStr = dGiaStr?.Replace('.', ',');
                        string chietkhau = hhdVu.SelectSingleNode("STCKhau")?.InnerText?.Replace('.', ',');
                        string ttien = hhdVu.SelectSingleNode("ThTien")?.InnerText;
                        ttien = FormatNumber(ttien);
                      //  ttien =ttien?.Replace('.', ',');
                        double.TryParse(chietkhau, out var dChietkhau);
                       
                        if (!string.IsNullOrEmpty(dvTinh) && !string.IsNullOrEmpty(thhdVu) && double.TryParse(sLuongStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double sLuong) && double.TryParse(dGiaStr, out var dGia) && double.TryParse(ttien, out var dttien))
                        {
                            //if (dChietkhau != 0)
                            //{
                            //    dGia -= dChietkhau;
                            //}
                            //dGia= Math.Round(dGia, MidpointRounding.ToEven);
                            thhdVu = Regex.Replace(thhdVu, @"\s+", " ");
                            string tenVattuVni = Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(thhdVu.Trim())).ToLower();
                            string dvTinhVni = Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(dvTinh)).ToLower();
                            string soHieuVattu = "";
                            bool hasVattu = false;
                            DataRow getrow = null;
                            //Kểm tra vật tư
                            List<Vattucheck> lstChonloc = new List<Vattucheck>();

                            string querykh = @" SELECT *  FROM tbRegister"; // Sử dụng ? thay cho @mst trong OleDb

                            result = ExecuteQuery(querykh, new OleDbParameter("?", ""));
                            string col1 = result.Rows[0]["Col1"].ToString();
                            string col2 = result.Rows[0]["Col2"].ToString();

                            foreach (DataRow row in existingTbImportVattu.Rows)
                            {
                                //if (AreNamesSimilar(row["TenVattu"].ToString().ToLower(), tenVattuVni.ToLower()))
                                //{
                                //    hasVattu = true;
                                //    getrow = row; 
                                //}
                                string getRowname = row["TenVattu"].ToString().ToLower();
                                var getpercent= Helpers.StringWordSimilarity.CalculateSimilarity(getRowname, tenVattuVni.ToLower());
                                int percent = 0;
                                if (col1 != "1")
                                    percent = 100;
                                else
                                    percent = 70;
                                if (getpercent >= percent)
                                {
                                    Vattucheck Vattucheck = new Vattucheck();
                                    Vattucheck.Namecore = tenVattuVni.ToLower();
                                    Vattucheck.Namecheck = getRowname;
                                    Vattucheck.Percent = getpercent;
                                    Vattucheck.SoHieu = row["SoHieu"]?.ToString();
                                    if(col2=="1")
                                    {
                                        if (row["DonVi"].ToString().ToLower() == dvTinhVni)
                                        {
                                            lstChonloc.Add(Vattucheck);
                                        }
                                    }
                                    else
                                    {
                                        lstChonloc.Add(Vattucheck);
                                    }
                                  
                                }
                            }
                            var chk = lstChonloc.OrderByDescending(m => m.Percent).Where(n=>n.Percent>=70).FirstOrDefault();
                            // if (!existingVatTu.ContainsKey($"{tenVattuVni}-{dvTinhVni}") || !hasVattu)
                            if (chk == null)
                            {
                                // Kiểm tra trong list tạm thời của file hiện tại
                                //var existingDetail = newFileImport.fileImportDetails.FirstOrDefault(d => Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(d.Ten.Trim())).ToLower() == tenVattuVni && Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(d.DVT.Trim())).ToLower() == dvTinhVni);

                                FileImportDetail existingDetail = null;
                                foreach (var it in peopleTemp)
                                {
                                    //existingDetail = it.fileImportDetails.Where(d => Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(d.Ten.Trim())).ToLower() == tenVattuVni && Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(d.DVT.Trim())).ToLower() == dvTinhVni).FirstOrDefault();
                                    existingDetail = it.fileImportDetails.Where(d => Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(d.Ten.Trim())).ToLower() == tenVattuVni).FirstOrDefault();
                                    if (existingDetail != null)
                                       break;
                                }

                                soHieuVattu = existingDetail?.SoHieu ?? GenerateResultString(NormalizeVietnameseString(thhdVu.Trim()));
                                // Thêm mới vật tư (nếu cần và bạn muốn lưu vào DB ngay)
                                // InitVatTu(soHieuVattu, Helpers.ConvertUnicodeToVni(thhdVu.Trim()), Helpers.ConvertUnicodeToVni(dvTinh));
                                // Cập nhật Dictionary (nếu cần)
                                // DataRow newVatTuRow = LoadVatTuByTenAndDVT(tenVattuVni, dvTinhVni);
                                // if (newVatTuRow != null) existingVatTu[$"{tenVattuVni}-{dvTinhVni}"] = newVatTuRow;
                            }
                            else
                            {
                                soHieuVattu = chk.SoHieu;
                                //soHieuVattu= row["SoHieu"]?.ToString();
                                //soHieuVattu = existingVatTu[$"{tenVattuVni}-{dvTinhVni}"]["SoHieu"]?.ToString();
                            }

                            FileImportDetail fileImportDetail = new FileImportDetail(NormalizeVietnameseString(thhdVu), newFileImport.ID, soHieuVattu?.ToUpper(), sLuong, dGia, dvTinh, "", tkNo.ToString(), tkCo.ToString(), dttien);
                            newFileImport.fileImportDetails.Add(fileImportDetail);
                        }
                        else if (thhdVu?.ToLower().Contains("chiết khấu") == true)
                        {
                            string thTienStr = hhdVu.SelectSingleNode("ThTien")?.InnerText ?? hhdVu.SelectSingleNode("THTien")?.InnerText;
                            if (double.TryParse(thTienStr, out var thTien))
                            {
                                string soHieuCK = "711";
                                if (hhdVuList.Count == 1)
                                {
                                    newFileImport.TKNo = "331";
                                    newFileImport.TKCo = "711";
                                    newFileImport.TkThue = 1331;
                                    newFileImport.Noidung = "Chiết khấu thương mại";
                                }
                                newFileImport.fileImportDetails.Add(new FileImportDetail(thhdVu, newFileImport.ID, soHieuCK, 0, thTien, "Exception", "", "", "",0));
                            }
                        }
                        else if (double.TryParse(hhdVu.SelectSingleNode("ThTien")?.InnerText ?? hhdVu.SelectSingleNode("THTien")?.InnerText, out var thTienNN) && thTienNN > 0)
                        {
                            newFileImport.fileImportDetails.Add(new FileImportDetail(thhdVu, newFileImport.ID, "6422", 0, thTienNN, "Exception", "", "", "",0));
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Lỗi khi xử lý chi tiết hóa đơn: {ex.Message}");
                    }
                }

                if (newFileImport.fileImportDetails.Count == 0 && type == 1)
                {
                    newFileImport.TKNo = "6422";
                }
            }

            // Xử lý file Excel
            //foreach (string excelFile in excelFiles)
            //{
            //    lblThongbao.Text = $"Đọc file Excel thứ {filesLoaded + 1}/{totalCount}";
            //    progressPanel1.Caption = "Đang xử lý file Excel...";
            //    progressPercentage = (filesLoaded * 100) / totalCount;
            //    progressBarControl1.EditValue = progressPercentage;
            //    Application.DoEvents();
            //    filesLoaded++;
            //    try
            //    {
            //        using (var workbook = new XLWorkbook(excelFile))
            //        {
            //            var worksheet = workbook.Worksheet(1);
            //            var currentRow = worksheet.Cell("A7"); // Bắt đầu từ dòng 7

            //            while (!currentRow.IsEmpty())
            //            {
            //                // Lấy giá trị ngày lập từ cột 5
            //                string getNgayLapStr = currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 5).GetFormattedString().Trim();
            //                if (DateTime.TryParse(getNgayLapStr, out DateTime getNgayLap) && getNgayLap >= startDate && getNgayLap <= endDate)
            //                {
            //                    // Tạo một đối tượng FileImport từ dữ liệu Excel
            //                    string soHDE = currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 2).GetFormattedString().Trim();
            //                    string kHHDE = "";
            //                    string tenKH = currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 3).GetFormattedString().Trim();
            //                    string mstKH = currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 4).GetFormattedString().Trim();
            //                    string diachiKH = "";
            //                    double tongTienExcel = 0;
            //                    double thueGTGTExcel = 0;
            //                    double tongThanhToanExcel = 0;

            //                    // Lấy giá trị từ các cột 6, 7, 8
            //                    if (!currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 6).IsEmpty())
            //                    {
            //                        string tongTienStr = currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 6).GetFormattedString().Trim().Replace(",", "");
            //                        double.TryParse(tongTienStr, out tongTienExcel);
            //                    }
            //                    if (!currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 7).IsEmpty())
            //                    {
            //                        string thueGTGTStr = currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 7).GetFormattedString().Trim().Replace(",", "");
            //                        double.TryParse(thueGTGTStr, out thueGTGTExcel);
            //                    }
            //                    if (!currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 8).IsEmpty())
            //                    {
            //                        string tongThanhToanStr = currentRow.Worksheet.Cell(currentRow.Address.RowNumber, 8).GetFormattedString().Trim().Replace(",", "");
            //                        double.TryParse(tongThanhToanStr, out tongThanhToanExcel);
            //                    }

            //                    int vatExcel = 0;
            //                    if (tongTienExcel > 0)
            //                        vatExcel = (int)Math.Round((thueGTGTExcel / tongTienExcel) * 100);

            //                    bool isAcessExcel = (type == 1 && mstKH == mstcongty) || (type == 2 && mstKH == mstcongty);

            //                    // Kiểm tra trùng lặp
            //                    if (!existingHoaDon.Rows.Cast<DataRow>().Any(row => row["SoHD"]?.ToString().Contains(soHDE) == true))
            //                    {
            //                        FileImport excelImport = new FileImport(excelFile, soHDE, kHHDE, getNgayLap, tenKH, "", "", "", vatExcel, mstKH, tongTienExcel, 0, 2, "", isAcessExcel, 0, tongThanhToanExcel, thueGTGTExcel);
            //                        peopleTemp.Add(excelImport);
            //                    }
            //                }
            //                currentRow = currentRow.Worksheet.Row(currentRow.Address.RowNumber + 1).Cell("A");
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine($"Lỗi khi đọc file Excel: {ex.Message}");
            //    }
            //}

            // Gán tài khoản mặc định và theo quy tắc
            foreach (var item in peopleTemp)
            {
                ApplyDefaultAndRuleBasedAccounts(item, tbDinhDanhtaikhoan, tbDinhDanhtaikhoanUuTien);
                if (item.fileImportDetails.Count > 0 && string.IsNullOrEmpty(item.Noidung))
                {
                    item.Noidung = item.fileImportDetails.FirstOrDefault().Ten;
                }
                foreach (var detail in item.fileImportDetails)
                {
                    detail.TKNo = item.TKNo;
                    detail.TKCo = item.TKCo;
                }
                if (!item.isAcess)
                {
                    item.Checked = false;
                }
            }

            progressBarControl1.EditValue = 100;
            lblThongbao.Text = "Hoàn thành";

            // Fill cho BindingList chính
            if (type == 1)
                people = peopleTemp;
            else if (type == 2)
                people2 = peopleTemp;
        }
        private void XulyFolder()
        {

        }
        private void XulyfilEexcel(int type,int month)
        {
            string filePath = savedPath;
            filePath += (type == 1 ? "\\HDVao" : "\\HDRa");
            // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
            var excelFiles = Directory.EnumerateFiles(filePath, "*.xlsx", SearchOption.AllDirectories).Where(file => IsFileInMonthRange(file, filePath, month, month)).ToList(); ; // Kiểm tra xem file có nằm trong khoảng tháng
            for (int j = 0; j < excelFiles.Count; j++)
            {
                using (var workbook = new XLWorkbook(excelFiles[j]))
                {
                    var worksheet = workbook.Worksheet(1); // Lấy worksheet đầu tiên
                    int rowCount = 0;
                    var currentCell = worksheet.Cell("A6"); // Bắt đầu từ ô A6
                    int demdong = 0;
                    var rowsToDelete = new List<int>();

                    for (int i = 7; i <= worksheet.LastRowUsed().RowNumber(); i++)
                    {
                        var row = worksheet.Row(i);
                        var data = row.Cell(5);
                        var dateCell = row.Cell(5).GetValue<string>();
                        DateTime getDate = DateTime.Parse(dateCell);
                        // Kiểm tra xem ngày có nhỏ hơn 12 không
                        if (getDate.Month != month)
                        {
                            rowsToDelete.Add(i);
                        }
                    }
                    // Xóa các dòng từ dưới lên
                    foreach (var rowNumber in rowsToDelete.OrderByDescending(r => r))
                    {
                        worksheet.Row(rowNumber).Delete();
                    }
                    workbook.SaveAs(excelFiles[j]);
                }
            }
        }
         public void GetdetailXML2(string nbmst, string khhdon, string shdon, string tokken,int invoiceType)
        {
            string url = "";
            if(invoiceType==4)
            url = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/detail?nbmst={nbmst}&khhdon={khhdon}&shdon={shdon}&khmshdon=1";
            if (invoiceType == 5)
                url = $"https://hoadondientu.gdt.gov.vn:30000/sco-query/invoices/detail?nbmst={nbmst}&khhdon={khhdon}&shdon={shdon}&khmshdon=1";
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                try
                {
                    HttpResponseMessage response=new HttpResponseMessage();
                    // Gửi yêu cầu GET đồng bộ 
                    try
                    {
                        Thread.Sleep(400);
                        response = client.GetAsync(url).GetAwaiter().GetResult();

                        // Đảm bảo phản hồi thành công
                        response.EnsureSuccessStatusCode();
                    }
                    catch (Exception ex)
                    {
                        var aa = ex.Message;
                        return;
                    }


                    // Đọc nội dung phản hồi đồng bộ
                    string responseBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    var rootObject = JsonConvert.DeserializeObject<Invoice>(responseBody);

                    // Tìm ID Cha mới nhất
                    string query = "SELECT * FROM tbimport WHERE SHDon=? AND KHHDon=?";
                    var parameterss = new OleDbParameter[]
                    {
                new OleDbParameter("?", shdon),
                new OleDbParameter("?", khhdon) 
                    };
                    var kq2 = ExecuteQuery(query, parameterss);
                    string getcode = "";

                    if (kq2.Rows.Count > 0)
                    {
                        // Xử lý tên
                        bool hasVattu = false;
                        foreach (var it in rootObject.Hdhhdvu)
                        {
                            if (!hasVattu)
                            {
                                // Update nội dung cho Parent
                                query = "UPDATE tbimport SET Noidung=? WHERE ID=?";
                                var parametersss = new OleDbParameter[]
                                {
                            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(it.Ten)),
                            new OleDbParameter("?", kq2.Rows[0]["ID"]),
                                };
                                ExecuteQueryResult(query, parametersss);
                                hasVattu = true;
                            }

                            // Chèn chi tiết hóa đơn
                            query = @"
                    INSERT INTO tbimportdetail (ParentId, SoHieu, SoLuong, DonGia, DVT, Ten, MaCT, TKNo, TKCo, TTien)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

                            var parameters = new OleDbParameter[]
                            {
                        new OleDbParameter("?", kq2.Rows[0]["ID"]),
                        new OleDbParameter("?", getcode),
                        new OleDbParameter("?", it.Sluong),
                        new OleDbParameter("?", it.Dgia),
                        new OleDbParameter("?", Helpers.ConvertUnicodeToVni(it.Dvtinh)),
                        new OleDbParameter("?", Helpers.ConvertUnicodeToVni(it.Ten)),
                        new OleDbParameter("?", ""),
                        new OleDbParameter("?", kq2.Rows[0]["TKNo"]),
                        new OleDbParameter("?", kq2.Rows[0]["TKCo"]),
                        new OleDbParameter("?", it.Thtien),
                            };
                            try
                            {
                                ExecuteQueryResult(query, parameters);
                            }
                            catch (Exception ex)
                            {
                                var aa = ex.Message;
                            }

                        }
                    }
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Request error: {e.Message}");
                }
            }
        }
        public async void GetdetailXML(string nbmst, string khhdon, string shdon, string tokken)
        {
            string url = @"https://hoadondientu.gdt.gov.vn:30000/query/invoices/detail?nbmst=" + nbmst + "&khhdon=" + khhdon + "&shdon=" + shdon + "&khmshdon=1";
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                try
                {
                    // Gửi yêu cầu GET 

                    HttpResponseMessage response = await client.GetAsync(url);

                    // Đảm bảo phản hồi thành công
                    response.EnsureSuccessStatusCode();

                    // Đọc nội dung phản hồi
                    string responseBody = await response.Content.ReadAsStringAsync();
                    var rootObject = JsonConvert.DeserializeObject<Invoice>(responseBody);
                    //Tìm ID Cha mới nhất
                    string query = "SELECT *   FROM tbimport where SHDon=? and KHHDon=? and Mst= ?"; // Giả sử có cột DateAdded
                    var parameterss = new OleDbParameter[]
                      {
                            new OleDbParameter("?",shdon),
                            new OleDbParameter("?",khhdon),
                            new OleDbParameter("?",nbmst)
                      };
                    var kq2 = ExecuteQuery(query, parameterss);
                    string getcode = "";
                    if (kq2.Rows.Count > 0)
                    {
                        foreach (var it in rootObject.Hdhhdvu)
                        {
                            //Xử lý tên
                            bool hasVattu = false;

                            //Update nội dung cho Parent
                            query = @"Update tbimport set Noidung=? where ID=?";
                            var parameters = new OleDbParameter[]
                             {
                        new OleDbParameter("?",Helpers.ConvertUnicodeToVni(rootObject.Hdhhdvu.FirstOrDefault().Ten)),
                        new OleDbParameter("?", kq2.Rows[0]["ID"]),
                             };
                            int resl = ExecuteQueryResult(query, parameters);

                            query = @"
                        INSERT INTO tbimportdetail (ParentId, SoHieu, SoLuong, DonGia, DVT, Ten,MaCT,TKNo,TKCo,TTien)
                        VALUES (?, ?, ?, ?, ?, ?,?,?,?,?)";

                            parameters = new OleDbParameter[]
                            {
                        new OleDbParameter("?", kq2.Rows[0]["ID"]),
                        new OleDbParameter("?", getcode),
                        new OleDbParameter("?", it.Sluong),
                        new OleDbParameter("?", it.Dgia),
                        new OleDbParameter("?", Helpers.ConvertUnicodeToVni(it.Dvtinh)),
                        new OleDbParameter("?", Helpers.ConvertUnicodeToVni(it.Ten)),
                        new OleDbParameter("?", ""),
                        new OleDbParameter("?", kq2.Rows[0]["TKNo"]),
                        new OleDbParameter("?", kq2.Rows[0]["TKCo"]),
                        new OleDbParameter("?", it.Thtien),
                            };

                            resl = ExecuteQueryResult(query, parameters);
                        }

                    }
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Request error: {e.Message}");
                }
            }

        }
        private void btnChonthang_Click(object sender, EventArgs e)
        {
          
            Id = 1;
             filesLoaded = 0;
            totalCount = 0;
            progressPanel1.Visible = true;
            XulyFolder();
            progressBarControl1.EditValue = 0;
            if (chkDauvao.Checked)
            {
                //LoadXmlFilesOptimized(savedPath, 1);
               // LoadExcel(savedPath, 1);
                //ImportHD(people, "HDVao");
                LoadDataGridview(1);
            }
            if (chkDaura.Checked)
            {
               // LoadXmlFilesOptimized(savedPath, 2);
               // ImportHD(people2, "HDRa");
                LoadDataGridview(2);

            }
            progressPanel1.Visible = false;
            //gridView1.BestFitColumns(); // Điều chỉnh kích thước cột 
            //gridView3.BestFitColumns();
            //gridView2.BestFitColumns();
            //gridView4.BestFitColumns();
        }
        public void LoadExcel(string filePath, int type)
        {
            filePath += (type == 1 ? "\\HDVao" : "\\HDRa");
            int fromMonth = dtTungay.DateTime.Month;
            int toMonth = dtDenngay.DateTime.Month;

            var excelFiles = Directory.EnumerateFiles(filePath, "*.xlsx", SearchOption.AllDirectories)
                .Where(file => IsFileInMonthRange(file, filePath, fromMonth, toMonth))
                .ToList();

            if (excelFiles.Count == 0)
                return;

            foreach (var excelFile in excelFiles)
            {
                ProcessExcelFile(excelFile);
            }
        }
        private int GetValidRowCount(IXLWorksheet worksheet)
        {
            int rowCount = 0;
            var currentCell = worksheet.Cell("A6");

            while (!currentCell.IsEmpty())
            {
                try
                {
                    DateTime invoiceDate = DateTime.Parse(worksheet.Cell(rowCount + 7, 5).Value.ToString().Trim());
                    if (invoiceDate >= dtTungay.DateTime && invoiceDate <= dtDenngay.DateTime)
                        rowCount++;

                    currentCell = currentCell.Worksheet.Row(currentCell.Address.RowNumber + 1).Cell("A");
                }
                catch
                {
                    break;
                }
            }

            return rowCount;
        }

        private void ProcessExcelFile(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                int rowCount = GetValidRowCount(worksheet);

                if (rowCount == 0)
                    return;

                for (int i = 7; i <= (rowCount + 6); i++)
                {
                    UpdateProgressUI();

                    try
                    {
                        var fileImport = ExtractFileImportData(worksheet, i, filePath);
                        if (ShouldSkipInvoice(fileImport))
                            continue;

                        ProcessCustomerInformation(fileImport);
                        AddInvoiceToCollection(fileImport);
                        ApplyTaxCodeRules(fileImport);
                    }
                    catch (Exception ex)
                    {
                        // Log error if needed
                        continue;
                    }
                }
            }
        }
        private bool ShouldSkipInvoice(FileImport fileImport)
        {
            // Check in HoaDon table
            string query = "SELECT * FROM HoaDon WHERE KyHieu = ? AND SoHD LIKE ?";
            var parameters = new OleDbParameter[]
            {
        new OleDbParameter("KyHieu", fileImport.KHHDon),
        new OleDbParameter("SoHD", "%" + fileImport.SHDon + "%")
            };

            if (ExecuteQuery(query, parameters).Rows.Count > 0)
                return true;

            // Check in tbimport table
            query = "SELECT * FROM tbimport WHERE KHHDon = ? AND SHDon LIKE ?";
            parameters = new OleDbParameter[]
            {
        new OleDbParameter("KHHDon", fileImport.KHHDon),
        new OleDbParameter("SHDon", "%" + fileImport.SHDon + "%")
            };

            return ExecuteQuery(query, parameters).Rows.Count > 0;
        }

        private void ProcessCustomerInformation(FileImport fileImport)
        {
            string query = "SELECT TOP 1 * FROM KhachHang WHERE MST = ?";
            System.Data.DataTable result = ExecuteQuery(query, new OleDbParameter("?", fileImport.Mst));

            if (result.Rows.Count == 0)
            {
                string customerCode = GetLastFourDigits(fileImport.Mst.Replace("-", ""));
                string convertedName = Helpers.ConvertUnicodeToVni(fileImport.Ten);
                //string convertedAddress = Helpers.ConvertUnicodeToVni(fileImport.Address);
                string convertedAddress = "";

                if (string.IsNullOrEmpty(convertedAddress))
                {
                    convertedAddress = Helpers.ConvertUnicodeToVni("Bô sung địa chỉ");
                }

                // Check for duplicate customer code
                query = "SELECT TOP 1 * FROM KhachHang WHERE SoHieu = ?";
                if (ExecuteQuery(query, new OleDbParameter("?", customerCode)).Rows.Count > 0)
                {
                    customerCode = "0" + customerCode;
                }

                InitCustomer(2, customerCode, convertedName, convertedAddress, fileImport.Mst);
            }
        }

        private void AddInvoiceToCollection(FileImport fileImport)
        {
            if (!people.Any(m => m.SHDon.Contains(fileImport.SHDon) && m.KHHDon == fileImport.KHHDon))
            {
                people.Add(fileImport);
            }
        }

        private void ApplyTaxCodeRules(FileImport fileImport)
        {
            string query = "SELECT * FROM tbDinhdanhtaikhoan where KeyValue like '%MST%'";
            System.Data.DataTable taxRules = ExecuteQuery(query);

            foreach (var item in people.Where(p => p.Mst == fileImport.Mst))
            {
                foreach (DataRow row in taxRules.Rows)
                {
                    string[] conditions = row["KeyValue"].ToString().Split('&');
                    int matchedConditions = 0;

                    foreach (string condition in conditions)
                    {
                        string[] parts = Regex.Split(condition, @"([><=%]+)");
                        if (parts.Length == 3 && parts[0] == "MST" && item.Mst == parts[2])
                        {
                            matchedConditions++;
                        }
                    }

                    if (matchedConditions == conditions.Length)
                    {
                        if (string.IsNullOrEmpty(item.Noidung))
                            item.Noidung = row["Type"].ToString();

                        item.TKCo = row[4].ToString();
                    }
                }
            }
        }

        private void UpdateProgressUI()
        {
            //if (totalCount > 0)
            //    progressPercentage = (filesLoaded * 100) / totalCount;
            //else
            //    progressPercentage = 0;

            //filesLoaded++;
           // progressBarControl1.EditValue = progressPercentage;
            //progressPanel1.Caption = $"Đọc file thứ {filesLoaded - 1}/{totalCount}";
            progressPanel1.Caption = $"Đọc file XML thứ {filesLoaded + 1}/{totalCount}";
            filesLoaded += 1;
            Application.DoEvents();
        }
        private FileImport ExtractFileImportData(IXLWorksheet worksheet, int row, string filePath)
        {
            // Extract basic information
            string invoiceNumber = worksheet.Cell(row, 4).Value.ToString().Trim();
            string invoiceSeries = worksheet.Cell(row, 3).Value.ToString();
            DateTime invoiceDate = DateTime.Parse(worksheet.Cell(row, 5).Value.ToString().Trim());
            string taxCode = worksheet.Cell(row, 6).Value.ToString().Trim();

            string query = @"select * FROM KhachHang 
                                    WHERE MST LIKE ? ";
            var parameterss = new OleDbParameter[]
            {
                new OleDbParameter("?",taxCode)
               };
            var kq = ExecuteQuery(query, parameterss);
            string customerName ="";
            if (kq.Rows.Count > 0)
                customerName = Helpers.ConvertVniToUnicode(kq.Rows[0]["Ten"].ToString());
            else
                customerName = worksheet.Cell(row, 7).Value.ToString();
            string address = worksheet.Cell(row, 8).Value.ToString();
            string status = worksheet.Cell(row, 16).Value.ToString();

            // Initialize amounts
            double amountBeforeTax = 0;
            double amountAfterTax = 0;
            int taxRate = 0;
            string description = "";

            // Parse financial data
            if (!string.IsNullOrEmpty(worksheet.Cell(row, 9).Value.ToString()))
            {
                amountBeforeTax = double.Parse(worksheet.Cell(row, 9).Value.ToString().Replace(",", ""));
                amountAfterTax = double.Parse(worksheet.Cell(row, 10).Value.ToString().Replace(",", ""));

                if (amountBeforeTax < 0)
                    description = "(*) Hóa đơn điều chỉnh âm";

                taxRate = amountAfterTax > 0
                    ? int.Parse(Math.Round((amountAfterTax / amountBeforeTax * 100)).ToString())
                    : 0;
            }
            else
            {
                amountBeforeTax = double.Parse(worksheet.Cell(row, 13).Value.ToString().Replace(",", ""));
            }

            if (status.Contains("điều chỉnh"))
            {
                description = "(*) Hóa đơn điều chỉnh";
            }

            // Create and return FileImport object
            return new FileImport(
                path: filePath,
                shdon: invoiceNumber,
                khhdon: invoiceSeries,
                nlap: invoiceDate,
                ten: customerName,
                noidung: description,
                tkno: "6422",  // Default debit account
                tkco: "1111",  // Default credit account
                tkthue: 1331,  // Default tax account
                mst: taxCode,
                tongTien: amountBeforeTax,
                vat: taxRate,
                type: 1,       // Assuming type 2 for these invoices
                tenTP: "",     // Empty for now
                isacess: true,
                tPhi: 0,
                tgTCThue: amountBeforeTax,
                tgTThue: amountAfterTax,
                true
            );
        }

        public static ChromeDriver Driver { get; private set; }
        #endregion

        #region Xu ly co quan thue
        int DoTask = 0;
        int Endtask = 0;
        private void btnTaicoquanthue_Click(object sender, EventArgs e)
        {
            //frmTaiCoQuanThue frmTaiCoQuanThue = new frmTaiCoQuanThue();
            //username = txtuser.Text;
            //pasword = txtpass.Text;
            //type = chkDauvao.Checked ? 1 : 2;
            //dtFrom = dtTungay.DateTime;
            //dtTo = dtDenngay.DateTime;
            //frmTaiCoQuanThue.frmMain = this;
            //frmTaiCoQuanThue.Show();
            int type = 0;
            if (chkDauvao.Checked && !chkDaura.Checked)
            {
                type = 1;
            }
            if (!chkDauvao.Checked && chkDaura.Checked)
            {
                type = 2;
            }
            if (chkDauvao.Checked && chkDaura.Checked)
            {
                type = 3;
            }
            Taihoadon(type);
        }
        private int Sohoadoncuathan = 0;
        private void Taihoadon(int type)
        {
            //Xử lý vòng lặp ngày
            Dictionary<DateTime, DateTime> lstDicDate = new Dictionary<DateTime, DateTime>();
            //Nếu trong 1 tháng
            if (dtTungay.DateTime.Month == dtDenngay.DateTime.Month)
            {
                lstDicDate.Add(dtTungay.DateTime, dtDenngay.DateTime);
            }
            //Nếu khác tháng
            else
            {
                int j = 0;
                for (int i = dtTungay.DateTime.Month; i <= dtDenngay.DateTime.Month; i++)
                {
                    int lastDay = DateTime.DaysInMonth(dtTungay.DateTime.Year, i);
                    // Tạo ngày cuối cùng của tháng
                    DateTime lastDateOfMonth = new DateTime(dtTungay.DateTime.Year, i, lastDay);
                    //Nếu lần đầu, từ ngày cặp đau tien lấy theo tu ngày
                    if (j == 0)
                        lstDicDate.Add(dtTungay.DateTime, lastDateOfMonth);
                    //
                    else
                    {
                        if(dtDenngay.DateTime< lastDateOfMonth)
                        lstDicDate.Add(new DateTime(dtTungay.DateTime.Year, i, 1), dtDenngay.DateTime);
                        else
                            lstDicDate.Add(new DateTime(dtTungay.DateTime.Year, i, 1), lastDateOfMonth);
                    }
                    j++;
                }
            }
            Driver = null;
            if (Driver == null)
            {
                var options = new ChromeOptions();
                // Tắt các cảnh báo bảo mật (Safe Browsing)

                // Tắt Safe Browsing và các tính năng bảo mật can thiệp
                options.AddArgument("--disable-features=SafeBrowsing,DownloadBubble,DownloadNotification");
                options.AddArgument("--safebrowsing-disable-extension-blacklist");
                options.AddArgument("--safebrowsing-disable-download-protection");

                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("safebrowsing.enabled", false);
                options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
                // Tối ưu hóa trình duyệt

                options.AddArguments(
                    "--disable-notifications",
                    "--start-maximized",
                    "--disable-extensions",
                    "--disable-infobars");
                //
                string downloadPath = "";
                if (type == 1)
                    downloadPath = savedPath + "\\HDVao";
                if (type == 2)
                    downloadPath = savedPath + "\\HDRa";
                options.AddUserProfilePreference("download.default_directory", downloadPath);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("disable-popup-blocking", "true");
                options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
                options.AddUserProfilePreference("safebrowsing.enabled", false); // Tắt Safe Browsing hoàn toàn
                var driverPath = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                ChromeDriverService chromeService = ChromeDriverService.CreateDefaultService(driverPath);
                chromeService.HideCommandPromptWindow = true; // Để ẩn cửa sổ CMD của driver


                Driver = new ChromeDriver(chromeService, options); 
                //
                try
                {
                    Driver.Navigate().GoToUrl("https://hoadondientu.gdt.gov.vn");
                    IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
                    js.ExecuteScript("window.scrollTo(0, 0);");
                    Thread.Sleep(1000);
                    var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(100));
                    var closeButton = wait.Until(driver => driver.FindElement(By.XPath("//span[@class='ant-modal-close-x']")));
                    closeButton.Click();
                    //
                    var loginButton = wait.Until(driver => driver.FindElement(By.XPath("//div[@class='ant-col home-header-menu-item']/span[text()='Đăng nhập']")));
                    loginButton.Click();
                    // Nhập tên đăng nhập, mật khẩu và captcha
                    var usernameField = Driver.FindElement(By.Id("username"));
                    var passwordField = Driver.FindElement(By.Id("password"));
                    //usernameField.SendKeys("3502501171"); // Thay your_username bằng tên đăng nhập thực tế
                    //passwordField.SendKeys("PDVT12345678aA@");
                    string username = txtuser.Text;
                    string password = txtpass.Text;
                    usernameField.SendKeys(username); // Thay your_username bằng tên đăng nhập thực tế
                    passwordField.SendKeys(password);
                    new Actions(Driver)
    .KeyDown(Keys.Tab).KeyUp(Keys.Tab)  // Tab lần 1
    .Pause(TimeSpan.FromMilliseconds(100))  // Đợi ngắn
    .KeyDown(Keys.Tab).KeyUp(Keys.Tab)  // Tab lần 2
    .Perform();

                    //Tìm capcha
               
                    var cvalue = Driver.FindElements(By.Id("cvalue"));

                    var imgElement = Driver.FindElements(By.XPath("//img[contains(@src, 'data:image')]"));

                    // In ra src của thẻ img
                    try
                    {
                        string src = imgElement[1].GetAttribute("src");

                        Testimg2(src);
                        Thread.Sleep(200);
                        string recap = Readcapcha();
                        cvalue[1].SendKeys(recap);
                        Thread.Sleep(200);
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    loginButton = Driver.FindElement(By.XPath("//button[contains(span/text(), 'Đăng nhập')]"));
                    loginButton.Click();
                    wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(200));
                    //chờ khi nao dang nhap xong
                    //                var button = wait.Until(d =>
                    //d.FindElement(By.CssSelector("button.ant-btn-icon-only i[aria-label='icon: user']"))
                    // .FindElement(By.XPath("./parent::button")));
                    wait.Until(d =>
                    d.FindElements(By.XPath("//div[contains(@class,'home-header-menu-item')]//span[text()='Đăng nhập']")).Count == 0);
                    //DoTask = int.Parse(comboBoxEdit1.SelectedItem.ToString());
                    //Endtask = int.Parse(comboBoxEdit2.SelectedItem.ToString());
                    DoTask = dtTungay.DateTime.Month;
                    Endtask = dtDenngay.DateTime.Month;
                   
                    if (type == 1)
                    {
                        foreach (var item in lstDicDate)
                        {
                            Xulysaudangnhap(DateTime.Parse(item.Key.ToString()), DateTime.Parse(item.Value.ToString()));
                        }
                    }
                    if (type == 2)
                    {
                        foreach(var item in lstDicDate)
                        {
                            Xulysaudangnhap2(DateTime.Parse(item.Key.ToString()),DateTime.Parse(item.Value.ToString()));
                        }
                    }
                    Driver.Quit();
                }
                catch (Exception ex)
                {
                    Driver.Close();
                    Environment.Exit(0);    
                    // MessageBox.Show($"Lỗi: {ex.Message}");
                }
            }
        }
        int globaltype = 0;
        Dictionary<int, int> dictionMonth = new Dictionary<int, int>();
        private int soThangtai=0;
        private void Xulychonngay(WebDriverWait wait, int type, DateTime fd, DateTime td)
        {

            // Tìm input với class 'ant-calendar-input' và placeholder 'Chọn thời điểm'
            var allInputs = Driver.FindElements(By.CssSelector("input.ant-calendar-picker-input"));
            Thread.Sleep(100);
            if (type == 1)
                allInputs[2].Click();
            if (type == 2)
                allInputs[0].Click();
            IWebElement monthSelect = Driver.FindElement(By.CssSelector("a.ant-calendar-month-select[title='Chọn tháng']"));
            monthSelect.Click();

            IWebElement monthItem = Driver.FindElement(By.XPath("//a[contains(@class,'ant-calendar-month-panel-month') and text()='Thg 0" + fd.Month.ToString() + "']"));
            monthItem.Click();

            var elements = Driver.FindElements(By.CssSelector("div.ant-calendar-date"));
            string xpath = $"//td[@role='gridcell' and @title='{fd.Day} tháng {fd.Month} năm 2025']";
            var tdElement = Driver.FindElement(By.XPath(xpath));

            // Cuộn tới phần tử
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", tdElement);

            // Đợi cho phần tử có thể click được
            wait.Until(d => d.FindElement(By.XPath(xpath)).Displayed && d.FindElement(By.XPath(xpath)).Enabled);
            tdElement.Click();
            Thread.Sleep(2000);
            // Lấy ngày cuối tháng
            DateTime selectedDate = dtTungay.DateTime;
            DateTime lastDay = new DateTime(selectedDate.Year, selectedDate.Month, DateTime.DaysInMonth(selectedDate.Year, selectedDate.Month));
            if (type == 1)
                allInputs[3].Click();
            if (type == 2)
                allInputs[1].Click();
            monthSelect = Driver.FindElement(By.CssSelector("a.ant-calendar-month-select[title='Chọn tháng']"));
            monthSelect.Click();

            monthItem = Driver.FindElement(By.XPath("//a[contains(@class,'ant-calendar-month-panel-month') and text()='Thg 0" + td.Month.ToString() + "']"));
            monthItem.Click();

            elements = Driver.FindElements(By.CssSelector("div.ant-calendar-date"));
            int day = td.Day;
            int month = td.Month;

            // Tạo XPath động dựa trên ngày và tháng
            xpath = $"//td[@role='gridcell' and @title='{day} tháng {month} năm 2025']";
            tdElement = Driver.FindElement(By.XPath(xpath));

            // Cuộn tới phần tử
            js = (IJavaScriptExecutor)Driver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", tdElement);

            // Đợi cho phần tử có thể click được
            wait.Until(d => d.FindElement(By.XPath(xpath)).Displayed && d.FindElement(By.XPath(xpath)).Enabled);
            tdElement.Click();
        }
        private int oldRow = 1;
        private void Xulysaudangnhap2(DateTime fromdate,DateTime todate)
        {
            Sohoadoncuathan = 0;

            Thread.Sleep(300);

            if (Driver == null)
            {
                MessageBox.Show("Vui lòng mở trình duyệt trước!");
                return;
            }

            Thread.Sleep(1000);
            var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(200));
            string targetUrl = "https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don";
            Driver.Navigate().GoToUrl(targetUrl);
            Thread.Sleep(1000);
            Xulychonngay(wait,2, fromdate, todate);
            Thread.Sleep(2000);
            new Actions(Driver)
                .SendKeys(Keys.Enter) // Tab lần 2
                .Perform();

            var button = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn') and .//span[text()='Tìm kiếm']])[1]")));
            while (TryClick(button)) ;
            waitLoading(wait);
            IReadOnlyCollection<IWebElement> rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
            int rowCount = rows.Count;
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(200));

            var divElement = wait.Until(d => d.FindElements(By.XPath("//div[@class='ant-select-selection-selected-value' and @title='15']")));

            // Kiểm tra nếu phần tử được tìm thấy và nhấp vào nó
            if (divElement != null && divElement[0].Displayed)
            {
                divElement[0].Click();
                waitLoading(wait);
            }
            Thread.Sleep(1000);
            var dropdownMenu = wait.Until(d => d.FindElement(By.ClassName("ant-select-dropdown-menu")));
            var option50 = wait.Until(d => dropdownMenu.FindElements(By.XPath(".//li[text()='50']")));

            // Nhấp vào phần tử "50"
            if (option50 != null)
            { 
                while (TryClick(option50[0])) ;
            }

            waitLoading(wait); 
            bool isPhantrang = false;
            Thread.Sleep(1000);
            while (!isPhantrang)
            {
                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(200));
                bool hasrow = false;

                try
                {
                    hasrow = true;
                }
                catch (Exception ex)
                {
                    hasrow = false;
                }

                if (hasrow)
                {
                    //rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
                    //var clickableRows = rows.Where(row =>
                    //{
                    //    try
                    //    {
                    //        return row.Displayed && row.Enabled && row.FindElements(By.CssSelector("td")).Any(td => td.Displayed);
                    //    }
                    //    catch
                    //    {
                    //        return false;
                    //    }
                    //}).ToList();

                   // rowCount = clickableRows.Count;
                    int currentRow = 1;
                    bool hasMoreRows = true;
                    List<string> lstHas = new List<string>();
                    int hasdata = 0;
                    rowCount = 100;
                    bool isnext = true;

                  

                    while (isnext)
                    {
                        try
                        {
                            waitLoading(wait);
                            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(3));
                            var row = wait.Until(d =>
                            {
                                try
                                {
                                    return d.FindElement(By.XPath($"(//tbody[@class='ant-table-tbody']/tr[contains(@class,'ant-table-row')])[{currentRow}]"));
                                }
                                catch (NoSuchElementException)
                                {
                                    return null; // Trả về null nếu không tìm thấy
                                }
                            });

                            var cellC25TYY = row.FindElement(By.XPath("./td[3]/span")).Text; // C25TYY
                            var cell22252 = row.FindElement(By.XPath("./td[4]")).Text; // 22252
                            oldRow += 1;
                            // Click vào dòng
                            while (TryClick(row)) ;
                            button = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[13]")));
                            while (TryClick(button)) ;
                            waitLoading(wait);

                            string fp = currentRow == 1
                                ? savedPath + "\\HDRa\\" + "invoice.zip"
                                //: savedPath + "\\HDRa\\" + "invoice (" + (currentRow - 1 - hasdata) + ").zip";

                                : savedPath + "\\HDRa\\" + "invoice.zip";
                            try
                            {
                                wait.Until(d => File.Exists(fp));
                                lstHas.Add(fp);
                                GiaiNenhoadon(2);
                                Sohoadoncuathan += 1;
                            }
                            catch (Exception ex) { }

                            currentRow++; // Chuyển sang dòng tiếp theo
                        }
                        catch (NoSuchElementException)
                        {
                            hasMoreRows = false; // Không còn dòng nào nữa
                            Console.WriteLine($"Đã xử lý hết {currentRow - 1} dòng");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Lỗi khi xử lý dòng {currentRow}: {ex.Message}");
                            isnext = false;
                            currentRow++; // Vẫn tiếp tục với dòng tiếp theo
                        }
                    }

                    if (lstHas.Count > 0)
                    {
                       
                    }

                    var buttonElement = Driver.FindElements(By.ClassName("ant-btn-primary"));
                    bool isDisabled = !buttonElement[3].Enabled;

                    if (!isDisabled)
                    {
                        Thread.Sleep(1000); 
                        while (TryClick(buttonElement[3])) ;
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        isPhantrang = true;
                        Thread.Sleep(1000);
                    }
                }
            }
            Thread.Sleep(1000);
            Xulymaytinhtien2(wait);
            dictionMonth.Add(DoTask, Sohoadoncuathan);
            //DoTask += 1;
            //Xulysaudangnhap2(); 
            this.Focus();
        }
        private void waitLoading(WebDriverWait wait)
        {
            var spinWrapper = wait.Until(d =>
            {
                var elements = d.FindElements(By.CssSelector(".enNzTo"));
                return elements.Count > 0 ? elements[0] : null; // Trả về phần tử nếu tìm thấy
            });
            // Lấy giá trị của thuộc tính 'style'
            wait.Until(d =>
            {
                string displayValue = (string)((IJavaScriptExecutor)Driver).ExecuteScript("return window.getComputedStyle(arguments[0]).display;", spinWrapper);
                return displayValue == "none";
            });
        }
        public bool TryClick(IWebElement element)
        {
            try
            {
                element.Click();
                return false;
            }
            catch (Exception ex)
            {
                Thread.Sleep(300);
                return true;
            }
        }
        private void Xulysaudangnhap(DateTime fromdate, DateTime todate)
        {
            Sohoadoncuathan = 0;
            //if (DoTask > Endtask)
            //{
            //    Driver.Quit(); // Đóng WebDriver 
            //    StringBuilder sb = new StringBuilder();
            //    this.Focus(); // Đặt focus cho form
            //    return;

            //}
            Thread.Sleep(200);
            if (Driver == null)
            {
                MessageBox.Show("Vui lòng mở trình duyệt trước!");
                return;
            }

            try
            {
                var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
                var notificationButton = wait.Until(d => d.FindElement(By.XPath("//i[@aria-label='icon: bell']/parent::button")));

                string targetUrl = "https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don";
                // Cách 1: Chuyển trang đơn giản
                Driver.Navigate().GoToUrl(targetUrl);
                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
                var tab = wait.Until(d => d.FindElement(
                    By.XPath("//div[@role='tab' and .//span[contains(text(),'Tra cứu hóa đơn điện tử mua vào')]]")
                ));
                Thread.Sleep(200);
                tab.Click();

                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
                Thread.Sleep(200);
                Xulychonngay(wait, 1, fromdate, todate);
                Thread.Sleep(200);
                new Actions(Driver)
.SendKeys(Keys.Enter) // Tab lần 2
.Perform();
                var button = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn') and .//span[text()='Tìm kiếm']])[2]")));
                while (TryClick(button)) ;
               // button.Click();

                //Chờ loading ẩn
                waitLoading(wait);
                //wait.Until(d => d.FindElements(By.CssSelector("tr.ant-table-row")).Count > 0);
                //IReadOnlyCollection<IWebElement> rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
                //int rowCount = rows.Count;
                //Chọn 50 rows
                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(200));


                var divElement = wait.Until(d => d.FindElements(By.XPath("//div[@class='ant-select-selection-selected-value' and @title='15']")));
                // Nếu tìm thấy ít nhất một phần tử
                if (divElement[1] != null)
                {
                    // Thực hiện cuộn đến phần tử với thời gian
                    var jsExecutor = (IJavaScriptExecutor)Driver;
                    int scrollDuration = 1000; // Thời gian cuộn (ms)
                    int scrollStep = 20; // Bước cuộn (px)

                    for (int i = 0; i < scrollDuration; i += scrollStep)
                    {
                        jsExecutor.ExecuteScript("window.scrollBy(0, arguments[0]);", scrollStep);
                        Thread.Sleep(scrollStep); // Thời gian nghỉ giữa các lần cuộn
                    }

                    // Cuộn đến phần tử cuối cùng
                    jsExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", divElement[1]);
                }
                else
                {
                    Console.WriteLine("Không tìm thấy phần tử");
                }
                // Kiểm tra nếu phần tử được tìm thấy và nhấp vào nó
                if (divElement != null && divElement[1].Displayed)
                {
                    divElement[1].Click();
                    Console.WriteLine("Đã nhấp vào phần tử.");
                }
                Thread.Sleep(500);
                var dropdownMenu = wait.Until(d => d.FindElement(By.ClassName("ant-select-dropdown-menu")));

                // Tìm phần tử <li> có nội dung là "50" và nhấp vào nó
                var option50 = wait.Until(d => dropdownMenu.FindElements(By.XPath(".//li[text()='50']")));
                while (TryClick(option50[0])) ;
                // Nhấp vào phần tử "50"

                //Click download XML
                waitLoading(wait);
                Thread.Sleep(1000);
                bool isPhantrang = false;

                try
                {
                    while (isPhantrang == false)
                    {
                       
                        int currentRow = 1;
                        bool hasMoreRows = true;
                        List<string> lstHas = new List<string>();
                        int hasdata = 0;
                        bool isnext = true;
                        //while ((currentRow) <= rowCount)
                        while (isnext)
                        {
                            try
                            {
                                // wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                               // var checkerror= wait.Until(d=>d.FindElement(By.ClassName("ant-notification-notice-message")));
                                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(5));
                                // Tìm dòng hiện tại
                                var row = wait.Until(d =>
                                {
                                    try
                                    {
                                        return d.FindElement(By.XPath($"(//tbody[@class='ant-table-tbody']/tr[contains(@class,'ant-table-row')])[{currentRow}]"));
                                    }
                                    catch (NoSuchElementException)
                                    {

                                        return null; // Trả về null nếu không tìm thấy
                                    }
                                });
                                var cellC25TYY = row.FindElement(By.XPath("./td[3]/span")).Text; // C25TYY
                                var cell22252 = row.FindElement(By.XPath("./td[4]")).Text; // 22252
                                                                                           //Kiểm tra xem  trong folder đã có chưa
                                string pathkt = savedPath + "\\HDVao\\" + dtTungay.DateTime.Month + "\\HD__" + dtTungay.DateTime.Month + "_" + cell22252 + "_" + cellC25TYY + ".xml";
                                if (File.Exists(pathkt))
                                {
                                    currentRow++;
                                    hasdata++;
                                    Thread.Sleep(200);
                                    continue;
                                }
                                cell22252 = Helpers.InsertZero(cell22252);
                                pathkt = savedPath + "\\HDVao\\" + dtTungay.DateTime.Month + "\\HD__" + dtTungay.DateTime.Month + "_" + cell22252 + "_" + cellC25TYY + ".xml";
                                if (File.Exists(pathkt))
                                {
                                    currentRow++;
                                    hasdata++;
                                    Thread.Sleep(200);
                                    continue;
                                }
                                // var a=
                                string query = "SELECT * FROM HoaDon WHERE KyHieu = ? AND SoHD LIKE ?";

                                // Click vào dòng
                                while (TryClick(row)) ;
                                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(15));
                                Thread.Sleep(200);
                                button = wait.Until(d =>
                                 d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[19]")));
                                while (TryClick(button)) ;
                                // Xử lý sau khi click (đợi tải, đóng popup,...)
                                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(20));
                                waitLoading(wait);
                                string fp = "";
                                if (currentRow == 1)
                                    fp = savedPath + "\\HDVao\\" + "invoice.zip";
                                else
                                    // fp = savedPath + "\\HDVao\\" + "invoice (" + (currentRow - 1 - hasdata) + ").zip";
                                    fp = savedPath + "\\HDVao\\" + "invoice.zip";
                                try
                                {
                                    wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(2));
                                    wait.Until(d => File.Exists(fp));
                                    lstHas.Add(fp);
                                    GiaiNenhoadon(1);
                                }
                                catch (Exception ex)
                                {

                                }
                                currentRow++; // Chuyển sang dòng tiếp theo
                            }
                            catch (NoSuchElementException)
                            {
                                hasMoreRows = false; // Không còn dòng nào nữa


                                Console.WriteLine($"Đã xử lý hết {currentRow - 1} dòng");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Lỗi khi xử lý dòng {currentRow}: {ex.Message}");
                                isnext = false;
                                currentRow++; // Vẫn tiếp tục với dòng tiếp theo
                            }
                        }
                        if (lstHas.Count > 0)
                        {
                            //var getlastlist = lstHas.LastOrDefault();
                            //wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
                            //wait.Until(d => File.Exists(getlastlist));
                            //GiaiNenhoadon(1);
                        }
                        //Xử lý phần trang 
                        var buttonElement = Driver.FindElements(By.ClassName("ant-btn-primary"));

                        // Kiểm tra xem button có bị vô hiệu hóa không
                        bool isDisabled = !buttonElement[7].Enabled;
                        if (isDisabled == false)
                        {
                            while (TryClick(buttonElement[7])) ;
                            Thread.Sleep(1000);
                        }
                        else
                        {
                            isPhantrang = true;
                        }
                    }

                }
                catch (Exception ex)
                {
                    throw ex;
                }
                Thread.Sleep(500);
                Cucthuekhngnhanma(wait, fromdate, todate);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}");
            }
        }
        public void  GiaiNenhoadon(int type) // Hoặc private async Task nếu hàm này chỉ dùng nội bộ
        {
            try
            {
                string typepath = "";
                if (type == 1)
                    typepath = "\\HDVao";
                if (type == 2)
                    typepath = "\\HDRa";
                string pah = savedPath + typepath;

                // Kiểm tra xem thư mục có tồn tại không trước khi lấy danh sách file
                if (!Directory.Exists(pah))
                {
                    Console.WriteLine($"Thư mục '{pah}' không tồn tại để giải nén.");
                    return; // Thoát nếu thư mục không tồn tại
                }

                // Directory.GetFiles là đồng bộ, nhưng việc này thường nhanh
                string[] zipFiles = Directory.GetFiles(pah, "*.zip");
                var i = 0;

                // List để theo dõi các tác vụ giải nén
                var extractionTasks = new System.Collections.Generic.List<Task>();

                foreach (var zipFile in zipFiles)
                {
                    string filename = "";
                    // Lưu ý: Logic đặt tên file này có thể cần được xem xét lại nếu
                    // tên file tải về không theo format invoice.zip, invoice (1).zip
                    // Tốt hơn là lấy tên file từ Path.GetFileName(zipFile)
                    if (i == 0)
                    {
                        filename = Path.GetFileName(zipFile); // Lấy tên file thực tế
                                                              // Nếu bạn vẫn muốn cố định tên là "invoice.zip" cho file đầu tiên,
                                                              // thì cần đảm bảo file đó thực sự được tải về với tên đó.
                                                              // filename = "invoice.zip"; 
                    }
                    else
                    {
                        filename = Path.GetFileName(zipFile); // Lấy tên file thực tế
                                                              // Nếu bạn vẫn muốn cố định tên là "invoice (i).zip",
                                                              // thì cần đảm bảo file đó thực sự được tải về với tên đó.
                                                              // filename = "invoice (" + i + ").zip"; 
                    }

                    // Gọi ExtractZip và chờ nó hoàn thành từng cái một (dù ExtractZip chạy trên thread pool)
                    // Nếu bạn muốn giải nén CÙNG LÚC nhiều file, bạn có thể thay đổi cách gọi ở đây.
                    // Để chạy tuần tự từng file một (nhưng bản thân mỗi file giải nén là async):
                     ExtractZip(pah, filename);

                    // Nếu bạn muốn giải nén SONG SONG các file:
                    // extractionTasks.Add(ExtractZip(pah, filename));

                    i++;
                }

                // Nếu bạn chọn giải nén SONG SONG, bạn cần chờ tất cả các tác vụ hoàn thành ở đây:
                // await Task.WhenAll(extractionTasks);
                // Console.WriteLine("Đã hoàn tất giải nén tất cả các file zip.");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi trong GiaiNenhoadon: {ex.Message}");
                // Không nên 'throw ex;' ở đây nếu đây là một event handler của UI hoặc không có handler cụ thể.
                // Thay vào đó, hãy ghi log lỗi.
                // Nếu bạn muốn ném lại ngoại lệ, hãy dùng 'throw;' (không phải 'throw ex;') để giữ nguyên stack trace.
                // throw;
            }
        }
        private void  ExtractZip(string path, string filename)
        {
            string downloadPath = path;
            string zipFileName = filename;

            string zipFilePath = System.IO.Path.Combine(downloadPath, zipFileName);
            string extractPath = System.IO.Path.Combine(downloadPath, System.IO.Path.GetFileNameWithoutExtension(zipFileName));

            try
            {
                Directory.CreateDirectory(extractPath);
                ZipFile.ExtractToDirectory(zipFilePath, extractPath);
                Console.WriteLine($"Đã giải nén thành công vào: {extractPath}");
                File.Delete(zipFilePath);

                string invoiceFilePath = System.IO.Path.Combine(extractPath, "invoice.xml");
                string htmlFilepath = System.IO.Path.Combine(extractPath, "invoice.html");

                var newname = Getnewname(invoiceFilePath, 1);
                var newnamehtml = Getnewname(invoiceFilePath, 2);

                var getsplit = (newname.Split('_'))[2];
                var getsplithtml = (newnamehtml.Split('_'))[2];

                string targetDirXml = Path.Combine(path, getsplit);
                string targetDirHtml = Path.Combine(path, getsplithtml);

                Directory.CreateDirectory(targetDirXml); // Tạo thư mục đích cho file XML
                Directory.CreateDirectory(targetDirHtml); // Tạo thư mục đích cho file HTML

                newname = Path.Combine(getsplit, newname); // Correctly build the path
                newnamehtml = Path.Combine(getsplithtml, newnamehtml); // Correctly build the path

                string newFilePath = Path.Combine(path, newname);
                string newFilePathhtml = Path.Combine(path, newnamehtml);

                if (File.Exists(invoiceFilePath))
                {
                    try
                    {
                        if (!File.Exists(newFilePath))
                        {
                            File.Move(invoiceFilePath, newFilePath);
                        }
                        else
                        {
                            File.Delete(invoiceFilePath);
                        }

                        if (File.Exists(htmlFilepath))
                        {
                            if (!File.Exists(newFilePathhtml))
                            {
                                File.Move(htmlFilepath, newFilePathhtml);
                            }
                            else
                            {
                                File.Delete(htmlFilepath);
                            }
                        }

                        Directory.Delete(extractPath, true);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Lỗi khi di chuyển/xóa file: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("Tệp invoice.xml không tồn tại trong thư mục giải nén.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi giải nén hoặc xử lý file: {ex.Message}");
            }
        }

        private string Getnewname(string path, int type)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(path); // Tải file XML

            // Lấy phần tử gốc
            XmlNode root = xmlDoc.DocumentElement;

            // Lấy phần tử <NDHDon>
            XmlNode ndhDonNode = root.SelectSingleNode("//NDHDon");
            XmlNode nTTChungNode = root.SelectSingleNode("//TTChung");
            XmlNode nBanNode = nTTChungNode.SelectSingleNode("Ten");
            XmlNode NgaylapNode = root.SelectSingleNode("//NLap");
            string SHDon = nTTChungNode.SelectSingleNode("SHDon")?.InnerText;
            string KHHDon = nTTChungNode.SelectSingleNode("KHHDon")?.InnerText;
            int month = 0;
            if (DateTime.TryParse(NgaylapNode.InnerText, out DateTime date))
            {
                // Lấy tháng từ DateTime
                month = date.Month;
            }
            string filename = "";
            if (type == 1)
                filename = ".xml";
            if (type == 2)
                filename = ".html";
            return "HD_" + "_" + month + "_" + SHDon + "_" + KHHDon + filename;
        }
        private void Xulymaytinhtien2(WebDriverWait wait)
        {
            var tabElement = wait.Until(d => d.FindElements(By.XPath("//div[@role='tab']"))
               .FirstOrDefault(e => e.Text.Trim() == "Hóa đơn có mã khởi tạo từ máy tính tiền"));
            if (tabElement != null)
            {
                tabElement.Click();
                Console.WriteLine("Đã nhấp vào tab.");
            }
            //var button = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn') and .//span[text()='Tìm kiếm']])[2]")));
            //button.Click();
            //
            bool isPhantrang = false;

            while (isPhantrang == false)
            {
               // wait.Until(d => d.FindElements(By.CssSelector("tr.ant-table-row")).Count > 0);
              
                //IReadOnlyCollection<IWebElement> rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
                //var clickableRows = rows.Where(row =>
                //{
                //    try
                //    {
                //        return row.Displayed && row.Enabled && row.FindElements(By.CssSelector("td")).Any(td => td.Displayed);
                //    }
                //    catch
                //    {
                //        return false;
                //    }
                //}).ToList();
                //int rowCount = clickableRows.Count;

                //Console.WriteLine($"Số dòng trong bảng: {rowCount}");



                int currentRow = 1;
                bool hasMoreRows = true;
                List<string> lstHas = new List<string>();
                int hasdata = 0;
                bool isnext = true;
                //while ((currentRow) <= rowCount)
                bool isfirst = true;
                
                while (isnext)
                {
                    try
                    {
                        wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(5));
                        // Tìm dòng hiện tại
                        var row = wait.Until(d =>
                        {
                            try
                            {
                                return d.FindElements(By.XPath($"(//tbody[@class='ant-table-tbody']/tr[contains(@class,'ant-table-row')])[{oldRow}]"));
                            }
                            catch (NoSuchElementException)
                            {

                                return null; // Trả về null nếu không tìm thấy
                            }
                        });
                        //var cellC25TYY = row[0].FindElement(By.XPath("./td[3]/span")).Text; // C25TYY
                        //var cell22252 = row[0].FindElement(By.XPath("./td[4]")).Text; // 22252
                        //if (string.IsNullOrEmpty(cellC25TYY))
                        //{
                        //    oldRow++; // Vẫn tiếp tục với dòng tiếp theo
                        //    continue;
                        //}
                        // Click vào dòng

                        while (TryClick(row[0])) ;
                        var button = wait.Until(d =>
                          d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[13]")));
                        while (TryClick(button)) ;
                        // Xử lý sau khi click (đợi tải, đóng popup,...)
                        wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(20));
                        waitLoading(wait);
                        string fp = ""; 
                        if (isfirst)
                        {
                            fp = savedPath + "\\HDRa\\" + "invoice.zip";
                            isfirst = false;
                        }
                        else
                            //fp = savedPath + "\\HDRa\\" + "invoice (" + (currentRow - 1 - hasdata) + ").zip";
                            fp = savedPath + "\\HDRa\\" + "invoice.zip";

                        wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
                        wait.Until(d => File.Exists(fp));
                        lstHas.Add(fp);
                        Sohoadoncuathan += 1;
                        GiaiNenhoadon(2);
                        currentRow++; // Chuyển sang dòng tiếp theo
                        oldRow += 1;
                    }
                    catch (NoSuchElementException)
                    {
                        hasMoreRows = false; // Không còn dòng nào nữa


                        Console.WriteLine($"Đã xử lý hết {currentRow - 1} dòng");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Lỗi khi xử lý dòng {currentRow}: {ex.Message}");
                        isnext = false;
                        currentRow++; // Vẫn tiếp tục với dòng tiếp theo
                    }
                }
                if (lstHas.Count == 0)
                    return;
                var getlastlist = lstHas.LastOrDefault();
                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
                wait.Until(d => File.Exists(getlastlist));

                var pp = savedPath + "\\HDRa";
                var pp2 = savedPath + "\\HDRa\\" + DoTask;
                // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
                string[] excelFiles = Directory.GetFiles(pp, "*.xlsx");
                if (excelFiles.Length > 0)
                {
                    string fileName = System.IO.Path.GetFileName(excelFiles[0]); // Lấy tên file
                    string destFilePath = System.IO.Path.Combine(pp2, fileName); // Tạo đường dẫn đích
                    try
                    {
                        File.Move(excelFiles[0], destFilePath);
                    }
                    catch (Exception ex)
                    {
                        File.Delete(excelFiles[0]);
                    }
                }

                // Di chuyển file

               
              //  GiaiNenhoadon(2);
                var buttonElement = Driver.FindElements(By.ClassName("ant-btn-primary"));

                // Kiểm tra xem button có bị vô hiệu hóa không
                bool isDisabled = !buttonElement[3].Enabled;
                if (isDisabled == false)
                { 
                    while (TryClick(buttonElement[3])) ;
                    Thread.Sleep(1000);
                    oldRow = 1;
                }
                else
                {
                    isPhantrang = true;
                }
            }
              
            //  LoadXmlFiles(savedPath);

            //End Xử lý hóa đơn từ máy tính tiền
        }
        private void Xulymaytinhtien(WebDriverWait wait)
        {
            //Xử lý hóa đơn từ máy tính tiền
            //By.Id("ttxly")
            Thread.Sleep(500); // Hoặc sử dụng WebDriverWait để chờ điều kiện phù hợp
            var divElement = wait.Until(d => d.FindElements(By.Id("ttxly")));
            if (divElement[1] != null)
            {
                var jsExecutor = (IJavaScriptExecutor)Driver;
                int scrollDuration = 2000; // Thời gian cuộn (ms)
                int scrollStep = 50; // Bước cuộn (px)

                for (int i = 0; i < scrollDuration; i += scrollStep)
                {
                    jsExecutor.ExecuteScript("window.scrollBy(0, arguments[0]);", scrollStep);
                    Thread.Sleep(scrollStep); // Thời gian nghỉ giữa các lần cuộn
                }

                // Cuộn đến phần tử cuối cùng
                jsExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", divElement[1]);
            }
            Thread.Sleep(500); // Hoặc sử dụng WebDriverWait để chờ điều kiện phù hợp
            divElement[1].Click();
            waitLoading(wait);
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
            var listItem = wait.Until(d => d.FindElements(By.XPath("//li[@role='option' and @class='ant-select-dropdown-menu-item']"))
            .FirstOrDefault(e => e.Text.Trim() == "Cục Thuế đã nhận hóa đơn có mã khởi tạo từ máy tính tiền"));

            // Kiểm tra nếu phần tử được tìm thấy và nhấp vào nó
            if (listItem != null)
            {
                listItem.Click();
                Console.WriteLine("Đã nhấp vào phần tử.");
            }
            else
            {
                Console.WriteLine("Không tìm thấy phần tử với văn bản cụ thể.");
            }
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
           
            var tabElement = wait.Until(d => d.FindElements(By.XPath("//div[@role='tab']"))
               .FirstOrDefault(e => e.Text.Trim() == "Hóa đơn có mã khởi tạo từ máy tính tiền"));
            if (tabElement != null)
            {
                tabElement.Click();
                Console.WriteLine("Đã nhấp vào tab.");
            }
            var button = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn') and .//span[text()='Tìm kiếm']])[2]")));
            while (TryClick(button)) ;
            //chờ loading

            waitLoading(wait);
             
            int currentRow = 1;
            bool hasMoreRows = true;
            List<string> lstHas = new List<string>();
            int hasdata = 0;
            bool isnext = true;
            //while ((currentRow) <= rowCount)
            while (isnext)
            {
                try
                {
                    wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(1));
                    // Tìm dòng hiện tại
                    var row = wait.Until(d =>
                    {
                        try
                        {
                            return d.FindElement(By.XPath($"(//tbody[@class='ant-table-tbody']/tr[contains(@class,'ant-table-row')])[{currentRow}]"));
                        }
                        catch (NoSuchElementException)
                        {

                            return null; // Trả về null nếu không tìm thấy
                        }
                    });
                    var cellC25TYY = row.FindElement(By.XPath("./td[3]/span")).Text; // C25TYY
                    var cell22252 = row.FindElement(By.XPath("./td[4]")).Text; // 22252

                    // Click vào dòng
                    while (TryClick(row)) ;
                    button = wait.Until(d =>
                     d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[19]")));
                    while (TryClick(button)) ;
                    // Xử lý sau khi click (đợi tải, đóng popup,...)
                    wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(20));
                    waitLoading(wait);
                    string fp = "";
                    if (currentRow == 15)
                    {
                        int aas = 10;
                    }
                    if (currentRow == 1)
                        fp = savedPath +"\\HDVao\\" + "invoice.zip";
                    else
                        // fp = savedPath +"\\HDVao\\" + "invoice (" + (currentRow - 1 - hasdata) + ").zip";
                        fp = savedPath + "\\HDVao\\" + "invoice.zip";
                    wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(3));
                    wait.Until(d => File.Exists(fp));
                    lstHas.Add(fp);
                    GiaiNenhoadon(1);
                    currentRow++; // Chuyển sang dòng tiếp theo
                }
                catch (NoSuchElementException)
                {
                    hasMoreRows = false; // Không còn dòng nào nữa


                    Console.WriteLine($"Đã xử lý hết {currentRow - 1} dòng");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi khi xử lý dòng {currentRow}: {ex.Message}");
                    isnext = false;
                    currentRow++; // Vẫn tiếp tục với dòng tiếp theo  
                }
            }
            if (lstHas.Count == 0)
            {
              
                DoTask += 1;
               // Xulysaudangnhap(); 
            }
            if (lstHas.Count > 0)
            {
                //var getlastlist = lstHas.LastOrDefault();
                //wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
                //wait.Until(d => File.Exists(getlastlist));
               
                DoTask += 1;
            }
          
        }
        private void XulyxoaExcel(DateTime fromdate,DateTime todate)
        {
            var pp = savedPath + "\\HDVao";
            var pp2 = savedPath + "\\HDVao\\" + fromdate.Month;
            // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
            string[] excelFiles = Directory.GetFiles(pp, "*.xlsx");
            if (excelFiles.Length > 0)
            {
                //string fileName = System.IO.Path.GetFileName(excelFiles[0]); // Lấy tên file
                string fileName = "Ex_" + fromdate.Day + "_" + todate.Day + ".xlsx";
                string destFilePath = System.IO.Path.Combine(pp2, fileName); // Tạo đường dẫn đích 
                try
                {
                    if (!File.Exists(destFilePath))
                        File.Move(excelFiles[0], destFilePath);
                    else
                         File.Delete(excelFiles[0]);
                }
                catch (Exception ex)
                { 
                    File.Delete(excelFiles[0]);
                }
            }
           
        }
        private void Cucthuekhngnhanma(WebDriverWait wait,DateTime fromdate,DateTime todate)
        {
            var divElement = wait.Until(d => d.FindElements(By.Id("ttxly")));
            if (divElement[1] != null)
            {
                var jsExecutor = (IJavaScriptExecutor)Driver;
                int scrollDuration = 2000; // Thời gian cuộn (ms)
                int scrollStep = 50; // Bước cuộn (px)

                for (int i = 0; i < scrollDuration; i += scrollStep)
                {
                    jsExecutor.ExecuteScript("window.scrollBy(0, arguments[0]);", scrollStep);
                    Thread.Sleep(scrollStep); // Thời gian nghỉ giữa các lần cuộn
                }

                // Cuộn đến phần tử cuối cùng
                jsExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", divElement[1]);
            }
            Thread.Sleep(500); // Hoặc sử dụng WebDriverWait để chờ điều kiện phù hợp
            // Nhấp vào phần tử đó

            divElement[1].Click();
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(100));
            var listItem = wait.Until(d => d.FindElements(By.XPath("//li[@role='option' and @class='ant-select-dropdown-menu-item']"))
            .FirstOrDefault(e => e.Text.Trim() == "Cục Thuế đã nhận không mã"));

            // Kiểm tra nếu phần tử được tìm thấy và nhấp vào nó
            if (listItem != null)
            {
                listItem.Click();
            }
            else
            {
                Console.WriteLine("Không tìm thấy phần tử với văn bản cụ thể.");
            }
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
            var tabElement = wait.Until(d => d.FindElements(By.XPath("//div[@role='tab']"))
              .FirstOrDefault(e => e.Text.Trim() == "Hóa đơn điện tử"));
            if (tabElement != null)
            {
                tabElement.Click();
                Console.WriteLine("Đã nhấp vào tab.");
            }
            var button = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn') and .//span[text()='Tìm kiếm']])[2]")));

            while (TryClick(button)) ;
            // 
            waitLoading(wait);
            //Thread.Sleep(2000);
            //button = wait.Until(d =>
            //    d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[18]")));
             
            //((IJavaScriptExecutor)Driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth'});", button);
            //Thread.Sleep(2000);
            ////button.Click();
            ////// Hover rồi mới click
            //new Actions(Driver)
            //    .MoveToElement(button)
            //    .Pause(TimeSpan.FromSeconds(2))
            //    .Click()
            //    .Perform();
            //Thread.Sleep(1000); 
                                // XulyxoaExcel(fromdate, todate);
            waitLoading(wait);
            //Lick 15 sang 50
              divElement = wait.Until(d => d.FindElements(By.XPath("//div[@class='ant-select-selection-selected-value' and @title='15']")));
            // Nếu tìm thấy ít nhất một phần tử
            if (divElement[0] != null)
            {
                // Thực hiện cuộn đến phần tử với thời gian
                var jsExecutor = (IJavaScriptExecutor)Driver;
                int scrollDuration = 1000; // Thời gian cuộn (ms)
                int scrollStep = 20; // Bước cuộn (px)

                for (int i = 0; i < scrollDuration; i += scrollStep)
                {
                    jsExecutor.ExecuteScript("window.scrollBy(0, arguments[0]);", scrollStep);
                    Thread.Sleep(scrollStep); // Thời gian nghỉ giữa các lần cuộn
                }

                // Cuộn đến phần tử cuối cùng
                jsExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", divElement[0]);
            }
            else
            {
                Console.WriteLine("Không tìm thấy phần tử");
            }
            // Kiểm tra nếu phần tử được tìm thấy và nhấp vào nó
            if (divElement != null && divElement[0].Displayed)
            {
                divElement[1].Click();
                Console.WriteLine("Đã nhấp vào phần tử.");
            }
            Thread.Sleep(500);
            var dropdownMenu = wait.Until(d => d.FindElement(By.ClassName("ant-select-dropdown-menu")));

            // Tìm phần tử <li> có nội dung là "50" và nhấp vào nó
            var option50 = wait.Until(d => dropdownMenu.FindElements(By.XPath(".//li[text()='50']")));
            while (TryClick(option50[0])) ;
            waitLoading(wait);
            Thread.Sleep(1000);
            bool isPhantrang = false;
            try
            {
                while (isPhantrang == false)
                {
                    int currentRow = 1;
                    bool hasMoreRows = true;
                    List<string> lstHas = new List<string>();
                    int hasdata = 0;
                    bool isnext = true;
                    while (isnext)
                    {
                        try
                        {
                            // wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                            // var checkerror= wait.Until(d=>d.FindElement(By.ClassName("ant-notification-notice-message")));
                            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(5));
                            // Tìm dòng hiện tại
                            var row = wait.Until(d =>
                            {
                                try
                                {
                                    return d.FindElement(By.XPath($"(//tbody[@class='ant-table-tbody']/tr[contains(@class,'ant-table-row')])[{currentRow}]"));
                                }
                                catch (NoSuchElementException)
                                {

                                    return null; // Trả về null nếu không tìm thấy
                                }
                            });
                            var cellC25TYY = row.FindElement(By.XPath("./td[3]/span")).Text; // C25TYY
                            var cell22252 = row.FindElement(By.XPath("./td[4]")).Text; // 22252
                                                                                       //Kiểm tra xem  trong folder đã có chưa
                            string pathkt = savedPath + "\\HDVao\\" + dtTungay.DateTime.Month + "\\HD__" + dtTungay.DateTime.Month + "_" + cell22252 + "_" + cellC25TYY + ".xml";
                            if (File.Exists(pathkt))
                            {
                                currentRow++;
                                hasdata++;
                                Thread.Sleep(200);
                                continue;
                            }
                            cell22252 = Helpers.InsertZero(cell22252);
                            pathkt = savedPath + "\\HDVao\\" + dtTungay.DateTime.Month + "\\HD__" + dtTungay.DateTime.Month + "_" + cell22252 + "_" + cellC25TYY + ".xml";
                            if (File.Exists(pathkt))
                            {
                                currentRow++;
                                hasdata++;
                                Thread.Sleep(200);
                                continue;
                            }
                            // var a=
                            string query = "SELECT * FROM HoaDon WHERE KyHieu = ? AND SoHD LIKE ?";

                            // Click vào dòng
                            while (TryClick(row)) ;
                            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(15));
                            Thread.Sleep(200);
                            button = wait.Until(d =>
                             d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[19]")));
                            while (TryClick(button)) ;
                            // Xử lý sau khi click (đợi tải, đóng popup,...)
                            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(20));
                            waitLoading(wait);
                            string fp = "";
                            if (currentRow == 1)
                                fp = savedPath + "\\HDVao\\" + "invoice.zip";
                            else
                                // fp = savedPath + "\\HDVao\\" + "invoice (" + (currentRow - 1 - hasdata) + ").zip";
                                fp = savedPath + "\\HDVao\\" + "invoice.zip";
                            try
                            {
                                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(2));
                                wait.Until(d => File.Exists(fp));
                                lstHas.Add(fp);
                                GiaiNenhoadon(1);
                            }
                            catch (Exception ex)
                            {

                            }
                            currentRow++; // Chuyển sang dòng tiếp theo
                        }
                        catch (NoSuchElementException)
                        {
                            hasMoreRows = false; // Không còn dòng nào nữa


                            Console.WriteLine($"Đã xử lý hết {currentRow - 1} dòng");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Lỗi khi xử lý dòng {currentRow}: {ex.Message}");
                            isnext = false;
                            currentRow++; // Vẫn tiếp tục với dòng tiếp theo
                        }
                    }
                    
                    //Xử lý phần trang 
                    var buttonElement = Driver.FindElements(By.ClassName("ant-btn-primary"));

                    // Kiểm tra xem button có bị vô hiệu hóa không
                    bool isDisabled = !buttonElement[7].Enabled;
                    if (isDisabled == false)
                    {
                        while (TryClick(buttonElement[7])) ;
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        isPhantrang = true;
                    }
                } 
            }
            catch(Exception ex)
            {

            }
            Xulymaytinhtien(wait);

        }

        #endregion
        #region Database Excute, query
        string mstNull = "00";
        static string ConvertToTenDigitNumber(string input)
        {
            // Tính tổng mã ASCII
            long sum = 0;
            foreach (char c in input)
            {
                sum += (int)c;
            }

            // Chuyển đổi thành chuỗi và đảm bảo có 10 chữ số
            string result = (sum % 10000000000).ToString("D10"); // Chuyển đổi thành chuỗi 10 chữ số
            return result;
        }
        static string CapitalizeFirstLetter(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input; // Kiểm tra chuỗi rỗng hoặc null

            return char.ToUpper(input[0]) + input.Substring(1);
        }
        public void InitCustomer(int Maphanloai, string Sohieu, string Ten, string Diachi, string Mst)
        {
            bool isadd = false;
            if (string.IsNullOrEmpty(Mst))
            {
                if (string.IsNullOrEmpty(Diachi))
                {
                    string qury = @"SELECT TOP 1 * FROM KhachHang AS kh
                         WHERE kh.SoHieu = ?"; // Sử dụng ? thay cho @mst trong OleDb

                    System.Data.DataTable rs = ExecuteQuery(qury, new OleDbParameter("?", "KV"));

                    if (rs.Rows.Count == 0)
                    {
                        Ten = Helpers.ConvertUnicodeToVni("Khách vãng lai");
                        Diachi = "...";
                        Mst = mstNull;
                        Sohieu = "KV";
                    }
                } 
            }
            else
            {
                //string qury = @"SELECT TOP 1 * FROM KhachHang AS kh
                //         WHERE MST = ?"; // Sử dụng ? thay cho @mst trong OleDb
                //DataTable rs = ExecuteQuery(qury, new OleDbParameter("?",Mst));
                //if (rs.Rows.Count == 0)
                //    isadd = true;
            }
            if (string.IsNullOrEmpty(Mst))
                return; 
            string query = @"
        INSERT INTO KhachHang (MaPhanLoai,SoHieu,Ten,DiaChi,MST)
        VALUES (?,?,?,?,?)";


            // Khai báo mảng tham số với đủ 10 tham số
            OleDbParameter[] parameters = new OleDbParameter[]
            {
        new OleDbParameter("?", Maphanloai),
          new OleDbParameter("?", Sohieu),
        new OleDbParameter("?", Ten),
        new OleDbParameter("?", Diachi),
        new OleDbParameter("?", Mst)
            };

            // Thực thi truy vấn và lấy kết quả
            int a = ExecuteQueryResult(query, parameters);
        }
        private System.Data.DataTable ExecuteQuery(string query, params OleDbParameter[] parameters)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    Console.WriteLine("Kết nối đến cơ sở dữ liệu thành công!");

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Thêm các tham số vào command
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }

                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                        {
                            dataAdapter.Fill(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }

            return dataTable; // Trả về DataTable chứa dữ liệu
        }
        private int ExecuteQueryResult(string query, params OleDbParameter[] parameters)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                Console.WriteLine("Kết nối đến cơ sở dữ liệu thành công!");

                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    // Thêm các tham số vào command
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }

                    int rowsAffected = command.ExecuteNonQuery(); // Thực thi câu lệnh
                    return rowsAffected;
                }
            }

            return -1;
        }
        #endregion
        #region Thuat toan
        public static string GetLastFourDigits(string input)
        {
            // Tìm vị trí của dấu '-'
            int dashIndex = input.IndexOf('-');

            // Nếu có dấu '-' trong chuỗi, lấy phần trước đó
            if (dashIndex != -1)
            {
                input = input.Substring(0, dashIndex);
            }

            // Lấy 4 ký tự cuối cùng
            if (input.Length >= 4)
            {
                return input.Substring(input.Length - 4);
            }
            else
            {
                return input; // Trả về toàn bộ chuỗi nếu độ dài nhỏ hơn 4
            }
        }
        public static string GenerateResultString(string input)
        {
            // Tìm từ đầu tiên (không cần loại bỏ dấu toàn bộ)
            string firstWord = input.Split(' ')[0];

            // Loại bỏ dấu tiếng Việt cho từ đầu tiên
            string normalizedFirstWord = RemoveVietnameseDiacritics(firstWord).Replace("á", "a");

            // Tạo 4 số ngẫu nhiên từ 1 đến 9
            string randomNumbers = GenerateRandomNumbers(4);

            // Kết hợp từ đầu tiên với 4 số ngẫu nhiên
            return normalizedFirstWord + randomNumbers;
        }
        private static string RemoveVietnameseDiacritics(string str)
        {
            // Mảng chứa ký tự có dấu
            str = str.ToLower();
            str = Regex.Replace(str, "[àáạảãâầấậẩẫăằắặẳẵ]", "a");
            str = Regex.Replace(str, "[èéẹẻẽêềếệểễ]", "e");
            str = Regex.Replace(str, "[ìíịỉĩ]", "i");
            str = Regex.Replace(str, "[òóọỏõôồốộổỗơờớợởỡ]", "o");
            str = Regex.Replace(str, "[ùúụủũưừứựửữ]", "u");
            str = Regex.Replace(str, "[ỳýỵỷỹ]", "y");
            str = Regex.Replace(str, "đ", "d");

            // Thay thế khoảng trắng bằng dấu gạch ngang
            str = Regex.Replace(str, " ", "-");
            str = str.Replace(",", "");
            str = str.Replace(".", "");

            // Thay thế tất cả các âm "o" có dấu thành "o" không dấu
            str = str.Replace("ó", "o");
            str = str.Replace("ò", "o");
            str = str.Replace("õ", "o");
            str = str.Replace("ọ", "o");
            str = str.Replace("ỏ", "o");
            str = str.Replace("ô", "o");
            str = str.Replace("ơ", "o");
            str = str.Replace("'", "");
            return str;
        }
        private static Random random = new Random(); // Tạo Random tĩnh để tái sử dụng
        private static string GenerateRandomNumbers(int length)
        {
            string randomNumbers = "";
            HashSet<int> generatedNumbers = new HashSet<int>(); // Sử dụng HashSet để lưu các số đã tạo

            while (randomNumbers.Length < length)
            {
                // Sinh số ngẫu nhiên từ 1 đến 9
                int number = random.Next(1, 10);

                // Kiểm tra nếu số đó chưa được tạo
                if (!generatedNumbers.Contains(number))
                {
                    randomNumbers += number.ToString();
                    generatedNumbers.Add(number); // Thêm số vào HashSet
                }
            }

            return randomNumbers;
        }

        private void gridControl1_DoubleClick(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = gridControl1.MainView as DevExpress.XtraGrid.Views.Grid.GridView;
            var hitInfo = gridView.CalcHitInfo(gridView.GridControl.PointToClient(MousePosition));


            // Kiểm tra nếu nhấp vào một ô
            if (hitInfo.InRowCell)
            {
                int columnIndex = hitInfo.Column.VisibleIndex; // Chỉ số cột
                if (columnIndex != 1)
                    return;
                WebBrowser webBrowser1 = new WebBrowser
                {
                    Dock = DockStyle.Fill // Đổ đầy không gian của form
                };
                // Lấy giá trị trong ô đã nhấp
                var hiddenValue = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["Path"]);
                frmWebbrowser frmCongTrinh = new frmWebbrowser();
                frmCongTrinh.Text = hiddenValue.ToString().Replace(".xml", "");
                frmCongTrinh.Show();
                frmCongTrinh.BringToFront();
                frmCongTrinh.Activate();
                // Thêm điều khiển WebBrowser vào Form
                frmCongTrinh.Controls.Add(webBrowser1);
                string filePath = hiddenValue.ToString().Replace(".xml", ".html");

                webBrowser1.Navigate("file:///" + filePath.Replace("\\", "/"));
            }
        }

        private void gridControl1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void gridControl1_Click(object sender, EventArgs e)
        {
           
        }

        public string hiddenValue { get; set; }
        public string hiddenValue2 { get; set; }
        public string hiddenValue3 { get; set; }
        private void GridcontrolKeyup(KeyEventArgs e, DevExpress.XtraGrid.Views.Grid.GridView gridView)
        {
           

            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                var selectedCells = gridView1.GetSelectedCells();

                // Kiểm tra nếu có ít nhất một ô được chọn
                if (selectedCells.Length > 0)
                {
                    // Lấy giá trị của ô đầu tiên
                    var firstCell = selectedCells[0];
                    var firstValue = gridView1.GetRowCellValue(firstCell.RowHandle, firstCell.Column.FieldName);

                    // Lặp qua tất cả các ô đã chọn
                    foreach (var cell in selectedCells)
                    {
                        // Gán giá trị của ô đầu tiên cho các ô khác
                        gridView1.SetRowCellValue(cell.RowHandle, cell.Column.FieldName, firstValue);
                    }

                    // Ngăn chặn việc xử lý sự kiện nhấn phím tiếp theo
                    e.Handled = true;
                }
                if (gridView.IsEditing)
                {
                    // Đóng trình chỉnh sửa
                    gridView.CloseEditor();
                    gridView.UpdateCurrentRow(); // Cập nhật giá trị
                }

                // Lấy chỉ số hàng hiện tại
                int currentRowHandle = gridView.FocusedRowHandle;

                // Lấy giá trị ô hiện tại
                var currentValue = gridView.GetRowCellValue(currentRowHandle, gridView.FocusedColumn);
                //if (currentValue.ToString().Contains("154") )
                //{
                  
                //    frmCongtrinh frmCongtrinh = new frmCongtrinh();
                //    frmCongtrinh.frmMain = this;
                //    frmCongtrinh.ShowDialog();
                //    if (currentValue.ToString().Contains("|"))
                //        currentValue = currentValue.ToString().Split('|')[0];
                //    if(currentValue.ToString().Contains("154"))
                //    gridView.SetRowCellValue(currentRowHandle, "TKNo", currentValue + "|" + hiddenValue);
                //    if (currentValue.ToString().Contains("511"))
                //    {
                //        gridView.SetRowCellValue(currentRowHandle, "TKCo", currentValue + "|" + hiddenValue);

                     
                //    }
                //    return;
                //}
                //if (currentValue.ToString().Contains("511"))
                //{
                //    if (!Kiemtrataikhoancon(currentValue.ToString()))
                //    {
                //        frmCongtrinh frmCongtrinh = new frmCongtrinh();
                //        frmCongtrinh.frmMain = this;
                //        frmCongtrinh.ShowDialog();
                //        if (currentValue.ToString().Contains("|"))
                //            currentValue = currentValue.ToString().Split('|')[0];
                //        if (currentValue.ToString().Contains("154"))
                //            gridView.SetRowCellValue(currentRowHandle, "TKNo", currentValue + "|" + hiddenValue);
                //        if (currentValue.ToString().Contains("511"))
                //        {
                //            gridView.SetRowCellValue(currentRowHandle, "TKCo", currentValue + "|" + hiddenValue);

                //        }
                //        return;
                //    }
                   
                //}
                // Di chuyển xuống hàng
                int nextRowHandle = currentRowHandle + 1;

                // Kiểm tra xem hàng tiếp theo có tồn tại không
                if (nextRowHandle < gridView.DataRowCount)
                {
                    // Gán giá trị cho cột trong hàng tiếp theo
                    gridView.SetRowCellValue(nextRowHandle, gridView.FocusedColumn, currentValue);
                    if (gridView.FocusedColumn == gridView.Columns["TKNo"])
                    {
                        //currentValue = gridView.GetRowCellValue(currentRowHandle, gridView.Columns["TKCo"]).ToString();
                        //gridView.SetRowCellValue(nextRowHandle, gridView.Columns["TKCo"], currentValue);
                        //currentValue = gridView.GetRowCellValue(currentRowHandle, gridView.Columns["Noidung"]).ToString();
                        //gridView.SetRowCellValue(nextRowHandle, gridView.Columns["Noidung"], currentValue);
                    }

                    // Di chuyển tiêu điểm đến ô trong hàng tiếp theo
                    gridView.FocusedRowHandle = nextRowHandle;
                    gridView.FocusedColumn = gridView.FocusedColumn; // Giữ nguyên cột
                    gridView.FocusedColumn = gridView.FocusedColumn; // Giữ nguyên cột
                    gridView.ShowEditor(); // Hiển thị editor của ô

                    // Chọn tất cả văn bản trong ô editor
                    if (gridView.ActiveEditor is DevExpress.XtraEditors.TextEdit textEdit)
                    {
                        textEdit.SelectAll(); // Chọn tất cả văn bản
                    }
                }

                e.Handled = true; // Ngăn chặn âm thanh "click" khi nhấn Enter
            }
        }
        private void gridControl1_KeyUp(object sender, KeyEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = gridControl1.MainView as DevExpress.XtraGrid.Views.Grid.GridView;

            // Kiểm tra nếu có hàng con nào đang mở
            if (gridView != null && IsAnyRowExpanded(gridView))
            {
                // Nếu có hàng con mở, xử lý cho GridView con
                HandleChildGridViewKeyUp(e, gridView);
                return; // Không xử lý cho GridView cha
            }

            // Xử lý sự kiện cho GridView cha
            GridcontrolKeyup(e, gridView);
        }

        private bool IsAnyRowExpanded(DevExpress.XtraGrid.Views.Grid.GridView gridView)
        {
            for (int rowHandle = 0; rowHandle < gridView.RowCount; rowHandle++)
            {
                // Kiểm tra xem hàng có detail view đang mở
                if (gridView.GetDetailView(rowHandle, 0) != null)
                {
                    return true; // Có hàng con đang mở
                }
            }
            return false; // Không có hàng con nào mở
        }

        private void HandleChildGridViewKeyUp(KeyEventArgs e, DevExpress.XtraGrid.Views.Grid.GridView gridView)
        {
            // Duyệt qua từng hàng để xử lý sự kiện cho GridView con
            for (int rowHandle = 0; rowHandle < gridView.RowCount; rowHandle++)
            {
                var childView = gridView.GetDetailView(rowHandle, 0) as DevExpress.XtraGrid.Views.Grid.GridView;
                if (childView != null)
                {
                    // Xử lý sự kiện KeyUp cho GridView con
                    GridcontrolKeyup(e, childView);
                }
            }
        }
        public string Kiemtracongtrinh(int id)
        {
            string querydinhdanh = @"SELECT * FROM TP154 WHERE SoHieu = ?"; // Sử dụng ? thay cho @mst trong OleDb
            var resultkm = ExecuteQuery(querydinhdanh, new OleDbParameter("?", "CTT0" + id));

            if (resultkm.Rows.Count == 0)
            {
                string query = @"INSERT INTO TP154(MaPhanLoai,SoHieu,TenVattu,DonVi,MaTK)
                         VALUES(?, ?, ?, ?,?)";
                var parameterss = new OleDbParameter[]
                {
            new OleDbParameter("?", "1"),
            new OleDbParameter("?", "CTT0"+id),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni("Công trình tạm số "+id)),
            new OleDbParameter("?","Ct"),
            new OleDbParameter("?","37")
                };
                int rr = ExecuteQueryResult(query, parameterss);
            }
            return "CTT0" + id;
        }
        private void ImportHDVao()
        {
            if (string.IsNullOrEmpty(savedPath))
            {
                XtraMessageBox.Show("Vui lòng thiết lập đường dẫn!", "Cảnh báo",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool isNull = false;
            if (people.Any(m => string.IsNullOrEmpty(m.TKCo) && m.Checked) ||
                people.Any(m => string.IsNullOrEmpty(m.TKNo) && m.Checked) ||
                people.Any(m => string.IsNullOrEmpty(m.Noidung) && m.Checked))
            {
                XtraMessageBox.Show("Thông tin không được để trống!", "Cảnh báo",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (var item in people)
            {
                if (string.IsNullOrEmpty(item.Mst))
                    item.Mst = mstNull;
                if (!item.Checked)
                {
                    continue;
                }

                if (item.TKCo == "711")
                {
                    item.TKNo = "711";
                    item.TKCo = "331";
                }

                if (string.IsNullOrEmpty(item.TKCo) ||
                    string.IsNullOrEmpty(item.TKNo) ||
                    string.IsNullOrEmpty(item.Noidung))
                {
                    isNull = true;
                    MessageBox.Show("Thông tin không được để trống");
                    break;
                }

                // Xử lý 154 cho notk
                if (item.TKNo.Contains('|'))
                {
                    var getsplits = item.TKNo.Split('|');
                    item.TKNo = getsplits[0].Trim();
                    item.SoHieuTP = getsplits[1].Trim();
                }

                if (item.Type == 3)
                    continue;

                if (item.TkThue == 0)
                {
                    if (item.TKNo.Contains("64"))
                        item.TkThue = 1331;
                    if (item.TKNo.Contains("15"))
                        item.TkThue = 1331;
                    if (item.TKNo.Contains("511"))
                        item.TkThue = 33311;
                    if (item.TkThue == 0)
                        item.TkThue = 1331;
                }

                string query = @"
            INSERT INTO tbImport (SHDon, KHHDon, NLap, Ten, Noidung, TKCo, TKNo, TkThue, Mst, Status, Ngaytao, TongTien, Vat, SohieuTP,TPhi,TgTCThue,TgTThue)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?)";

                string newTen = Helpers.ConvertUnicodeToVni(item.Ten);
                string newNoidung = Helpers.ConvertUnicodeToVni(item.Noidung);

                OleDbParameter[] parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", item.SHDon),
            new OleDbParameter("?", item.KHHDon),
            new OleDbParameter("?", item.NLap),
            new OleDbParameter("?", newTen),
            new OleDbParameter("?", newNoidung),
            new OleDbParameter("?", item.TKCo),
            new OleDbParameter("?", item.TKNo),
            new OleDbParameter("?", item.TkThue),
            new OleDbParameter("?", item.Mst),
            new OleDbParameter("?", "0"),
            new OleDbParameter("?", DateTime.Now.ToShortDateString()),
            new OleDbParameter("?", item.TongTien.ToString()),
            new OleDbParameter("?", item.Vat.ToString()),
            new OleDbParameter("?", item.SoHieuTP.ToString()),
            new OleDbParameter("?", item.TPhi.ToString()),
            new OleDbParameter("?", item.TgTCThue.ToString()),
            new OleDbParameter("?", item.TgTThue.ToString())
                };

                int a = ExecuteQueryResult(query, parameters);

                if (a > 0)
                {
                    if (item.TKNo.Contains("152") || item.TKNo.Contains("153") || item.TKNo.Contains("156") || item.TKNo.Contains("154") || item.TKNo.Contains("711"))
                    {
                        string tableName = "tbImport";
                        query = $"SELECT MAX(ID) FROM {tableName}";

                        int parentID = (int)ExecuteQuery(query, new OleDbParameter("?", null)).Rows[0][0];
                        int idconttrinh = 1;

                        foreach (var it in item.fileImportDetails)
                        {
                            if (item.TKNo == "154")
                            {
                                string sc = Helpers.ConvertUnicodeToVni("Sửa");
                                if (it.Ten.Contains(sc))
                                {
                                    it.TKNo = "154";
                                    it.MaCT = Kiemtracongtrinh(idconttrinh);
                                    idconttrinh += 1;
                                }
                                else
                                {
                                    it.TKNo = "152";
                                }
                            }
                            if (item.TKNo == "711")
                            {
                                it.TKNo = "711";
                                it.TKCo = "331";
                            }
                            query = @"
                        INSERT INTO tbimportdetail (ParentId, SoHieu, SoLuong, DonGia, DVT, Ten,MaCT,TKNo,TKCo)
                        VALUES (?, ?, ?, ?, ?, ?,?,?,?)";

                            parameters = new OleDbParameter[]
                            {
                        new OleDbParameter("?", parentID),
                        new OleDbParameter("?", it.SoHieu),
                        new OleDbParameter("?", it.Soluong),
                        new OleDbParameter("?", it.Dongia),
                        new OleDbParameter("?", Helpers.ConvertUnicodeToVni(it.DVT)),
                        new OleDbParameter("?", it.Ten),
                        new OleDbParameter("?", it.MaCT),
                        new OleDbParameter("?", it.TKNo),
                         new OleDbParameter("?", it.TKCo)
                            };

                            int resl = ExecuteQueryResult(query, parameters);
                            InsertHangHoa(Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(it.DVT)), it.SoHieu, it.Ten);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Thêm dữ liệu thất bại.");
                }

                try
                {
                    var htmlPath = Path.Combine(savedPath, "HDVao");
                    var month = "\\" + item.NLap.Month;
                    htmlPath += month;

                    var htmlFiles = Directory.EnumerateFiles(htmlPath, "*.html", SearchOption.AllDirectories);
                    foreach (var it in htmlFiles)
                    {
                        try
                            {
                                File.Move(it, it.Replace("HDVao", "HDVaoChonLoc"));
                            }
                            catch(Exception ex)
                            {

                            }
                        File.Delete(it);
                    }

                    try
                    {
                        
                        try
                        {
                            File.Move(item.Path, item.Path.Replace("HDVao", "HDVaoChonLoc"));
                        }
                        catch (Exception ex)
                        {
                                File.Delete(item.Path);
                            }
                    }
                    catch
                    {
                      
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            if (!isNull)
            {
                XtraMessageBox.Show("Lấy dữ liệu thành công!", "Thông báo",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                isClick = true;
                this.Close();
            }
        }
        private bool Kiemtrataikhoancon(string tk)
        {
            string query = @"
                        select * from  HeThongTK where SoHieu =?";
            var resultkm = ExecuteQuery(query, new OleDbParameter("?",tk));
            if (resultkm.Rows.Count > 0)
            {
                var getTK_ID2 = resultkm.Rows[0]["TK_ID2"].ToString();
                if(getTK_ID2=="0")
                    return true;
            }
            return false;
        }
        private void ImportHDRa()
        {
            if (string.IsNullOrEmpty(savedPath))
            {
                XtraMessageBox.Show("Vui lòng thiết lập đường dẫn!", "Cảnh báo",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool isNull = false;
            if (people2.Any(m => string.IsNullOrEmpty(m.TKCo) && m.Checked) ||
                people2.Any(m => string.IsNullOrEmpty(m.TKNo) && m.Checked) ||
                people2.Any(m => string.IsNullOrEmpty(m.Noidung) && m.Checked))
            {
                XtraMessageBox.Show("Thông tin không được để trống!", "Cảnh báo",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (var item in people2)
            {
                //Đảo ngược tk
                string temp = "";
                temp = item.TKCo;
                item.TKCo = item.TKNo;
                item.TKNo = temp;
                if (string.IsNullOrEmpty(item.Mst))
                    item.Mst = mstNull;
                if (!item.Checked)
                {
                    continue;
                }

                if (item.TKCo == "711")
                {
                    item.TKNo = "711";
                    item.TKCo = "331";
                }

                if (string.IsNullOrEmpty(item.TKCo) ||
                    string.IsNullOrEmpty(item.TKNo) ||
                    string.IsNullOrEmpty(item.Noidung))
                {
                    isNull = true;
                    MessageBox.Show("Thông tin không được để trống");
                    break;
                }

                // Xử lý 154 cho notk
                if (item.TKNo.Contains('|'))
                {
                    var getsplits = item.TKNo.Split('|');
                    item.TKNo = getsplits[0].Trim();
                    item.SoHieuTP = getsplits[1].Trim();
                }

                if (item.Type == 3)
                    continue;

                if (item.TkThue == 0)
                {
                    if (item.TKNo.Contains("64"))
                        item.TkThue = 1331;
                    if (item.TKNo.Contains("15"))
                        item.TkThue = 1331;
                    if (item.TKNo.Contains("511"))
                        item.TkThue = 33311;
                    if (item.TkThue == 0)
                        item.TkThue = 1331;
                }

                string query = @" INSERT INTO tbImport (SHDon, KHHDon, NLap, Ten, Noidung, TKCo, TKNo, TkThue, Mst, Status, Ngaytao, TongTien, Vat, SohieuTP,TPhi,TgTCThue,TgTThue)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?)";

                string newTen = Helpers.ConvertUnicodeToVni(item.Ten);
                string newNoidung = Helpers.ConvertUnicodeToVni(item.Noidung);

                OleDbParameter[] parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", item.SHDon),
            new OleDbParameter("?", item.KHHDon),
            new OleDbParameter("?", item.NLap),
            new OleDbParameter("?", newTen),
            new OleDbParameter("?", newNoidung),
            new OleDbParameter("?", item.TKCo),
            new OleDbParameter("?", item.TKNo),
            new OleDbParameter("?", item.TkThue),
            new OleDbParameter("?", item.Mst),
            new OleDbParameter("?", "0"),
            new OleDbParameter("?", DateTime.Now.ToShortDateString()),
            new OleDbParameter("?", item.TongTien.ToString()),
            new OleDbParameter("?", item.Vat.ToString()),
            new OleDbParameter("?", item.SoHieuTP.ToString()),
            new OleDbParameter("?", item.TPhi.ToString()),
            new OleDbParameter("?", item.TgTCThue.ToString()),
            new OleDbParameter("?", item.TgTThue.ToString())
                };


                int a = ExecuteQueryResult(query, parameters);

                if (a > 0)
                {
                    if (item.TKNo.Contains("5111"))
                    {
                        string tableName = "tbImport";
                        query = $"SELECT MAX(ID) FROM {tableName}";

                        int parentID = (int)ExecuteQuery(query, new OleDbParameter("?", null)).Rows[0][0];
                        int idconttrinh = 1;

                        foreach (var it in item.fileImportDetails)
                        {
                            if (item.TKNo == "154")
                            {
                                string sc = Helpers.ConvertUnicodeToVni("Sửa");
                                if (it.Ten.Contains(sc))
                                {
                                    it.TKNo = "154";
                                    it.MaCT = Kiemtracongtrinh(idconttrinh);
                                    idconttrinh += 1;
                                }
                                else
                                {
                                    it.TKNo = "152";
                                }
                            }
                             temp = "";
                            temp = it.TKCo;
                            it.TKCo = it.TKNo;
                            it.TKNo = temp;
                            query = @"
                        INSERT INTO tbimportdetail (ParentId, SoHieu, SoLuong, DonGia, DVT, Ten,MaCT,TKNo,TKCo)
                        VALUES (?, ?, ?, ?, ?, ?,?,?,?)";

                            parameters = new OleDbParameter[]
                            {
                        new OleDbParameter("?", parentID),
                        new OleDbParameter("?", it.SoHieu),
                        new OleDbParameter("?", it.Soluong),
                        new OleDbParameter("?", it.Dongia),
                            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(it.DVT)),
                        new OleDbParameter("?", it.Ten),
                        new OleDbParameter("?", it.MaCT),
                        new OleDbParameter("?", it.TKNo),
                         new OleDbParameter("?", it.TKCo),
                            };

                            int resl = ExecuteQueryResult(query, parameters);
                            InsertHangHoa(Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(it.DVT)), it.SoHieu,it.Ten);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Thêm dữ liệu thất bại.");
                }

                try
                {
                    var htmlPath = Path.Combine(savedPath, "HDRa");
                    var month = "\\" + item.NLap.Month;
                    htmlPath += month;

                    var htmlFiles = Directory.EnumerateFiles(htmlPath, "*.html", SearchOption.AllDirectories);
                    foreach (var it in htmlFiles)
                    {
                        File.Move(it, it.Replace("HDRa", "HDRaChonLoc"));
                    }

                    try
                    {
                        File.Move(item.Path, item.Path.Replace("HDRa", "HDRaChonLoc"));
                    }
                    catch
                    {
                        File.Delete(item.Path);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            if (!isNull)
            {
                XtraMessageBox.Show("Lấy dữ liệu thành công!", "Thông báo",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                isClick = true;
                this.Close();
            }
        }
        private void btnimport_Click(object sender, EventArgs e)
        {
            progressPanel1.Visible = true;
           
          
            if (chkDauvao.Checked)
            {

                foreach (var item in lstImportVao)
                {
                    if(string.IsNullOrEmpty(item.Noidung) && item.Checked)
                    { 
                        XtraMessageBox.Show("Hóa đơn "+ item.SHDon + ", Nội dung không được đê trống!");
                        return;
                    }
                }
            }
            if (chkDaura.Checked)
            {
                foreach (var item in lstImportRa)
                {
                    if (string.IsNullOrEmpty(item.Noidung) && item.Checked)
                    {
                        var index = lstImportRa.IndexOf(item);
                        XtraMessageBox.Show("Hóa đơn " + item.SHDon + ", Nội dung không được đê trống!");
                        return;
                    }
                }
            }
            progressPanel1.Caption = "Đang xử lý";
            progressPanel1.Visible = true;
            //Cho tất cả hóa đơn dạng ko dc check, trừ nhung hoa don =1 hoặc bằng 2
            string queryc = @"UPDATE tbimport SET  Status =-1 WHERE Status=0 ";
            int rowsAffecteds = ExecuteQueryResult(queryc, null);
            if (chkDauvao.Checked)
            {
                foreach (var item in lstImportVao)
                {
                    progressPanel1.Caption = "Đang xử lý file thứ " + (lstImportVao.IndexOf(item) + 1) + " / " + lstImportVao.Count();
                    Application.DoEvents();
                    //Thực hiện 154
                    if (item.TKNo.Contains("|"))
                    {
                        var getsplit = item.TKNo.Split('|');
                        item.TKNo = getsplit[0];
                        item.SoHieuTP = getsplit[1];
                        string query = @"UPDATE tbimport SET  TKNo =?,SohieuTP=? WHERE ID=?";
                        var parameters = new OleDbParameter[]
                         {
                              new OleDbParameter("?", item.TKNo),
                               new OleDbParameter("?", item.SoHieuTP),
                              new OleDbParameter("?", item.ID)
                         };
                        int rowsAffected = ExecuteQueryResult(query, parameters);
                        //Xóa đi con 
                         query = @"delete from  tbimportdetail WHERE ParentId=?";
                         parameters = new OleDbParameter[]
                         {
            new OleDbParameter("?", item.ID)
                         };
                         rowsAffected = ExecuteQueryResult(query, parameters);
                        item.fileImportDetails = new List<FileImportDetail>();

                    }
                    else
                    {
                        if (item.TKNo.Contains("154"))
                        {
                         var   query = @"delete from  tbimportdetail WHERE ParentId=?";
                          var  parameters = new OleDbParameter[]
                            {
            new OleDbParameter("?", item.ID)
                            };
                           var rowsAffected = ExecuteQueryResult(query, parameters);
                            item.fileImportDetails = new List<FileImportDetail>();
                        }
                    }
                    if (item.Checked)
                    {

                        string query = @"UPDATE tbimport SET  Status =0 WHERE ID=?";
                        var parameters = new OleDbParameter[]
                         {
            new OleDbParameter("?", item.ID)
                         };
                        int rowsAffected = ExecuteQueryResult(query, parameters);

                        if (!item.isHaschild)
                        {
                             query = @"delete tbimportdetail  WHERE ParentId=?";
                             parameters = new OleDbParameter[]
                             {
                              new OleDbParameter("?", item.ID)
                             };
                             rowsAffected = ExecuteQueryResult(query, parameters);
                        }
                        foreach (var it in item.fileImportDetails)
                        {
                            if (it.TKNo.Contains("|"))
                            {
                                var getsplit = it.TKNo.Split('|');
                                it.TKNo = getsplit[0];
                                it.MaCT = getsplit[1];
                                query = @"UPDATE tbimportdetail SET  TKNo =?,MaCT=? WHERE ID=?";
                                parameters = new OleDbParameter[]
                                {
                              new OleDbParameter("?", it.TKNo),
                               new OleDbParameter("?", it.MaCT),
                              new OleDbParameter("?", it.ID)
                                };
                                int rowsAffected2 = ExecuteQueryResult(query, parameters);
                            }
                            //Sửa cho trường hợp có công trình
                            if (item.TKNo.Contains("15"))
                                InsertHangHoa(Helpers.ConvertUnicodeToVni(it.DVT), it.SoHieu, Helpers.ConvertUnicodeToVni(it.Ten));
                        }
                    }
                    // Thực hiện insert Vật tư
                  

                }
            }
            else
            {
                foreach (var item in lstImportRa)
                {
                    //Thực hiện 154
                    if (item.TKCo.Contains("|"))
                    {
                        var getsplit = item.TKCo.Split('|');
                        item.TKCo = getsplit[0];
                        item.SoHieuTP = getsplit[1];
                        string query = @"UPDATE tbimport SET  TKCo =?,SohieuTP=? WHERE ID=?";
                        var parameters = new OleDbParameter[]
                         {
                              new OleDbParameter("?", item.TKCo),
                               new OleDbParameter("?", item.SoHieuTP),
                              new OleDbParameter("?", item.ID)
                         };
                        int rowsAffected = ExecuteQueryResult(query, parameters);
                        //Xóa đi con 
                        query = @"delete from  tbimportdetail WHERE ParentId=?";
                        parameters = new OleDbParameter[]
                        {
            new OleDbParameter("?", item.ID)
                        };
                        rowsAffected = ExecuteQueryResult(query, parameters);
                        item.fileImportDetails = new List<FileImportDetail>();

                    }
                    if (!item.Checked)
                    {
                        string query = @"UPDATE tbimport SET  Status =-1 WHERE ID=?";
                        var parameters = new OleDbParameter[]
                         {
            new OleDbParameter("?", item.ID)
                         };
                        int rowsAffected = ExecuteQueryResult(query, parameters);
                    }
                    else
                    {
                        string query = @"UPDATE tbimport SET  Status =0 WHERE ID=?";
                        var parameters = new OleDbParameter[]
                         {
            new OleDbParameter("?", item.ID)
                         };
                        int rowsAffected = ExecuteQueryResult(query, parameters);
                    }

                    if (!item.isHaschild)
                    {
                      var  query = @"delete from tbimportdetail  WHERE ParentId=?";
                      var  parameters = new OleDbParameter[]
                        {
                              new OleDbParameter("?", item.ID)
                        };
                    var    rowsAffected = ExecuteQueryResult(query, parameters);
                    }
                    // Thực hiện insert Vật tư
                    foreach (var it in item.fileImportDetails)
                    {
                        if (it.TKCo.Contains("|"))
                        {
                            var getsplit = it.TKCo.Split('|');
                            it.TKCo = getsplit[0];
                            it.MaCT = getsplit[1];
                            string query = @"UPDATE tbimportdetail SET  TKCo =?,MaCT=? WHERE ID=?";
                            var parameters = new OleDbParameter[]
                             {
                              new OleDbParameter("?", it.TKCo),
                               new OleDbParameter("?", it.MaCT),
                              new OleDbParameter("?", it.ID)
                             };
                            int rowsAffected = ExecuteQueryResult(query, parameters);
                        }
                        if (!it.TKCo.Contains("5113"))
                            InsertHangHoa(Helpers.ConvertUnicodeToVni(it.DVT), it.SoHieu, Helpers.ConvertUnicodeToVni(it.Ten));
                    }

                }
            }
            //Nếu isHAschild

            isClick = true;
            this.Close(); 
        }
        private void ImportHD(BindingList<FileImport> data, string type)
        {
            if (string.IsNullOrEmpty(savedPath))
            {
                XtraMessageBox.Show("Vui lòng thiết lập đường dẫn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //if (data.Any(m => string.IsNullOrEmpty(m.TKCo) && m.Checked) ||
            //    data.Any(m => string.IsNullOrEmpty(m.TKNo) && m.Checked) ||
            //    data.Any(m => string.IsNullOrEmpty(m.Noidung) && m.Checked))
            //{
            //    XtraMessageBox.Show("Thông tin không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            bool isNull = false;
            using (OleDbConnection connection = new OleDbConnection(connectionString)) // Thay thế chuỗi kết nối
            {
                connection.Open();
                using (OleDbTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        foreach (var item in data)
                        {
                            if (string.IsNullOrEmpty(item.Mst))
                            {
                                item.Mst = mstNull;
                            }
                            if (!item.Checked)
                            {
                                continue;
                            }

                            // Đảo ngược TKNo và TKCo cho HDRa
                            //if (type == "HDRa")
                            //{
                            //    string temp = item.TKCo;
                            //    item.TKCo = item.TKNo;
                            //    item.TKNo = temp;
                            //}

                            if (item.TKCo == "711")
                            {
                                item.TKNo = "711";
                                item.TKCo = "331";
                            }

                            //if (string.IsNullOrEmpty(item.TKCo) ||
                            //    string.IsNullOrEmpty(item.TKNo) ||
                            //    string.IsNullOrEmpty(item.Noidung))
                            //{
                            //    isNull = true;
                            //    MessageBox.Show("Thông tin không được để trống");
                            //    transaction.Rollback(); // Rollback giao dịch nếu có lỗi
                            //    return;
                            //}
                            bool parenthasCT = false;
                            // Xử lý 154 cho notk
                            if (item.TKNo.Contains('|'))
                            {
                                parenthasCT = true;
                                var getsplits = item.TKNo.Split('|');
                                item.TKNo = getsplits[0].Trim();
                                item.SoHieuTP = getsplits[1].Trim();
                            }

                            if (item.Type == 3)
                                continue;

                            if (item.TkThue == 0)
                            {
                                if (item.TKNo.Contains("64") || item.TKNo.Contains("15"))
                                    item.TkThue = 1331;
                                else if (item.TKNo.Contains("511"))
                                    item.TkThue = 33311;
                                else
                                    item.TkThue = 1331;
                            }

                            string query = @"
                        INSERT INTO tbImport (SHDon, KHHDon, NLap, Ten, Noidung, TKCo, TKNo, TkThue, Mst, Status, Ngaytao, TongTien, Vat, SohieuTP, TPhi, TgTCThue, TgTThue,Type,Path,IsHaschild)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?)";

                            string newTen = Helpers.ConvertUnicodeToVni(item.Ten);
                            string newNoidung = Helpers.ConvertUnicodeToVni(item.Noidung);
                             
                            OleDbParameter[] parameters = new OleDbParameter[]
                            {
                        new OleDbParameter("?", item.SHDon),
                        new OleDbParameter("?", item.KHHDon),
                        new OleDbParameter("?", item.NLap),
                        new OleDbParameter("?", newTen),
                        new OleDbParameter("?", newNoidung),
                        new OleDbParameter("?", item.TKCo),
                        new OleDbParameter("?", item.TKNo),
                        new OleDbParameter("?", item.TkThue),
                        new OleDbParameter("?", item.Mst),
                        new OleDbParameter("?", "0"),
                        new OleDbParameter("?", DateTime.Now.ToShortDateString()),
                        new OleDbParameter("?", item.TongTien.ToString()),
                        new OleDbParameter("?", item.Vat.ToString()),
                        new OleDbParameter("?", item.SoHieuTP.ToString()),
                        new OleDbParameter("?", item.TPhi.ToString()),
                        new OleDbParameter("?", item.TgTCThue.ToString()),
                        new OleDbParameter("?", item.TgTThue.ToString()),
                        new OleDbParameter("?", item.Type.ToString()),
                        new OleDbParameter("?", item.Path.ToString()),
                        new OleDbParameter("?",1)
                            };

                            int rowsAffected = ExecuteQueryResult(query, parameters);
                            string tableName = "tbImport";
                            query = $"SELECT MAX(ID) FROM {tableName}";

                            int parentID = (int)ExecuteQuery(query, new OleDbParameter("?", null)).Rows[0][0];
                            if (rowsAffected > 0 )
                            {
                                if (item.isHaschild)
                                {
                                    if (item.TKNo.Contains("152") || item.TKNo.Contains("153") || item.TKNo.Contains("156") || (item.TKNo.Contains("154") && !parenthasCT) || item.TKNo.Contains("711") || (item.TKCo.Contains("511") && !parenthasCT)) //them 5111
                                    {
                                        if (!item.TKCo.Contains("5113"))
                                        {
                                            InsertImportDetails(connection, transaction, item, parentID, type);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Thêm dữ liệu vào tbImport thất bại.");
                                transaction.Rollback();
                                return; // Dừng import nếu thêm vào tbImport thất bại
                            }
                        }

                        transaction.Commit();
                        MoveHtmlFiles(data, type); // Move files after successful import 
                        if (!isNull)
                        {
                           // XtraMessageBox.Show("Lấy dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //isClick = true; 
                        }
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show("Lỗi trong quá trình import: " + ex.Message);
                    }
                }
            }
        }
        private void MoveHtmlFiles(BindingList<FileImport> data, string type)
        {
            if (string.IsNullOrEmpty(savedPath))
            {
                return; // Không làm gì nếu đường dẫn không hợp lệ
            }

            string destFolder = Path.Combine(savedPath, type + "ChonLoc");
            if (!Directory.Exists(destFolder))
            {
                Directory.CreateDirectory(destFolder); // Tạo thư mục đích nếu nó không tồn tại
            }
            foreach (var item in data)
            {
                string htmlPath = Path.Combine(savedPath, type, item.NLap.Month.ToString()); // Đã sửa lỗi đường dẫn tháng
                string[] htmlFiles = Directory.GetFiles(htmlPath, "*.xml", SearchOption.AllDirectories);

                foreach (string file in htmlFiles)
                {
                    string destFile = file.Replace(type, type + "ChonLoc");
                    try
                    {
                        if (!File.Exists(destFile))
                            File.Move(file, destFile);
                        else
                            File.Delete(file);
                    }
                    catch (Exception ex)
                    { 
                    }
                } 
               
            }
            foreach (var item in data)
            {
                string htmlPath = Path.Combine(savedPath, type, item.NLap.Month.ToString()); // Đã sửa lỗi đường dẫn tháng
                string[] htmlFiles = Directory.GetFiles(htmlPath, "*.html", SearchOption.AllDirectories);

                foreach (string file in htmlFiles)
                {
                    string destFile = file.Replace(type, type + "ChonLoc");
                    try
                    {
                        if (!File.Exists(destFile))
                            File.Move(file, destFile);
                        else
                            File.Delete(file);
                    }
                    catch (Exception ex)
                    { 
                    }
                }
            }

            //Xóa excel
            if (type == "1")
            {
                var pp = savedPath + "\\HDVao"+dtTungay.DateTime.Month;
                var pp2 = savedPath + "\\HDVao\\" + +dtTungay.DateTime.Month;
                // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
                string[] excelFiles = Directory.GetFiles(pp, "*.xlsx");
                
            }
        }
        private void InsertImportDetails(OleDbConnection connection, OleDbTransaction transaction, FileImport item, int parentId,string type)
        {
            int idconttrinh = 1;
            foreach (var detail in item.fileImportDetails)
            { 
                string tkNo = detail.TKNo;
                if (detail.TKNo.Contains("154"))
                {
                    if (detail.TKNo.Contains("|"))
                    {
                        var getsplit = detail.TKNo.Split('|');
                        detail.TKNo = getsplit[0];
                        detail.MaCT = getsplit[1];
                    }
                    else
                    {
                        return;
                    }
                    //string sc = Helpers.ConvertUnicodeToVni("Sửa");
                    //if (detail.Ten.Contains(sc))
                    //{
                    //    tkNo = "154";
                    //    detail.MaCT = Kiemtracongtrinh(idconttrinh);
                    //    idconttrinh++;
                    //}
                    //else
                    //{
                    //    tkNo = "152";
                    //}
                }
                else if (item.TKNo == "711") // Thêm xử lý cho TK 711
                {
                    tkNo = "711";
                    detail.TKCo = "331";
                }
                if (detail.TKCo.Contains("511"))
                {
                    if (detail.TKCo.Contains("|"))
                    {
                        var getsplit = detail.TKCo.Split('|');
                        detail.TKCo = getsplit[0];
                        detail.MaCT = getsplit[1];
                    } 
                }
                //if (type == "HDRa")
                //{
                //    string temp = detail.TKCo;
                //    detail.TKCo = detail.TKNo;
                //    detail.TKNo = temp;

                //}

                string query = @"
            INSERT INTO tbimportdetail (ParentId, SoHieu, SoLuong, DonGia, DVT, Ten, MaCT, TKNo, TKCo,TTien)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?,?)";
                detail.Ten = Helpers.ConvertUnicodeToVni(detail.Ten);
                OleDbParameter[] parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", parentId),
            new OleDbParameter("?", detail.SoHieu),
            new OleDbParameter("?", detail.Soluong),
            new OleDbParameter("?", detail.Dongia),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(detail.DVT)),
            new OleDbParameter("?", detail.Ten),
            new OleDbParameter("?", detail.MaCT),
            new OleDbParameter("?", detail.TKNo),
            new OleDbParameter("?", detail.TKCo),
            new OleDbParameter("?", detail.TTien)
                };

                ExecuteQueryResult(query, parameters);
                //if (!detail.TKCo.Contains("5113"))
                //    InsertHangHoa(Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(detail.DVT)), detail.SoHieu, detail.Ten);
            }
        }
        bool isClick = false;
        private void InsertHangHoa(string DVTinh, string sohieu, string newName)
        { 
            sohieu = sohieu.ToUpper();

            if (string.IsNullOrEmpty(DVTinh) || DVTinh == "Exception" || DVTinh == "kWh")
            {
                return;
            }

            string kmSearchTerm = "%" + Helpers.ConvertUnicodeToVni("khuyến mãi").ToLower() + "%";
            string km2SearchTerm = "%" + Helpers.ConvertUnicodeToVni("khuyến mại").ToLower() + "%";
            string nhomHangTamSearchTerm = "%" + Helpers.ConvertUnicodeToVni("Nhóm hàng tạm").ToLower() + "%";

            using (OleDbConnection connection = new OleDbConnection(connectionString)) // Thay YourConnectionString
            {
                connection.Open();
                using (OleDbTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // Kiểm tra và thêm mới phân loại vật tư "Hàng khuyến mãi" nếu chưa tồn tại
                        string queryCheckPhanLoai = @"SELECT MaSo FROM PhanLoaiVattu 
                                    WHERE LCase(TenPhanLoai) LIKE ? OR LCase(TenPhanLoai) LIKE ?";
                        int maPhanLoai;
                        using (OleDbCommand checkPhanLoaiCmd = new OleDbCommand(queryCheckPhanLoai, connection, transaction))
                        {
                            checkPhanLoaiCmd.Parameters.AddWithValue("?", kmSearchTerm);
                            checkPhanLoaiCmd.Parameters.AddWithValue("?", km2SearchTerm);
                            object resultKm = checkPhanLoaiCmd.ExecuteScalar();
                            if (resultKm == null)
                            {
                                // Nếu chưa có thì thêm mới
                                string insertPhanLoaiQuery = @"
                            INSERT INTO PhanLoaiVattu (SoHieu, TenPhanLoai, Cap, MaTK)
                            VALUES (?, ?, ?, ?)";
                                using (OleDbCommand insertPhanLoaiCmd = new OleDbCommand(insertPhanLoaiQuery, connection, transaction))
                                {
                                    insertPhanLoaiCmd.Parameters.AddWithValue("?", "HKM");
                                    insertPhanLoaiCmd.Parameters.AddWithValue("?", Helpers.ConvertUnicodeToVni("Hàng khuyến mãi"));
                                    insertPhanLoaiCmd.Parameters.AddWithValue("?", 1);
                                    insertPhanLoaiCmd.Parameters.AddWithValue("?", 39);
                                    insertPhanLoaiCmd.ExecuteNonQuery();
                                }
                                maPhanLoai = GetMaPhanLoai(connection, transaction, kmSearchTerm, km2SearchTerm);
                            }
                            else
                            {
                                maPhanLoai = (int)resultKm;
                            }
                        }

                        // Kiểm tra và thêm mới vật tư
                        string queryCheckVatTu = @"SELECT MaSo FROM Vattu WHERE LCase(SoHieu) = LCase(?)";
                        using (OleDbCommand checkVatTuCmd = new OleDbCommand(queryCheckVatTu, connection, transaction))
                        {
                            checkVatTuCmd.Parameters.AddWithValue("?", sohieu.ToLower());
                            //checkVatTuCmd.Parameters.AddWithValue("?", DVTinh.ToLower());
                            object resultVatTu = checkVatTuCmd.ExecuteScalar();
                            if (resultVatTu == null)
                            {
                                // Nếu không tìm thấy, thêm mới vật tư và nhóm nếu cần
                                if (!newName.Contains(Helpers.ConvertUnicodeToVni("khuyến mãi").ToLower()) && !newName.Contains(Helpers.ConvertUnicodeToVni("khuyến mại").ToLower()))
                                {
                                    maPhanLoai = GetMaPhanLoai(connection, transaction, nhomHangTamSearchTerm);
                                    if (maPhanLoai == 0)
                                    {
                                        // Tạo nhóm tạm nếu chưa tồn tại
                                        string insertNhomHangTamQuery = @"
                                    INSERT INTO PhanLoaiVattu (SoHieu, TenPhanLoai, Cap, MaTK)
                                    VALUES (?, ?, ?, ?)";
                                        using (OleDbCommand insertNhomHangTamCmd = new OleDbCommand(insertNhomHangTamQuery, connection, transaction))
                                        {
                                            insertNhomHangTamCmd.Parameters.AddWithValue("?", "NHT");
                                            insertNhomHangTamCmd.Parameters.AddWithValue("?", Helpers.ConvertUnicodeToVni("Nhóm hàng tạm"));
                                            insertNhomHangTamCmd.Parameters.AddWithValue("?", 1);
                                            insertNhomHangTamCmd.Parameters.AddWithValue("?", 39);
                                            insertNhomHangTamCmd.ExecuteNonQuery();
                                        }
                                        maPhanLoai = GetMaPhanLoai(connection, transaction, nhomHangTamSearchTerm);
                                    }
                                }
                                string insertVatTuQuery = @"
                            INSERT INTO Vattu (MaPhanLoai, SoHieu, TenVattu, DonVi)
                            VALUES (?, ?, ?, ?)";
                                using (OleDbCommand insertVatTuCmd = new OleDbCommand(insertVatTuQuery, connection, transaction))
                                {
                                    insertVatTuCmd.Parameters.AddWithValue("?", maPhanLoai);
                                    insertVatTuCmd.Parameters.AddWithValue("?", sohieu);
                                    insertVatTuCmd.Parameters.AddWithValue("?", newName);
                                    insertVatTuCmd.Parameters.AddWithValue("?", DVTinh);
                                    insertVatTuCmd.ExecuteNonQuery();
                                }
                            }
                        }
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show(ex.Message + "  " + newName + "  so hieu: " + sohieu);
                    }
                }
            }
        }

        private int GetMaPhanLoai(OleDbConnection connection, OleDbTransaction transaction, string searchTerm1, string searchTerm2 = null)
        {
            string query = @"SELECT MaSo FROM PhanLoaiVattu WHERE LCase(TenPhanLoai) LIKE ?";
            if (searchTerm2 != null)
            {
                query += " OR LCase(TenPhanLoai) LIKE ?";
            }
            using (OleDbCommand cmd = new OleDbCommand(query, connection, transaction))
            {
                cmd.Parameters.AddWithValue("?", searchTerm1);
                if (searchTerm2 != null)
                {
                    cmd.Parameters.AddWithValue("?", searchTerm2);
                }
                object result = cmd.ExecuteScalar();
                return result != null ? (int)result : 0;
            }
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (string.IsNullOrEmpty(savedPath))
                return;
            if (isClick)
                File.WriteAllText(savedPath + "\\status.txt", "ButtonClicked");
            else
                File.WriteAllText(savedPath + "\\status.txt", "");
        }

        

        private void btnMdtk_Click(object sender, EventArgs e)
        {
            frmDinhdanh frmDinhdanh = new frmDinhdanh();
            frmDinhdanh.ShowDialog();
        }
        private void Matdinhtaikhoan(string MST, string Tensp, ref string TKNo, ref string TKCo, ref string TKThue)
        {

        }

        private void txtuser_TextChanged(object sender, EventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            config.AppSettings.Settings["username"].Value = txtuser.Text;
            config.AppSettings.Settings["password"].Value = txtpass.Text;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void panelControl3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtpass_TextChanged(object sender, EventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            config.AppSettings.Settings["username"].Value = txtuser.Text;
            config.AppSettings.Settings["password"].Value = txtpass.Text;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void txtuser_TextChanged_1(object sender, EventArgs e)
        {
            
            string query = "UPDATE tbRegister SET username = ?"; 
            // Khai báo mảng tham số với đủ 10 tham số
            OleDbParameter[] parameters = new OleDbParameter[]
            {
        new OleDbParameter("?", txtuser.Text) 
            };

            // Thực thi truy vấn và lấy kết quả
            int a = ExecuteQueryResult(query, parameters);
        }

        private void txtpass_TextChanged_1(object sender, EventArgs e)
        {
            

            string query = "UPDATE tbRegister SET [Password] = ?";
            // Khai báo mảng tham số với đủ 10 tham số
            OleDbParameter[] parameters = new OleDbParameter[]
            {
        new OleDbParameter("?", txtpass.Text)
            };

            // Thực thi truy vấn và lấy kết quả
            int a = ExecuteQueryResult(query, parameters);
        }

        private void xtraTabControl2_Click(object sender, EventArgs e)
        {
            

        }

        private void gridControl2_Click(object sender, EventArgs e)
        {
            
        }

        private void gridControl2_KeyUp(object sender, KeyEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = gridControl2.MainView as DevExpress.XtraGrid.Views.Grid.GridView;

            // Kiểm tra nếu có hàng con nào đang mở
            if (gridView != null && IsAnyRowExpanded(gridView))
            {
                // Nếu có hàng con mở, xử lý cho GridView con
                HandleChildGridViewKeyUp(e, gridView);
                return; // Không xử lý cho GridView cha
            }

            // Xử lý sự kiện cho GridView cha
            GridcontrolKeyup(e, gridView);
        }

        private void xtraTabPage1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void gridControl1_KeyDown(object sender, KeyEventArgs e)
        {
           
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        { 
            Process.Start("explorer.exe", savedPath);
          
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            

            Process.Start(dbPath.Trim());

        }

        private void gridView1_MasterRowEmpty(object sender, MasterRowEmptyEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            FileImport dt = view.GetRow(e.RowHandle) as FileImport;
            if (dt != null)
                e.IsEmpty = !lstImportVao.Any(m => m.fileImportDetails.Any(j => j.ParentId == dt.ID));
        }

        private void gridView1_MasterRowGetChildList(object sender, MasterRowGetChildListEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            FileImport dt = view.GetRow(e.RowHandle) as FileImport;
            if (dt != null)
            {
                var fileImportDetails = dt.fileImportDetails;
               // fileImportDetails.ForEach(m => m.Ten = Helpers.ConvertVniToUnicode(m.Ten));
                e.ChildList = fileImportDetails; // Gán danh sách đã sửa đổi
            }
        }

        private void gridView1_MasterRowGetRelationCount(object sender, MasterRowGetRelationCountEventArgs e)
        {
            e.RelationCount = 1;
        }

        private void gridView1_MasterRowGetRelationName(object sender, MasterRowGetRelationNameEventArgs e)
        {
            e.RelationName = "Detail";
        }
        static string RemoveLeadingSpecialCharacters(string input)
        {
            // Sử dụng LINQ để lấy các ký tự không phải là ký tự đặc biệt
            return new string(input.SkipWhile(c => !char.IsLetterOrDigit(c)).ToArray());
        }

        private void dtTungay_EditValueChanged(object sender, EventArgs e)
        {
            DateTime selectedDate = dtTungay.DateTime;
            // Lấy ngày cuối cùng của tháng
            DateTime lastDay = new DateTime(selectedDate.Year, selectedDate.Month, DateTime.DaysInMonth(selectedDate.Year, selectedDate.Month));
            dtDenngay.DateTime = lastDay;
        }
        private ToolTipController toolTipController = new ToolTipController();

        private void gridView1_CustomRowCellEdit(object sender, CustomRowCellEditEventArgs e)
        {
            // Lấy giá trị của cột 10
            object cellValue = gridView1.GetRowCellValue(e.RowHandle, gridView1.Columns["isAcess"]);
           
            // Nếu giá trị là false, vô hiệu hóa ô chỉnh sửa
            //if (cellValue is bool && !(bool)cellValue)
            //{
            //    e.RepositoryItem = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            //    e.RepositoryItem.ReadOnly = true; // Hoặc có thể sử dụng một loại điều khiển khác
            //}
        }

        private void gridView1_ShownEditor(object sender, EventArgs e)
        {
           
        }

        private void xtraTabControl2_DoubleClick(object sender, EventArgs e)
        {
         
        }

        private void gridControl2_DoubleClick(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = gridControl2.MainView as DevExpress.XtraGrid.Views.Grid.GridView;
            var hitInfo = gridView.CalcHitInfo(gridView.GridControl.PointToClient(MousePosition));


            // Kiểm tra nếu nhấp vào một ô
            if (hitInfo.InRowCell)
            {
                int columnIndex = hitInfo.Column.VisibleIndex; // Chỉ số cột
                if (columnIndex != 1)
                    return;
                WebBrowser webBrowser1 = new WebBrowser
                {
                    Dock = DockStyle.Fill // Đổ đầy không gian của form
                };
                // Lấy giá trị trong ô đã nhấp
                var hiddenValue = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["Path"]);
                frmWebbrowser frmCongTrinh = new frmWebbrowser();
                frmCongTrinh.Text = hiddenValue.ToString().Replace(".xml", "");
                frmCongTrinh.Show();
                frmCongTrinh.BringToFront();
                frmCongTrinh.Activate();
                // Thêm điều khiển WebBrowser vào Form
                frmCongTrinh.Controls.Add(webBrowser1);
                string filePath = hiddenValue.ToString().Replace(".xml", ".html");

                webBrowser1.Navigate("file:///" + filePath.Replace("\\", "/"));
            }
        }

        private void gridView3_MasterRowEmpty(object sender, MasterRowEmptyEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            FileImport dt = view.GetRow(e.RowHandle) as FileImport;
            if (dt != null)
                e.IsEmpty = !lstImportRa.Any(m => m.fileImportDetails.Any(j => j.ParentId == dt.ID));
        }

        private void gridView3_MasterRowGetChildList(object sender, MasterRowGetChildListEventArgs e)
        {

            DevExpress.XtraGrid.Views.Grid.GridView view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            FileImport dt = view.GetRow(e.RowHandle) as FileImport;
            if (dt != null)
            {
                var fileImportDetails = dt.fileImportDetails;
               // fileImportDetails.ForEach(m => m.Ten = Helpers.ConvertVniToUnicode(m.Ten));
                e.ChildList = fileImportDetails; // Gán danh sách đã sửa đổi
            }
        }

        private void gridView3_MasterRowGetRelationCount(object sender, MasterRowGetRelationCountEventArgs e)
        {
            e.RelationCount = 1;
        }

        private void gridView3_MasterRowGetRelationName(object sender, MasterRowGetRelationNameEventArgs e)
        {
            e.RelationName = "Detail";
        }

        private void gridView1_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 1) // Nếu là single click
            {
                DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;

                // Lấy tên cột đã nhấp
                string columnName = gridView.FocusedColumn.FieldName;
                if(columnName != "isHaschild")
                gridView1.ShowEditor(); // Hiển thị chế độ chỉnh sửa
                else
                {
                    var gethaschild = (bool)gridView1.GetRowCellValue(e.RowHandle, "isHaschild");
                    int id = (int)gridView1.GetRowCellValue(e.RowHandle, "ID");
                    var tkco = (string)gridView1.GetRowCellValue(e.RowHandle, "TKNo");
                    foreach (var item in lstImportVao)
                    {
                        if (item.TKNo == tkco)
                        {
                            string query = @"UPDATE tbimport SET IsHaschild=? where ID=? ";

                            var parameters = new OleDbParameter[]
                             {
                            new OleDbParameter("?", gethaschild==true?0:1),
                               new OleDbParameter("?",item.ID)
                             };
                            int rowsAffected = ExecuteQueryResult(query, parameters);
                        }
                    }


                    LoadDataGridview(1);
                }
            }
            if (e.RowHandle >= 0) // Đảm bảo là hàng hợp lệ
            {
                string columnName = gridView1.FocusedColumn.FieldName;
                if (columnName == "Index")
                {
                    bool isExpanded = gridView1.GetMasterRowExpanded(e.RowHandle);
                    gridView1.SetMasterRowExpanded(e.RowHandle, !isExpanded); // Chuyển đổi trạng thái

                }

            }
        }

        private void gridView3_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 1) // Nếu là single click
            {
                DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;

                string columnName = gridView.FocusedColumn.FieldName;
                if (columnName != "isHaschild") 
                    gridView3.ShowEditor(); // Hiển thị chế độ chỉnh sửa
                else
                {
                    var gethaschild= (bool)gridView3.GetRowCellValue(e.RowHandle, "isHaschild");
                    int id=(int)gridView3.GetRowCellValue(e.RowHandle, "ID");
                    var tkco=(string)gridView3.GetRowCellValue(e.RowHandle, "TKCo");
                    foreach (var item in lstImportRa)
                    {
                        if (item.TKCo == tkco)
                        {
                            string query = @"UPDATE tbimport SET IsHaschild=? where ID=? ";

                            var parameters = new OleDbParameter[]
                             {
                            new OleDbParameter("?", gethaschild==true?0:1),
                               new OleDbParameter("?",item.ID)
                             };
                            int rowsAffected = ExecuteQueryResult(query, parameters);
                        }
                    }

            
                    LoadDataGridview(2);
                }
            }
            if (e.RowHandle >= 0) // Đảm bảo là hàng hợp lệ
            {
                string columnName = gridView3.FocusedColumn.FieldName;
                if (columnName == "Index")
                {
                    bool isExpanded = gridView3.GetMasterRowExpanded(e.RowHandle);
                    gridView3.SetMasterRowExpanded(e.RowHandle, !isExpanded); // Chuyển đổi trạng thái

                }
                    
            }
        }

        private void gridView1_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (e.Clicks == 2) // Nhấp đúp
            {
                // Lấy thông tin hàng được nhấp
                var row = gridView1.GetRow(e.RowHandle);

                // Thực hiện hành động bạn muốn, ví dụ: mở form chỉnh sửa
                //MessageBox.Show("Nhấp đúp vào hàng: " + row.ToString());
            }
            if (e.Column.FieldName == "Checked") // Thay đổi tên cột cho phù hợp
            {
               var getAcess= (bool)gridView1.GetRowCellValue(e.RowHandle, "isAcess");
                if (getAcess)
                {
                    // Lấy giá trị hiện tại của checkbox
                    bool currentValue = (bool)gridView1.GetRowCellValue(e.RowHandle, e.Column);

                    // Đảo ngược giá trị
                    gridView1.SetRowCellValue(e.RowHandle, e.Column, !currentValue);
                }
                else
                {
                    gridView1.SetRowCellValue(e.RowHandle, e.Column, false);
                }
            }
            if (e.Column.FieldName == "isHaschild")
            {
                //bool currentValue = (bool)gridView1.GetRowCellValue(e.RowHandle, e.Column);
                //gridView1.SetRowCellValue(e.RowHandle, e.Column, !currentValue);
                //var gettkCo = gridView1.GetRowCellValue(e.RowHandle, "TKNo").ToString();
                //foreach (var item in people)
                //{
                //    if (item.TKNo == gettkCo)
                //    {
                //        item.isHaschild = !currentValue;
                //        var index = people.IndexOf(item);
                //        gridView1.SetRowCellValue(index, e.Column, !currentValue);
                //    }
                //}
                //gridControl1.RefreshDataSource();
            }
        }

        private void gridView3_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (e.Column.FieldName == "Checked" ) // Thay đổi tên cột cho phù hợp
            {
                // Lấy giá trị hiện tại của checkbox
                bool currentValue = (bool)gridView3.GetRowCellValue(e.RowHandle, e.Column);
                // Đảo ngược giá trị
                gridView3.SetRowCellValue(e.RowHandle, e.Column, !currentValue);

            }
            if (e.Column.FieldName == "isHaschild")
            {
                //bool currentValue = (bool)gridView3.GetRowCellValue(e.RowHandle, e.Column);
                //var gettkCo = gridView3.GetRowCellValue(e.RowHandle, "TKCo").ToString();
                //foreach (var item in people2)
                //{
                //    if (item.TKCo == gettkCo)
                //        item.isHaschild = !currentValue;
                //}
                //gridControl2.RefreshDataSource();
            }
        }

        private void dtDenngay_EditValueChanged(object sender, EventArgs e)
        {
            //DateTime fromDate = (DateTime)dtTungay.EditValue;

            //DateTime toDate = fromDate.AddMonths(1).AddDays(-1); 
            //if (dtDenngay.DateTime > toDate)
            //    dtDenngay.EditValue = toDate;
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            progressPanel1.Visible = true;
            Application.DoEvents();
            string query = "SELECT * FROM tbimport"; 
            var kq = ExecuteQuery(query, null);
            if (kq.Rows.Count > 0)
            {
                foreach (DataRow item in kq.Rows)
                {
                    query = "SELECT * FROM HoaDon where SoHD=?  and KyHieu=? ";
                    var parameters = new OleDbParameter[]
                         {
            new OleDbParameter("?", item["SHDon"]),          // Sử dụng chỉ số mà không cần tên
            new OleDbParameter("?", item["KHHDon"])  // Thêm ký tự % cho LIKE
                         };
                    var kq2 = ExecuteQuery(query, parameters);
                    if (kq2.Rows.Count > 0)
                        continue;
                    //Nếu ko có thì xóa
                    query = "Delete  FROM tbimportdetail where ParentId=? ";
                     parameters = new OleDbParameter[]
                       {
            new OleDbParameter("?", item["ID"]),   
                       };
                    var kq3 = ExecuteQuery(query, parameters);
                    //
                    query = "Delete  FROM tbimport where ID=? ";
                    parameters = new OleDbParameter[]
                      {
            new OleDbParameter("?", item["ID"]),
                      };
                    var kq4 = ExecuteQuery(query, parameters);
                }
            }

           // XtraMessageBox.Show("Đã xóa dữ liệu dư thứa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            progressPanel1.Visible = false;
        }
        int uutienselect = 0;
        private void chkDaura_CheckedChanged(object sender, EventArgs e)
        { 
         
            if (chkDaura.Checked && (uutienselect == 2 || uutienselect == 0))
            {
                uutienselect = 2;
                chkDaura.Checked = true;
                chkDauvao.Checked = false;
                xtraTabControl2.SelectedTabPageIndex = 1;
            }
            int getthang = 0;
            //try
            //{
            //    string path = savedPath + @"\HDRa";
            //    var files = Directory.EnumerateFiles(path, "*.xml", SearchOption.AllDirectories).FirstOrDefault();
            //    if (files != null)
            //    {
            //        var getsplit = files.Split(new string[] { "\\" }, StringSplitOptions.None);
                   
            //        try
            //        {
            //            getthang = int.Parse(getsplit[getsplit.Length - 2].ToString());
            //            DateTime now = DateTime.Now; // Ngày hiện tại
            //            int year = now.Year; // Năm hiện tại 

            //            DateTime lastDayOfMonth = new DateTime(year, getthang, DateTime.DaysInMonth(year, getthang));
            //            dtTungay.DateTime = new DateTime(year, getthang, 1);
            //            dtDenngay.DateTime = new DateTime(year, getthang, lastDayOfMonth.Day);
            //        }
            //        catch(Exception ex)
            //        {

            //        }
            //    }
            //    else
            //    {

            //    }
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message + "   " + getthang.ToString());
            //}
        }

        private void chkDauvao_CheckedChanged(object sender, EventArgs e)
        {
             
            xtraTabControl2.SelectedTabPageIndex =0;
        }

        private void gridView1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            //// Lấy giá trị của cột 10 (chỉ số 9)
            try
            {
                object cellValue = gridView1.GetRowCellValue(e.RowHandle, gridView1.Columns["isAcess"]);

                // Nếu giá trị của cột 10 là false
                if (cellValue is bool && !(bool)cellValue)
                {
                    // Đặt màu nền và màu chữ để thể hiện dòng đã bị vô hiệu hóa
                    e.Appearance.BackColor = System.Drawing.Color.Red;
                }
            }
            catch( Exception ex)
            {

            }
            
        }

        private void gridView4_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
          
        }

        private void gridView3_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            try
            {
                object cellValue = gridView3.GetRowCellValue(e.RowHandle, gridView3.Columns["isAcess"]);

                //// Nếu giá trị của cột 10 là false
                if (cellValue is bool && !(bool)cellValue)
                {
                    // Đặt màu nền và màu chữ để thể hiện dòng đã bị vô hiệu hóa
                    e.Appearance.BackColor = System.Drawing.Color.Red; 
                }
            }
            catch(Exception ex)
            {

            }
          
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {

        }

        private void chkDaura_Click(object sender, EventArgs e)
        {
            uutienselect = 2;
        }

        private void chkDauvao_Click(object sender, EventArgs e)
        {
            uutienselect = 1;
        }

        private void chkDauvao_CheckedChanged_1(object sender, EventArgs e)
        {
            
            if (chkDaura.Checked && (uutienselect == 1 || uutienselect == 0))
            {
                uutienselect = 1;
                xtraTabControl2.SelectedTabPageIndex = 0;
                if (chkDauvao.Checked)
                {
                    chkDaura.Checked = false;
                    chkDauvao.Checked = true;
                }
            }
            int getthang = 0;
            try
            {
                string path = savedPath + @"\HDVao";
                var files = Directory.EnumerateFiles(path, "*.xml", SearchOption.AllDirectories).FirstOrDefault();
                if (files != null)
                {
                    var getsplit = files.Split(new string[] { "\\" }, StringSplitOptions.None);
                    getthang = int.Parse(getsplit[getsplit.Length - 2].ToString());
                    DateTime now = DateTime.Now; // Ngày hiện tại
                    int year = now.Year; // Năm hiện tại 

                    DateTime lastDayOfMonth = new DateTime(year, getthang, DateTime.DaysInMonth(year, getthang));
                    dtTungay.DateTime = new DateTime(year, getthang, 1);
                    dtDenngay.DateTime = new DateTime(year, getthang, lastDayOfMonth.Day);
                }
              
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message +"  "+ getthang.ToString());
            }
            
        }

        private void dtDenngay_EditValueChanged_1(object sender, EventArgs e)
        {

        }

        private void chkDauvao_MouseClick(object sender, MouseEventArgs e)
        {
        }

        private void simpleButton2_Click_1(object sender, EventArgs e)
        {
            var options = new ChromeOptions();
            // Tắt các cảnh báo bảo mật (Safe Browsing)

            // Tắt Safe Browsing và các tính năng bảo mật can thiệp
            //options.AddArgument("--disable-features=SafeBrowsing,DownloadBubble,DownloadNotification");
            options.AddArgument("--safebrowsing-disable-extension-blacklist");
            options.AddArgument("--safebrowsing-disable-download-protection");

            options.AddUserProfilePreference("download.prompt_for_download", false);
            //options.AddUserProfilePreference("safebrowsing.enabled", false);
            //options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
            // Tối ưu hóa trình duyệt

            options.AddArguments(
                //"--disable-notifications",
                "--start-maximized",
                "--disable-extensions",
                "--disable-infobars");
            //
            string downloadPath = "";
            downloadPath = savedPath + "\\HDVao";
            options.AddUserProfilePreference("download.default_directory", downloadPath);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            //options.AddUserProfilePreference("disable-popup-blocking", "true");
            options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
            options.AddUserProfilePreference("safebrowsing.enabled", false); // Tắt Safe Browsing hoàn toàn
            var driverPath = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Driver = new ChromeDriver(driverPath, options);
            Driver.Navigate().GoToUrl("https://hoadondientu.gdt.gov.vn");
        }

        private void gridView1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void gridView1_RowCellDefaultAlignment(object sender, DevExpress.XtraGrid.Views.Base.RowCellAlignmentEventArgs e)
        {

        }

        private void gridControl1_DataSourceChanged(object sender, EventArgs e)
        {
            progressPanel1.Visible = false;
        }

        private void gridView1_AsyncCompleted(object sender, EventArgs e)
        {
            
        }

        private void gridControl2_DataSourceChanged(object sender, EventArgs e)
        {
            progressPanel1.Visible = false;
        }

        private void gridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Escape)
            {
                frmTaikhoan frmTaikhoan = new frmTaikhoan();
                frmTaikhoan.frmMain = this;
                frmTaikhoan.ShowDialog();
              
            }
                if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                getMessage = true;
                DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;

                int currentRowHandle = gridView.FocusedRowHandle;

                // Lấy tên của cột hiện tại
                string currentColumnName = gridView.FocusedColumn.FieldName;
                object cellValue = gridView.GetRowCellValue(currentRowHandle, currentColumnName);

                if (cellValue != null)
                {
                    if (cellValue.ToString().Contains("154"))
                    {
                        // Thực hiện hành động mong muốn khi nhấn phím Tab
                        frmCongtrinh frmCongtrinh = new frmCongtrinh();
                        frmCongtrinh.frmMain = this;
                        frmCongtrinh.VatTu vatTu = new frmCongtrinh.VatTu();
                        vatTu.SoHieu = cellValue.ToString();
                        frmCongtrinh.dtoVatTu = vatTu;
                        frmCongtrinh.ShowDialog();
                        if (cellValue.ToString().Contains("|"))
                            cellValue = cellValue.ToString().Split('|')[0];
                        if (cellValue.ToString().Contains("154"))
                            gridView.SetRowCellValue(currentRowHandle, "TKNo", cellValue + "|" + hiddenValue);
                        if (cellValue.ToString().Contains("511"))
                        {
                            gridView.SetRowCellValue(currentRowHandle, "TKCo", cellValue + "|" + hiddenValue);

                        }
                        // Nếu bạn muốn ngăn chặn hành động mặc định của phím Tab

                        e.SuppressKeyPress = true;
                    }
                  
                }
                if (currentColumnName == "Ten")
                {
                     TooogleKH(gridView, cellValue.ToString(), currentRowHandle, gridView.GetRowCellValue(currentRowHandle, "Mst").ToString(), gridView.GetRowCellValue(currentRowHandle, "Ten").ToString());
                }
              
            }
           
        }

        private void gridView1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
        private int Kiemtrahttkcon(string tk)
        {
            string querydinhdanh = @"SELECT * FROM HeThongTK WHERE SoHieu LIKE ?";
            var resultkm = ExecuteQuery(querydinhdanh, new OleDbParameter("?", tk + "%"));
            return resultkm.Rows.Count;
        }
        private void gridView1_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            toolTipController.HideHint(); 

            var columns = e.Column;

            if (e.Column.ToString() == "TKNo")
            {
                var newValue = e.Value;
                var check = newValue is bool ? -1 : Kiemtrahttkcon(newValue.ToString());
                if (!newValue.ToString().Contains("|"))
                {
                    if (check == 0)
                    {
                        XtraMessageBox.Show("Số tài khoản không tồn tại trong hệ thống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SendKeys.Send("{F2}");

                        return;
                    }
                    if (check > 1)
                    {
                        XtraMessageBox.Show("Tài khoản " + newValue + " có tài khoản con, vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SendKeys.Send("{F2}");

                        return;
                    }
                }
            }
           
          
            DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;

            // Kiểm tra nếu hàng có detail view đang mở
            //if (gridView.GetDetailView(e.RowHandle, 0) != null)
            //{
            //    return; // Nếu có hàng con mở, không thực hiện hành động
            //}

            // Chỉ thực hiện khi không có hàng con mở
            foreach (var item in lstImportVao)
            {
                foreach (var it2 in item.fileImportDetails)
                {
                    it2.TKCo = item.TKCo;
                    it2.TKNo = item.TKNo;
                }
            }

            var gridViews = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            var idValue = gridViews.GetRowCellValue(e.RowHandle, "ID");
            UpdateData(int.Parse(idValue.ToString()), lstImportVao);
            gridControl1.DataSource = lstImportVao;
            gridControl1.RefreshDataSource();
        }
        private void UpdateData(int id,BindingList<FileImport> lst)
        {
            var getfip = lst.Where(m => m.ID == id).FirstOrDefault();
            string query = @"UPDATE tbimport SET NoiDung=?, TKNo=?, TKCo=?, TKThue=?,SohieuTP=? where ID=? ";
            //if (getfip.TKNo.Contains("|"))
            //{
            //    var getsplit = getfip.TKNo.Split('|');
            //    getfip.TKNo = getsplit[0];
            //    getfip.SoHieuTP= getsplit[1];
            //}
           var parameters = new OleDbParameter[]
            {
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(getfip.Noidung)),
            new OleDbParameter("?", getfip.TKNo),
            new OleDbParameter("?",  getfip.TKCo) ,
              new OleDbParameter("?", getfip.TkThue),
                new OleDbParameter("?", getfip.SoHieuTP),
               new OleDbParameter("?", getfip.ID)
            };
            int rowsAffected = ExecuteQueryResult(query, parameters);
            //
             query = @"UPDATE tbimportdetail SET  TKNo=?, TKCo=? where ParentId=? ";
             parameters = new OleDbParameter[]
           { 
            new OleDbParameter("?", getfip.TKNo),
            new OleDbParameter("?",  getfip.TKCo) , 
               new OleDbParameter("?", getfip.ID)
           };
             rowsAffected = ExecuteQueryResult(query, parameters);
        }

        private void gridView1_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void gridView2_KeyUp(object sender, KeyEventArgs e)
        {
           
        }

        private void frmMain_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void gridView2_RowCellClick(object sender, RowCellClickEventArgs e)
        {

        }
        private int clickCount = 0;
        private System.Windows.Forms.Timer clickTimer;
        private void gridView1_MouseDown(object sender, MouseEventArgs e)
        {
            
        }
        private List<int> lstrowSohieu = new List<int>();
        private void gridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                lstrowSohieu = new List<int>();
                getMessage = true;
                DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;

                int currentRowHandle = gridView.FocusedRowHandle;

                // Lấy tên của cột hiện tại
                string currentColumnName = gridView.FocusedColumn.FieldName;
                object cellValue = gridView.GetRowCellValue(currentRowHandle, currentColumnName);

                if (cellValue != null)
                {
                    if (cellValue.ToString().Contains("154"))
                    {
                        // Thực hiện hành động mong muốn khi nhấn phím Tab
                        frmCongtrinh frmCongtrinh = new frmCongtrinh();
                        frmCongtrinh.frmMain = this;
                        frmCongtrinh.ShowDialog();
                        if (cellValue.ToString().Contains("|"))
                            cellValue = cellValue.ToString().Split('|')[0];
                        if (cellValue.ToString().Contains("154"))
                            gridView.SetRowCellValue(currentRowHandle, "TKNo", cellValue + "|" + hiddenValue);
                        if (cellValue.ToString().Contains("511"))
                        {
                            gridView.SetRowCellValue(currentRowHandle, "TKCo", cellValue + "|" + hiddenValue);

                        }
                        // Nếu bạn muốn ngăn chặn hành động mặc định của phím Tab

                        e.SuppressKeyPress = true;
                    }
                    else
                    {
                        if (currentColumnName.ToString() == "SoHieu")
                        {
                            frmHangHoa frmHangHoa = new frmHangHoa();
                            frmHangHoa.frmMain = this;

                             VatTu vatTu = new VatTu();
                            vatTu.SoHieu = cellValue.ToString();
                            //Kiểm tra xem có phải so hiệu tự tạo 
                            string querydinhdanh = @"SELECT * FROM Vattu WHERE SoHieu = ?";
                            var checkSH = ExecuteQuery(querydinhdanh, new OleDbParameter("?", vatTu.SoHieu));
                            //Nếu chưa có, tìm hết danh sách
                            if (checkSH.Rows.Count == 0)
                            {
                                foreach(var item in lstImportVao)
                                {
                                    foreach(var it in item.fileImportDetails)
                                    {
                                        if(it.SoHieu== vatTu.SoHieu)
                                        {
                                            lstrowSohieu.Add(it.ID);
                                        }
                                    }
                                }
                            }

                            var tenvattu = gridView.GetRowCellValue(currentRowHandle, "Ten").ToString();
                            vatTu.TenVattu = tenvattu;
                            var dvt = gridView.GetRowCellValue(currentRowHandle, "DVT").ToString();
                            vatTu.DonVi = dvt;
                            frmHangHoa.dtoVatTu = vatTu;
                            if (checkSH.Rows.Count > 0)
                                currentselectId = int.Parse(checkSH.Rows[0]["MaPhanLoai"].ToString());
                            else
                                currentselectId = 0;
                            frmHangHoa.ShowDialog();

                        
                            
                            if (!string.IsNullOrEmpty(hiddenValue) && frmHangHoa.isChange)
                            {
                                gridView.SetRowCellValue(currentRowHandle, currentColumnName, hiddenValue);
                                gridView.SetRowCellValue(currentRowHandle, "DVT", hiddenValue2);
                                gridView.SetRowCellValue(currentRowHandle, "Ten", hiddenValue3);
                                //Fill lai cho lstrowSohieu
                                foreach(var item in lstrowSohieu)
                                {
                                    foreach(var it in lstImportVao)
                                    {
                                        var findchild = it.fileImportDetails.Where(m => m.ID == item).FirstOrDefault();
                                        if (findchild != null)
                                        {
                                            findchild.SoHieu = hiddenValue;
                                            findchild.DVT = hiddenValue2;
                                            findchild.Ten = hiddenValue3;
                                            //Update vào database
                                         var   query = @"UPDATE tbimportdetail SET  SoHieu=?,DVT=?,Ten=? where ID=? ";
                                          var  parameters = new OleDbParameter[]
                                          {
                                                new OleDbParameter("?", hiddenValue),
                                                new OleDbParameter("?",  Helpers.ConvertUnicodeToVni(hiddenValue2)) ,
                                                new OleDbParameter("?", Helpers.ConvertUnicodeToVni(hiddenValue3)),
                                                new OleDbParameter("?", findchild.ID)
                                          };
                                          var  rowsAffected = ExecuteQueryResult(query, parameters);
                                        }
                                    }
                                }
                                lstrowSohieu = new List<int>();
                            }

                        }
                    }
                }

            }

            if (e.KeyCode == System.Windows.Forms.Keys.Delete)
            {
               
               
            } 
           
        }

        private void gridView3_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            toolTipController.HideHint();
            if (e.Column.ToString() == "TKCo")
            {
                var newValue = e.Value;
                var check = newValue is bool ? -1 : Kiemtrahttkcon(newValue.ToString());
                if (!newValue.ToString().Contains("|"))
                {
                    if (check == 0)
                    {
                        XtraMessageBox.Show("Số tài khoản không tồn tại trong hệ thống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SendKeys.Send("{F2}");

                        return;
                    }
                    if (check > 1)
                    {
                        XtraMessageBox.Show("Tài khoản " + newValue + " có tài khoản con, vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SendKeys.Send("{F2}");

                        return;
                    }
                }
            }
               
          
            DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;

            // Kiểm tra nếu hàng có detail view đang mở
            //if (gridView.GetDetailView(e.RowHandle, 0) != null)
            //{
            //    return; // Nếu có hàng con mở, không thực hiện hành động
            //}

            // Chỉ thực hiện khi không có hàng con mở
            foreach (var item in lstImportRa)
            {
                foreach (var it2 in item.fileImportDetails)
                {
                    it2.TKCo = item.TKCo;
                    it2.TKNo = item.TKNo;
                }
            }

            var gridViews = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            var idValue = gridViews.GetRowCellValue(e.RowHandle, "ID");
            UpdateData(int.Parse(idValue.ToString()), lstImportRa);
            gridControl2.DataSource = lstImportRa;
            gridControl2.RefreshDataSource();
        }

        private void gridView4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                getMessage = true;
                DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;

                int currentRowHandle = gridView.FocusedRowHandle;

                // Lấy tên của cột hiện tại
                string currentColumnName = gridView.FocusedColumn.FieldName;
                object cellValue = gridView.GetRowCellValue(currentRowHandle, currentColumnName);

                if (cellValue != null)
                {
                    if (cellValue.ToString().Contains("511"))
                    {
                        // Thực hiện hành động mong muốn khi nhấn phím Tab
                        frmCongtrinh frmCongtrinh = new frmCongtrinh();
                        frmCongtrinh.frmMain = this;
                        frmCongtrinh.VatTu vatTu = new frmCongtrinh.VatTu();
                        vatTu.SoHieu = cellValue.ToString();  
                        frmCongtrinh.dtoVatTu = vatTu;
                        frmCongtrinh.ShowDialog();
                        if (cellValue.ToString().Contains("|"))
                            cellValue = cellValue.ToString().Split('|')[0];
                        if (cellValue.ToString().Contains("154"))
                            gridView.SetRowCellValue(currentRowHandle, "TKNo", cellValue + "|" + hiddenValue);
                        if (cellValue.ToString().Contains("511"))
                        {
                            gridView.SetRowCellValue(currentRowHandle, "TKCo", cellValue + "|" + hiddenValue);

                        }
                        // Nếu bạn muốn ngăn chặn hành động mặc định của phím Tab

                        e.SuppressKeyPress = true;
                    }
                    else
                    {
                        if (currentColumnName.ToString() == "SoHieu")
                        {
                            frmHangHoa frmHangHoa = new frmHangHoa();
                            frmHangHoa.frmMain = this;

                           VatTu vatTu = new VatTu();
                            vatTu.SoHieu = cellValue.ToString();

                            string querydinhdanh = @"SELECT * FROM Vattu WHERE SoHieu = ?";
                            var checkSH = ExecuteQuery(querydinhdanh, new OleDbParameter("?", vatTu.SoHieu));
                            //Nếu chưa có, tìm hết danh sách
                            if (checkSH.Rows.Count == 0)
                            {
                                foreach (var item in lstImportRa)
                                {
                                    foreach (var it in item.fileImportDetails)
                                    {
                                        if (it.SoHieu == vatTu.SoHieu)
                                        {
                                            lstrowSohieu.Add(it.ID);
                                        }
                                    }
                                }
                            }

                            var tenvattu = gridView.GetRowCellValue(currentRowHandle, "Ten").ToString();
                            vatTu.TenVattu = tenvattu;
                            var dvt = gridView.GetRowCellValue(currentRowHandle, "DVT").ToString();
                            vatTu.DonVi = dvt;
                            frmHangHoa.dtoVatTu = vatTu;
                            if (checkSH.Rows.Count > 0)
                                currentselectId = int.Parse(checkSH.Rows[0]["MaPhanLoai"].ToString());
                            else
                                currentselectId = 0;
                            frmHangHoa.ShowDialog();


                            if (!string.IsNullOrEmpty(hiddenValue) && frmHangHoa.isChange)
                            {
                                gridView.SetRowCellValue(currentRowHandle, currentColumnName, hiddenValue);
                                gridView.SetRowCellValue(currentRowHandle, "DVT", hiddenValue2);
                                gridView.SetRowCellValue(currentRowHandle, "Ten", hiddenValue3);
                                foreach (var item in lstrowSohieu)
                                {
                                    foreach (var it in lstImportRa)
                                    {
                                        var findchild = it.fileImportDetails.Where(m => m.ID == item).FirstOrDefault();
                                        if (findchild != null)
                                        {
                                            findchild.SoHieu = hiddenValue;
                                            findchild.DVT = hiddenValue2;
                                            findchild.Ten = hiddenValue3;
                                            //Update vào database
                                            var query = @"UPDATE tbimportdetail SET  SoHieu=?,DVT=?,Ten=? where ID=? ";
                                            var parameters = new OleDbParameter[]
                                            {
                                        new OleDbParameter("?", hiddenValue),
                                        new OleDbParameter("?",  Helpers.ConvertUnicodeToVni(hiddenValue2)) ,
                                        new OleDbParameter("?", Helpers.ConvertUnicodeToVni(hiddenValue3)),
                                        new OleDbParameter("?", findchild.ID)
                                            };
                                            var rowsAffected = ExecuteQueryResult(query, parameters);
                                        }
                                    }
                                }
                                lstrowSohieu = new List<int>();
                            }

                        }
                    }
                }

            }
        }
        public DataTable tbKhachhang = new DataTable();
        private void TooogleKH(DevExpress.XtraGrid.Views.Grid.GridView gridView,string cellValue,int currentRowHandle,string mst,string ten)
        {
            frmKhachhang frmKhachhang = new frmKhachhang();
            frmKhachhang.frmMain = this;

            frmKhachhang.Khachhang vatTu = new frmKhachhang.Khachhang();
            DataRow kh =null ;
            if (!string.IsNullOrEmpty(mst))
            {
                 kh = tbKhachhang.AsEnumerable().Where(row => row.Field<string>("MST") == mst).FirstOrDefault();
            }
            else
            {
                if (!string.IsNullOrEmpty(ten))
                {
                    ten = Helpers.ConvertUnicodeToVni(ten);
                    kh = tbKhachhang.AsEnumerable().Where(row => Helpers.RemoveVietnameseDiacritics(row.Field<string>("Ten").ToLower()) == Helpers.RemoveVietnameseDiacritics(ten.ToLower())).FirstOrDefault();
                }
            }
            vatTu.Ten = Helpers.ConvertVniToUnicode(ten);
            vatTu.Mst = kh["MST"].ToString();
            vatTu.DiaChi = kh["DiaChi"].ToString();
            vatTu.SoHieu= kh["SoHieu"].ToString();
            vatTu.MaPhanLoai= int.Parse(kh["MaPhanLoai"].ToString());
            frmKhachhang.dtoVatTu = vatTu;
            frmKhachhang.ShowDialog();
        }
        private void gridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                getMessage = true;
                DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;

                int currentRowHandle = gridView.FocusedRowHandle;

                // Lấy tên của cột hiện tại
                string currentColumnName = gridView.FocusedColumn.FieldName;
                object cellValue = gridView.GetRowCellValue(currentRowHandle, currentColumnName);

                if (cellValue != null)
                {
                    if (cellValue.ToString().Contains("511"))
                    {
                        if (!Kiemtrataikhoancon(cellValue.ToString()))
                        {
                            // Thực hiện hành động mong muốn khi nhấn phím Tab
                            frmCongtrinh frmCongtrinh = new frmCongtrinh();
                            frmCongtrinh.frmMain = this;
                            frmCongtrinh.VatTu vatTu = new frmCongtrinh.VatTu();
                            vatTu.SoHieu = cellValue.ToString();
                            frmCongtrinh.dtoVatTu = vatTu;
                            frmCongtrinh.ShowDialog();
                            if (cellValue.ToString().Contains("|"))
                                cellValue = cellValue.ToString().Split('|')[0];
                            if (cellValue.ToString().Contains("511"))
                            {
                                gridView.SetRowCellValue(currentRowHandle, "TKCo", cellValue + "|" + hiddenValue);

                            }
                        }
                        // Nếu bạn muốn ngăn chặn hành động mặc định của phím Tab

                        e.SuppressKeyPress = true;
                    }
                    if (currentColumnName == "Ten")
                    {
                        TooogleKH(gridView, cellValue.ToString(), currentRowHandle, gridView.GetRowCellValue(currentRowHandle, "Mst").ToString(), gridView.GetRowCellValue(currentRowHandle, "Ten").ToString());
                    }
                }

            }
        }

        private void gridControl1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void gridView2_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 1) // Nếu là single click
            {
                gridView2.ShowEditor(); // Hiển thị chế độ chỉnh sửa
            }
        }

        private void gridView4_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 1) // Nếu là single click
            {
                gridView4.ShowEditor(); // Hiển thị chế độ chỉnh sửa
            }
        } 

        private void gridView3_ShownEditor(object sender, EventArgs e)
        {
           
        }

        private void gridView3_KeyUp(object sender, KeyEventArgs e)
        {
           
        }

        private void gridView3_CellValueChanging(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            gridControl1.ToolTipController = toolTipController1;
            DevExpress.XtraGrid.Views.Grid.GridView view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            var newValue = e.Value;
            if (newValue.ToString().Length < 4)
            {
                toolTipController.HideHint();
                return;
            }
            if (view.FocusedColumn.FieldName == "TKCo") // Thay đổi tên cột theo nhu cầu
            {
                // Hiển thị tooltip
                string querydinhdanh = @"SELECT * FROM HeThongTK WHERE SoHieu LIKE ?";
                var resultkm = ExecuteQuery(querydinhdanh, new OleDbParameter("?", newValue + "%"));
                var str = "";
                var strBuilder = new System.Text.StringBuilder();

                if (resultkm.Rows.Count > 0)
                {
                    for (int i = 0; i < resultkm.Rows.Count; i++)
                    {
                        strBuilder.AppendLine(resultkm.Rows[i]["SoHieu"].ToString()+"-"+ Helpers.ConvertVniToUnicode(resultkm.Rows[i]["Ten"].ToString())); // Thêm xuống dòng
                    }
                }
                if (resultkm.Rows.Count > 1)
                {
                    string tooltipText = strBuilder.ToString().Trim(); // Loại bỏ dòng trống cuối cùng (nếu có)

                    var gridViewInfo = view.GetViewInfo() as GridViewInfo;

                    GridCellInfo cellInfo = gridViewInfo.GetGridCellInfo(e.RowHandle, view.FocusedColumn);

                    int rowHandle = view.FocusedRowHandle;
                    System.Drawing.Rectangle cellRect = cellInfo.Bounds;
                    System.Drawing.Point screenPoint = view.GridControl.PointToScreen(
                        new System.Drawing.Point(cellRect.X, cellRect.Bottom + 2)
                    );

                    toolTipController.ShowHint(tooltipText, screenPoint);
                }
                
            }
        }

        private void gridView1_CellValueChanging(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            gridControl1.ToolTipController = toolTipController1;
            DevExpress.XtraGrid.Views.Grid.GridView view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            var newValue = e.Value;
            if (newValue.ToString().Length < 3)
            {
                toolTipController.HideHint();
                return;
            }
            if (view.FocusedColumn.FieldName == "TKNo") // Thay đổi tên cột theo nhu cầu
            {
                // Hiển thị tooltip
                string querydinhdanh = @"SELECT * FROM HeThongTK WHERE SoHieu LIKE ?";
                var resultkm = ExecuteQuery(querydinhdanh, new OleDbParameter("?", newValue + "%"));
                var str = "";
                var strBuilder = new System.Text.StringBuilder();

                if (resultkm.Rows.Count > 0)
                {
                    for (int i = 0; i < resultkm.Rows.Count; i++)
                    {
                        strBuilder.AppendLine(resultkm.Rows[i]["SoHieu"].ToString() + "-" + Helpers.ConvertVniToUnicode(resultkm.Rows[i]["Ten"].ToString())); // Thêm xuống dòng
                    }
                }
                if (resultkm.Rows.Count > 1)
                {
                    string tooltipText = strBuilder.ToString().Trim(); // Loại bỏ dòng trống cuối cùng (nếu có)

                    var gridViewInfo = view.GetViewInfo() as GridViewInfo;

                    GridCellInfo cellInfo = gridViewInfo.GetGridCellInfo(e.RowHandle, view.FocusedColumn);

                    int rowHandle = view.FocusedRowHandle;
                    System.Drawing.Rectangle cellRect = cellInfo.Bounds;
                    System.Drawing.Point screenPoint = view.GridControl.PointToScreen(
                        new System.Drawing.Point(cellRect.X, cellRect.Bottom + 2)
                    );

                    toolTipController.ShowHint(tooltipText, screenPoint);
                }

            }
        }

        private void gridView1_Click(object sender, EventArgs e)
        {

        }

        private void gridView2_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            
        }

        private void gridView2_RowStyle(object sender, RowStyleEventArgs e)
        {
           
        }
        private void LoadCustomDrawcell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column.FieldName == "SoHieu")
            {
                if (e.CellValue == null)
                    return;
                var getvalue = e.CellValue.ToString();
                // 1. Vẽ nền mặc định (giữ nguyên background)
                //e.DefaultDraw();
                string querydinhdanh = @"SELECT * FROM Vattu WHERE SoHieu LIKE ?";
                var resultkm = ExecuteQuery(querydinhdanh, new OleDbParameter("?", getvalue));
                // 2. Vẽ chữ màu đè lên (màu đỏ)
                if (resultkm.Rows.Count == 0)
                {
                    System.Drawing.Font boldFont = new System.Drawing.Font(e.Appearance.Font, FontStyle.Bold);

                    e.Cache.DrawString(
                    e.DisplayText,
                    boldFont,
                    Brushes.DarkBlue,  // Chữ màu đỏ
                    e.Bounds,
                    e.Appearance.GetStringFormat()
                );

                    e.Handled = true; // Ngăn vẽ mặc định
                }

            }
        }
        private void gridView2_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            LoadCustomDrawcell(sender, e);
        }

        private void gridView4_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            LoadCustomDrawcell(sender, e);
        }
        private void ConvertImageTOexcel()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Thiết lập các thuộc tính cho hộp thoại mở file
            openFileDialog.Title = "Chọn file PDF hóa đơn"; // Tiêu đề của hộp thoại
            openFileDialog.Filter = "Tệp PDF (*.pdf)|*.pdf|Tất cả các tệp (*.*)|*.*"; // Lọc loại file
            openFileDialog.FilterIndex = 1; // Chọn bộ lọc mặc định là PDF
            openFileDialog.RestoreDirectory = true; // Nhớ thư mục đã mở lần trước
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Người dùng đã chọn một file
                string inputFile = openFileDialog.FileName; // Lấy đường dẫn file đã chọn

                // Định nghĩa đường dẫn file đầu ra (bạn có thể muốn cho người dùng chọn cả đường dẫn này)
                // Hiện tại, chúng ta sẽ đặt tên và thư mục mặc định gần file đầu vào
                string outputDirectory = Path.Combine(Path.GetDirectoryName(inputFile), "Converted_Excel");
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                string outputFile = savedPath + @"\output.xlsx";

                // Gọi phương thức để thực hiện chuyển đổi
                // (Phương thức này sẽ chứa logic gọi ABBYY FineReader qua CMD)

                // Định nghĩa đường dẫn đến file thực thi của ABBYY FineReader
                string abbyyExePath = @"C:\Program Files\ABBYY FineReader 16\finereaderocr.exe";

                // --- Kiểm tra trước khi chạy ---
                if (!File.Exists(abbyyExePath))
                {
                    Console.WriteLine($"Lỗi: Không tìm thấy file thực thi của ABBYY FineReader tại: {abbyyExePath}");
                    Console.WriteLine("Vui lòng kiểm tra lại đường dẫn cài đặt ABBYY FineReader.");
                    return;
                }

                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Lỗi: File đầu vào không tồn tại: {inputFile}");
                    return;
                }

                // Tạo thư mục đầu ra nếu chưa có 
                if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                    Console.WriteLine($"Đã tạo thư mục đầu ra: {outputDirectory}");
                }
                // --- Kết thúc kiểm tra ---

                // Xây dựng chuỗi đối số cho lệnh ABBYY
                // Đảm bảo các đường dẫn có khoảng trắng được đặt trong dấu ngoặc kép kép (escape " thành "")
                string arguments = $@"""{inputFile}"" /out ""{outputFile}"" /format ""XLSX"" /lang ""Vietnamese""";

                Console.WriteLine($"\nĐang cố gắng gọi ABBYY FineReader với lệnh:");
                Console.WriteLine($"Application: {abbyyExePath}");
                Console.WriteLine($"Arguments: {arguments}");

                try
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo
                    {
                        FileName = abbyyExePath,     // Đường dẫn đến file .exe cần chạy
                        Arguments = arguments,       // Các tham số dòng lệnh
                        UseShellExecute = false,     // Rất quan trọng để chuyển hướng output và ẩn cửa sổ
                        RedirectStandardOutput = true, // Chuyển hướng output chuẩn
                        RedirectStandardError = true,  // Chuyển hướng lỗi chuẩn
                        CreateNoWindow = true        // Không hiển thị cửa sổ console riêng
                    };

                    using (Process process = Process.Start(startInfo))
                    {
                        // Đọc luồng output và error một cách bất đồng bộ để tránh deadlock
                        string output = process.StandardOutput.ReadToEnd();
                        string error = process.StandardError.ReadToEnd();

                        process.WaitForExit(); // Chờ process hoàn thành

                        Console.WriteLine("\n--- Kết quả từ ABBYY Process ---");
                        Console.WriteLine("Standard Output:\n" + output);
                        Console.WriteLine("Standard Error:\n" + error);
                        Console.WriteLine($"Mã thoát (Exit Code): {process.ExitCode}");
                        Console.WriteLine("--------------------------------\n");

                        if (process.ExitCode == 0)
                        {
                            Console.WriteLine("Lệnh ABBYY FineReader đã hoàn tất (Exit Code 0).");
                            if (File.Exists(outputFile))
                            {

                            }
                            else
                            {
                                Console.WriteLine("Lệnh hoàn tất với Exit Code 0 nhưng file Excel đầu ra không được tìm thấy. Có thể do giới hạn giấy phép ABBYY FineReader PDF thông thường.");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Lệnh ABBYY FineReader kết thúc với **mã lỗi {process.ExitCode}**. Vui lòng kiểm tra lỗi trên.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Đã xảy ra lỗi khi cố gắng chạy ABBYY FineReader: {ex.Message}");
                    Console.WriteLine($"StackTrace: {ex.StackTrace}");
                }
            }
            else
            {
                MessageBox.Show("Bạn chưa chọn file nào.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            // Đường dẫn đến file PDF và file đầu ra
            // Định nghĩa các đường dẫn file

        }
        public class Nganhang
        {
            public int Stt { get; set; }
            public DateTime NgayGD { get; set; }
            public string Maso { get; set; }
            public string Diengiai { get; set; }
            public double ThanhTien { get; set; }
            public double ThanhTien2 { get; set; }
            public string TKNo { get; set; }
            public string TKCo { get; set; }
        }
        private List<Nganhang> lstNganhan = new List<Nganhang>();
        System.Windows.Forms.BindingSource bdNganhang = new System.Windows.Forms.BindingSource();
        public static class DateHelper
        {
            /// <summary>
            /// Phân tích một chuỗi ngày tháng có thể bị lỗi ở phần năm (ví dụ: "dd/MM/yyyyX HH:mm:ss")
            /// hoặc có ký tự thừa ở cuối (ví dụ: "dd/MM/yyyy HH:mm:ss Z").
            /// </summary>
            /// <param name="invalidDateString">Chuỗi ngày tháng cần phân tích.</param>
            /// <returns>Đối tượng DateTime nếu phân tích thành công, ngược lại là null.</returns>
            public static DateTime? ParseAndCorrectDate(string dateStringFromExcel)
            {
                if (string.IsNullOrWhiteSpace(dateStringFromExcel))
                {
                    return null;
                }

                DateTime parsedDate;
                // Định dạng đầy đủ mà chúng ta muốn khớp (dd/MM/yyyy HH:mm:ss)
                string fullFormat = "dd/MM/yyyy HH:mm:ss";
                // Định dạng chỉ ngày (dd/MM/yyyy)
                string dateFormat = "dd/MM/yyyy";

                // Bước 1: Cố gắng phân tích chuỗi gốc với định dạng đầy đủ trước.
                // Điều này xử lý các trường hợp chuỗi đã đúng định dạng (có hoặc không có ký tự thừa sau giây).
                if (DateTime.TryParseExact(dateStringFromExcel, fullFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                {
                    return parsedDate;
                }

                // Bước 2: Nếu phân tích trực tiếp thất bại, thử phân tích và sửa lỗi bằng Regex.
                // Regex để bắt các phần ngày (dd/MM/), năm (có thể 4 hoặc 5 chữ số), và TÙY CHỌN phần thời gian HH:mm:ss.
                // Nó sẽ bỏ qua bất kỳ thứ gì sau phần thời gian (nếu có).
                // ^(\d{2}/\d{2}/)(\d+)          : Bắt đầu chuỗi, ngày/tháng/ và năm (4 hoặc 5 chữ số)
                // (?:\s*(\d{2}:\d{2}:\d{2}))? : TÙY CHỌN: khoảng trắng, sau đó là HH:mm:ss.
                //                               Nếu không có, Group[3] sẽ không khớp.
                // .* : Khớp bất kỳ ký tự nào còn lại (để bỏ qua phần '5139.88796' hoặc các ký tự thừa khác)
                Match match = Regex.Match(dateStringFromExcel, @"^(\d{2}/\d{2}/)(\d+)(?:\s*(\d{2}:\d{2}:\d{2}))?.*");

                if (match.Success)
                {
                    string datePartPrefix = match.Groups[1].Value; // Ví dụ: "01/04/"
                    string yearPart = match.Groups[2].Value;       // Ví dụ: "2025" hoặc "20251"
                                                                   // Nếu Group[3] khớp (tức là tìm thấy thời gian chuẩn), thì lấy giá trị đó.
                                                                   // Ngược lại, mặc định thời gian là "00:00:00".
                    string timePart = match.Groups[3].Success ? match.Groups[3].Value : "00:00:00";

                    // Sửa phần năm nếu có 5 chữ số (như "20251" thành "2025")
                    if (yearPart.Length == 5)
                    {
                        yearPart = yearPart.Substring(0, 4);
                    }

                    // Kết hợp lại thành chuỗi định dạng chuẩn để thử phân tích
                    string combinedString = $"{datePartPrefix}{yearPart} {timePart}";

                    if (DateTime.TryParseExact(combinedString, fullFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                    {
                        return parsedDate;
                    }
                }

                // Bước 3: Nếu Regex phức tạp không khớp được, thử một cách đơn giản hơn: chỉ phân tích phần ngày.
                // Điều này hữu ích nếu chuỗi chỉ chứa ngày và có thể có bất kỳ thứ gì sau đó.
                // Ví dụ: "01/04/2025 5139.88796" hoặc "05/03/2020"
                match = Regex.Match(dateStringFromExcel, @"^(\d{2}/\d{2}/\d{4})"); // Chỉ bắt dd/MM/yyyy từ đầu chuỗi

                if (match.Success)
                {
                    string dateOnlyString = match.Groups[1].Value; // Lấy "01/04/2025"
                    if (DateTime.TryParseExact(dateOnlyString, dateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                    {
                        // Nếu chỉ có ngày, thời gian sẽ là 00:00:00 (mặc định của DateTime khi chỉ parse ngày)
                        return parsedDate;
                    }
                }

                // Nếu không thể phân tích được với bất kỳ phương pháp nào, trả về null
                return null;
            }
        }
            private void ReadExcelBank(string filePath)
        {
            // Xóa dữ liệu cũ nếu phương thức này được gọi nhiều lần
            lstNganhan.Clear();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // Lấy sheet đầu tiên

                // --- TÌM HÀNG TIÊU ĐỀ VÀ CHỈ SỐ CỘT ---
                int indexDiengiai = 0;
                int indexNgaygiaodich = 0;
                int indexMaGD = 0;
                int indexTKCo = 0; // Số tiền rút (Debit)
                int indexTKNo = 0; // Số tiền gửi (Credit)
                int headerRowNumber = 0; // Số hàng chứa tiêu đề

                // Duyệt qua tất cả các ô có dữ liệu để tìm hàng tiêu đề và các chỉ số cột
                // headerRowNumber sẽ lưu lại số hàng cao nhất mà một tiêu đề được tìm thấy,
                // giả định đó là hàng tiêu đề chính và dữ liệu sẽ bắt đầu từ hàng tiếp theo.
                bool findSTT = true;
                foreach (var cell in worksheet.CellsUsed())
                {
                    var getdata = cell.GetString().Trim(); // Lấy nội dung ô và loại bỏ khoảng trắng thừa

                    if (getdata.Contains("STT") && findSTT)
                    {
                        // Cập nhật headerRowNumber nếu tìm thấy STT ở hàng thấp hơn hàng đã lưu
                        if (cell.Address.RowNumber > headerRowNumber)
                        {
                            headerRowNumber = cell.Address.RowNumber;
                            findSTT = false;
                        }
                    }
                    //if (headerRowNumber!=0 && (cell.Address.RowNumber + 1) > headerRowNumber)
                    //    break;
                    // Tìm chỉ số cột, không phụ thuộc vào headerRowNumber hiện tại
                    if (getdata.Contains("Ngày GD") || getdata.Contains("Ngày giao dịch") || getdata.Contains("Tran Date"))
                    {
                        indexNgaygiaodich = cell.Address.ColumnNumber;
                        headerRowNumber= cell.Address.RowNumber;
                    }
                    else if (getdata.Contains("Description") || getdata.Contains("Transactions"))
                    {
                        indexDiengiai = cell.Address.ColumnNumber;
                    }
                    else if ( (getdata.Contains("GD") || getdata.Contains("GDV") || getdata.Contains("So CT")) && indexMaGD==0)
                    {
                        indexMaGD = cell.Address.ColumnNumber;
                    }
                    else if (getdata.Contains("Sô tiên rút") || getdata.Contains("Debit") || getdata.Contains("Phát sinh nợ"))
                    {
                        indexTKCo = cell.Address.ColumnNumber;
                    }
                    else if (getdata.Contains("SÔ tiên gửi") || getdata.Contains("Credit") || getdata.Contains("Phát sinh có"))
                    {
                        indexTKNo = cell.Address.ColumnNumber;
                    }
                    if (headerRowNumber>0 && (headerRowNumber < cell.Address.RowNumber))
                        break;
                }

                // Kiểm tra xem đã tìm thấy đủ các cột cần thiết chưa
                //if (headerRowNumber == 0 || indexNgaygiaodich == 0 || indexDiengiai == 0 || indexMaGD == 0 || indexTKCo == 0 || indexTKNo == 0)
                //{
                //    // Thông báo lỗi nếu không tìm thấy các tiêu đề cần thiết
                //    System.Windows.Forms.MessageBox.Show("Không tìm thấy đủ các cột tiêu đề cần thiết (Ngày GD, Diễn giải, Mã GD, Số tiền rút/gửi). Vui lòng kiểm tra lại file Excel.", "Lỗi Định dạng", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                //    return; // Thoát phương thức
                //}

                // --- ĐỌC DỮ LIỆU TỪ CÁC HÀNG ---
                // Bắt đầu vòng lặp từ hàng ngay sau hàng tiêu đề (headerRowNumber + 1)
                // và kết thúc ở hàng cuối cùng có dữ liệu trong worksheet.
                int sttValue= 1; 
                for (int rowIndex = headerRowNumber + 1; rowIndex <= worksheet.LastCellUsed().Address.RowNumber; rowIndex++)
                {
                    var row = worksheet.Row(rowIndex);

                    // --- Xác thực cơ bản để bỏ qua các hàng trống hoặc không phải dữ liệu ---
                    // Kiểm tra xem ô STT hoặc ô Ngày GD có dữ liệu không.
                    // Giả định cột 1 là cột STT, nếu nó trống, bỏ qua hàng.
                    string sttCellContent = row.Cell(1).GetString().Trim();
                    string ngayGDCellContent = row.Cell(indexNgaygiaodich).GetString().Trim();

                    if (string.IsNullOrWhiteSpace(sttCellContent) && string.IsNullOrWhiteSpace(ngayGDCellContent))
                    {
                        continue; // Bỏ qua nếu cả ô STT và Ngày GD đều trống, có thể là hàng trống
                    }

                    // Cố gắng phân tích STT để đảm bảo đây là hàng dữ liệu thực sự (nếu STT là số)
                    //int sttValue;
                    //if (!int.TryParse(sttCellContent, out sttValue))
                    //{
                    //    // Nếu cột STT không phải là số, có thể đây là hàng tổng kết hoặc không phải dữ liệu giao dịch
                    //    // Bạn có thể chọn continue hoặc break tùy thuộc vào cấu trúc file của bạn.
                    //    continue;
                    //}

                    try
                    {
                        Nganhang nganhang = new Nganhang();
                        nganhang.Stt = sttValue; // Gán STT đã đọc/parse
                        sttValue += 1;
                        // Ngay Giao dịch - Sử dụng hàm ParseAndCorrectDate() đã cung cấp
                        if (indexMaGD == 0)
                        {
                            try
                            {
                                var getsplit = ngayGDCellContent.Split(' ');
                                ngayGDCellContent = getsplit[0];
                                nganhang.Maso = getsplit[1];
                            }
                            catch(Exception ex)
                            {

                            }
                        }
                        else
                            nganhang.Maso = row.Cell(indexMaGD).GetString().Trim();
                        DateTime? parsedNgayGD = DateHelper.ParseAndCorrectDate(ngayGDCellContent);
                        if (parsedNgayGD.HasValue)
                        {
                            nganhang.NgayGD = parsedNgayGD.Value;
                        }
                        else
                        {
                            Console.WriteLine($"Cảnh báo: Không thể phân tích hoặc sửa ngày '{ngayGDCellContent}' tại hàng {rowIndex}. Bỏ qua dòng.");
                            continue; // Bỏ qua dòng nếu ngày không hợp lệ
                        }

                        // Diễn giải
                        nganhang.Diengiai = row.Cell(indexDiengiai).GetString().Trim();

                        // Mã Giao dịch
                       

                        // Số tiền rút (TK Có) và Số tiền gửi (TK Nợ)
                        double tkco = 0;
                        // Sử dụng InvariantCulture để đảm bảo việc phân tích số nhất quán, loại bỏ dấu phẩy/chấm
                        double.TryParse(row.Cell(indexTKCo).GetString().Replace(",", "").Replace(".", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out tkco);

                        double tkno = 0;
                        double.TryParse(row.Cell(indexTKNo).GetString().Replace(",", "").Replace(".", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out tkno);

                        // Logic xác định ThanhTien, TKCo, TKNo dựa trên số tiền rút hoặc gửi
                        string tknganhan = lblTKNganhang.Text.Split('-')[0].ToString();
                        //Kiểm tra mật định
                        string tkDoiung = "1111";
                        if (dtMatdinhnganhang.Rows.Count > 0)
                        {
                            foreach(DataRow dtRow in dtMatdinhnganhang.Rows)
                            {
                                string getNoidung = RemoveVietnameseDiacritics(dtRow["Noidung"].ToString().ToLower());
                                //Kiểm tra có chứa mặc định không
                                if (RemoveVietnameseDiacritics(nganhang.Diengiai).Contains(getNoidung))
                                {
                                    tkDoiung = dtRow["TK"].ToString();
                                }
                            }
                        }
                        if (tkco != 0)
                        {
                            nganhang.ThanhTien = tkco;
                            nganhang.TKCo = tknganhan;  // Tài khoản Có mẫu
                            nganhang.TKNo = tkDoiung; // Tài khoản Nợ mẫu
                        }
                        else if (tkno != 0) // Dùng else if để tránh trường hợp cả hai đều có giá trị (nếu có lỗi dữ liệu)
                        {
                            nganhang.ThanhTien2 = tkno;
                          
                            nganhang.TKCo = tkDoiung; // Tài khoản Có mẫu
                            nganhang.TKNo = tknganhan;  // Tài khoản Nợ mẫu
                        }
                        else
                        {
                            // Nếu cả tkco và tkno đều là 0, có thể đây không phải là dòng giao dịch tiền tệ
                            Console.WriteLine($"Cảnh báo: Dòng {rowIndex} không có giá trị Debit/Credit (TKCo/TKNo). Bỏ qua dòng.");
                            continue; // Bỏ qua dòng này
                        }

                        lstNganhan.Add(nganhang); // Thêm đối tượng Nganhang vào danh sách
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Lỗi khi đọc dữ liệu tại hàng {rowIndex}: {ex.Message}");
                        // Bạn có thể ghi log lỗi chi tiết hơn hoặc thông báo cho người dùng
                        continue; // Tiếp tục xử lý các dòng khác ngay cả khi có lỗi ở dòng hiện tại
                    }
                }

                // --- GÁN DỮ LIỆU VÀO ĐIỀU KHIỂN UI (GridControl) ---
                // Đảm bảo bdNganhang và gridControl3 đã được khởi tạo và có thể truy cập được
                // trong Form của bạn (ví dụ: là các thuộc tính công khai hoặc trường).
                if (bdNganhang != null)
                {
                    bdNganhang.DataSource = null; // Xóa nguồn dữ liệu cũ
                    bdNganhang.DataSource = lstNganhan; // Gán danh sách mới
                }
                if (gridControl3 != null)
                {
                    gridControl3.DataSource = bdNganhang; // GridControl lấy dữ liệu từ BindingSource
                    gridControl3.RefreshDataSource(); // Yêu cầu GridControl làm mới để hiển thị dữ liệu
                    xtraTabControl2.SelectedTabPageIndex = 3;
                }
            }
        }

        System.Data.DataTable dtMatdinhnganhang = new System.Data.DataTable();
        private void btnReadPDF_Click(object sender, EventArgs e)
        {
            if (lblTKNganhang.Text == "...")
            {
                XtraMessageBox.Show("Vui lòng chọn tài khoản ngân hàng!");
                return;
            }
            if (dtMatdinhnganhang.Rows.Count == 0)
            {
                string queryCheckVatTu = @"SELECT * FROM tbDinhdanhNganhang ";
                var parameterss = new OleDbParameter[]
                {
                 new OleDbParameter("?",null)
                   };
                dtMatdinhnganhang = ExecuteQuery(queryCheckVatTu, parameterss);
            }
           // ConvertImageTOexcel();
            string path = savedPath + @"\output.xlsx";
            ReadExcelBank(path);
        }

        private void gridView2_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            // Lấy gridView2
            GridView gridView2 = sender as GridView;

            // Lấy gridView1 thông qua Parent
            GridControl gridControl = gridView2.GridControl;
            GridView gridView1 = gridControl.MainView as GridView;

            // Lấy chỉ số dòng của gridView1
            int rowIndex = gridView1.FocusedRowHandle; // Thay vì CurrentRow.Index
            int rowIndex2 = gridView2.FocusedRowHandle; // Thay vì CurrentRow.Index 
            // Hoặc nếu bạn muốn sử dụng FieldName
            string fieldName = e.Column.FieldName;
            if (fieldName == "Soluong" || fieldName == "Dongia")
            {
                double soluong = (double)gridView2.GetRowCellValue(rowIndex2, "Soluong"); // Thay "ColumnName" bằng tên cột thực tế
                double dongia = (double)gridView2.GetRowCellValue(rowIndex2, "Dongia"); // Thay "ColumnName" bằng tên cột thực tế
                double Thanhtien = soluong * dongia;
                gridView2.SetRowCellValue(rowIndex2, "TTien", Thanhtien);
            }
            // Kiểm tra nếu chỉ số dòng hợp lệ
            if (rowIndex >= 0)
            {
                // Thực hiện các hành động với row của gridView1
                int rowData = (int)gridView1.GetRowCellValue(rowIndex, "ID"); // Thay "ColumnName" bằng tên cột thực tế
                string SoHieu = gridView2.GetRowCellValue(rowIndex2, "SoHieu").ToString();

                string queryCheckVatTu = @"SELECT * FROM tbimportdetail WHERE  ID = ?";
                var parameterss = new OleDbParameter[]
                { 
                 new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "ID").ToString())
                   };
                var kq = ExecuteQuery(queryCheckVatTu, parameterss);
                if (kq.Rows.Count > 0)
                {
                    string query = @"UPDATE tbimportdetail SET SoHieu=?, SoLuong=?, DonGia=?, DVT=?, Ten=?, TKNo=?,TKCo=?, TTien=? WHERE ID=?";
                    var parameters = new OleDbParameter[]
                      {
            new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "SoHieu").ToString()),
            new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "Soluong").ToString()),
            new OleDbParameter("?",  gridView2.GetRowCellValue(rowIndex2, "Dongia").ToString()),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(gridView2.GetRowCellValue(rowIndex2, "DVT").ToString())),
            new OleDbParameter("?",Helpers.ConvertUnicodeToVni(gridView2.GetRowCellValue(rowIndex2, "Ten").ToString())),
            new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "TKNo").ToString()),
             new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "TKCo").ToString()),
              new OleDbParameter("?",  gridView2.GetRowCellValue(rowIndex2, "TTien").ToString()),
               new OleDbParameter("?",gridView2.GetRowCellValue(rowIndex2, "ID").ToString()),
                      };
                    int rowsAffected = ExecuteQueryResult(query, parameters);
                }
            }

        }

        private void gridView4_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            // Lấy gridView2
            GridView gridView2 = sender as GridView;

            // Lấy gridView1 thông qua Parent
            GridControl gridControl = gridView2.GridControl;
            GridView gridView1 = gridControl.MainView as GridView;

            // Lấy chỉ số dòng của gridView1
            int rowIndex = gridView1.FocusedRowHandle; // Thay vì CurrentRow.Index
            int rowIndex2 = gridView2.FocusedRowHandle; // Thay vì CurrentRow.Index
            // Kiểm tra nếu chỉ số dòng hợp lệ
            if (rowIndex >= 0)
            {
                // Thực hiện các hành động với row của gridView1
                int rowData = (int)gridView1.GetRowCellValue(rowIndex, "ID"); // Thay "ColumnName" bằng tên cột thực tế
                string SoHieu = gridView2.GetRowCellValue(rowIndex2, "SoHieu").ToString();
                string fieldName = e.Column.FieldName;
                if (fieldName == "Soluong" || fieldName == "Dongia")
                {
                    double soluong = (double)gridView2.GetRowCellValue(rowIndex2, "Soluong"); // Thay "ColumnName" bằng tên cột thực tế
                    double dongia = (double)gridView2.GetRowCellValue(rowIndex2, "Dongia"); // Thay "ColumnName" bằng tên cột thực tế
                    double Thanhtien = soluong * dongia;
                    gridView2.SetRowCellValue(rowIndex2, "TTien", Thanhtien);
                } 
                string queryCheckVatTu = @"SELECT * FROM tbimportdetail WHERE  ID = ?";
                var parameterss = new OleDbParameter[]
                {
                 new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "ID").ToString())
                   };
                var kq = ExecuteQuery(queryCheckVatTu, parameterss);
                if (kq.Rows.Count > 0)
                {
                    string query = @"UPDATE tbimportdetail SET SoHieu=?, SoLuong=?, DonGia=?, DVT=?, Ten=?, TKNo=?,TKCo=?, TTien=? WHERE ID=?";
                    var parameters = new OleDbParameter[]
                      {
            new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "SoHieu").ToString()),
            new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "Soluong").ToString()),
            new OleDbParameter("?",  gridView2.GetRowCellValue(rowIndex2, "Dongia").ToString()),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(gridView2.GetRowCellValue(rowIndex2, "DVT").ToString())),
            new OleDbParameter("?",Helpers.ConvertUnicodeToVni(gridView2.GetRowCellValue(rowIndex2, "Ten").ToString())),
            new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "TKNo").ToString()),
             new OleDbParameter("?", gridView2.GetRowCellValue(rowIndex2, "TKCo").ToString()),
              new OleDbParameter("?",  gridView2.GetRowCellValue(rowIndex2, "TTien").ToString()),
               new OleDbParameter("?",gridView2.GetRowCellValue(rowIndex2, "ID").ToString()),
                      };
                    int rowsAffected = ExecuteQueryResult(query, parameters);
                }
            }

        }

        private void gridView1_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
            {
             if (e.Column.FieldName == "Index" && e.IsGetData)
            {
                e.Value = e.ListSourceRowIndex + 1; // Gán giá trị số thứ tự
            }
        }

        private void gridView3_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
           if (e.Column.FieldName == "Index")
            {
                if (e.IsGetData) // Kiểm tra xem có phải đang lấy dữ liệu không
                {
                    e.Value = e.ListSourceRowIndex + 1; // Gán số thứ tự
                }
            }
        }

        private void gridView3_MouseDown(object sender, MouseEventArgs e)
        {
            
        }

        private void gridView4_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Clicks == 2 && e.Button == MouseButtons.Left)
            {
                DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;
                string currentColumnName = gridView.FocusedColumn.FieldName;
                int currentRowHandle = gridView.FocusedRowHandle;
                object cellValue = gridView.GetRowCellValue(currentRowHandle, currentColumnName);

                if (currentColumnName.ToString() == "SoHieu")
                {
                    frmHangHoa frmHangHoa = new frmHangHoa();
                    frmHangHoa.frmMain = this;

                    VatTu vatTu = new VatTu();
                    vatTu.SoHieu = cellValue.ToString();

                    string querydinhdanh = @"SELECT * FROM Vattu WHERE SoHieu = ?";
                    var checkSH = ExecuteQuery(querydinhdanh, new OleDbParameter("?", vatTu.SoHieu));
                    //Nếu chưa có, tìm hết danh sách
                    if (checkSH.Rows.Count == 0)
                    {
                        foreach (var item in lstImportRa)
                        {
                            foreach (var it in item.fileImportDetails)
                            {
                                if (it.SoHieu == vatTu.SoHieu)
                                {
                                    lstrowSohieu.Add(it.ID);
                                }
                            }
                        }
                    }

                    var tenvattu =gridView.GetRowCellValue(currentRowHandle, "Ten").ToString();
                    vatTu.TenVattu = tenvattu;
                    var dvt = gridView.GetRowCellValue(currentRowHandle, "DVT").ToString();
                    vatTu.DonVi = dvt;
                    frmHangHoa.dtoVatTu = vatTu;
                    if (checkSH.Rows.Count > 0)
                        currentselectId = int.Parse(checkSH.Rows[0]["MaPhanLoai"].ToString());
                    else
                        currentselectId = 0;
                    frmHangHoa.ShowDialog();


                    if (!string.IsNullOrEmpty(hiddenValue) && frmHangHoa.isChange)
                    {
                        gridView.SetRowCellValue(currentRowHandle, currentColumnName, hiddenValue);
                        gridView.SetRowCellValue(currentRowHandle, "DVT", hiddenValue2);
                        gridView.SetRowCellValue(currentRowHandle, "Ten", hiddenValue3);
                        foreach (var item in lstrowSohieu)
                        {
                            foreach (var it in lstImportRa)
                            {
                                var findchild = it.fileImportDetails.Where(m => m.ID == item).FirstOrDefault();
                                if (findchild != null)
                                {
                                    findchild.SoHieu = hiddenValue;
                                    findchild.DVT = hiddenValue2;
                                    findchild.Ten = hiddenValue3;
                                    //Update vào database
                                    var query = @"UPDATE tbimportdetail SET  SoHieu=?,DVT=?,Ten=? where ID=? ";
                                    var parameters = new OleDbParameter[]
                                    {
                                        new OleDbParameter("?", hiddenValue),
                                        new OleDbParameter("?",  Helpers.ConvertUnicodeToVni(hiddenValue2)) ,
                                        new OleDbParameter("?", Helpers.ConvertUnicodeToVni(hiddenValue3)),
                                        new OleDbParameter("?", findchild.ID)
                                    };
                                    var rowsAffected = ExecuteQueryResult(query, parameters);
                                }
                            }
                        }
                        lstrowSohieu = new List<int>();
                    }

                }
            }

        }

        private void gridView2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Clicks == 2 && e.Button == MouseButtons.Left)
            {
                DevExpress.XtraGrid.Views.Grid.GridView gridView = sender as DevExpress.XtraGrid.Views.Grid.GridView;
                string currentColumnName = gridView.FocusedColumn.FieldName;
                int currentRowHandle = gridView.FocusedRowHandle;
                object cellValue = gridView.GetRowCellValue(currentRowHandle, currentColumnName);
                if (currentColumnName.ToString() == "SoHieu")
                {
                    frmHangHoa frmHangHoa = new frmHangHoa();
                    frmHangHoa.frmMain = this;

                   VatTu vatTu = new VatTu();
                    vatTu.SoHieu = cellValue.ToString();
                    //Kiểm tra xem có phải so hiệu tự tạo 
                    string querydinhdanh = @"SELECT * FROM Vattu WHERE SoHieu = ?";
                    var checkSH = ExecuteQuery(querydinhdanh, new OleDbParameter("?", vatTu.SoHieu));
                    //Nếu chưa có, tìm hết danh sách
                    if (checkSH.Rows.Count == 0)
                    {
                        foreach (var item in lstImportVao)
                        {
                            foreach (var it in item.fileImportDetails)
                            {
                                if (it.SoHieu == vatTu.SoHieu)
                                {
                                    lstrowSohieu.Add(it.ID);
                                }
                            }
                        }
                    }

                    var tenvattu = gridView.GetRowCellValue(currentRowHandle, "Ten").ToString();
                    vatTu.TenVattu = tenvattu;
                    var dvt = gridView.GetRowCellValue(currentRowHandle, "DVT").ToString();
                    vatTu.DonVi = dvt;
                    frmHangHoa.dtoVatTu = vatTu;
                    if (checkSH.Rows.Count > 0)
                        currentselectId = int.Parse(checkSH.Rows[0]["MaPhanLoai"].ToString());
                    else
                        currentselectId = 0;
                    frmHangHoa.ShowDialog();



                    if (!string.IsNullOrEmpty(hiddenValue) && frmHangHoa.isChange)
                    {
                        gridView.SetRowCellValue(currentRowHandle, currentColumnName, hiddenValue);
                        gridView.SetRowCellValue(currentRowHandle, "DVT", hiddenValue2);
                        gridView.SetRowCellValue(currentRowHandle, "Ten", hiddenValue3);
                        //Fill lai cho lstrowSohieu
                        foreach (var item in lstrowSohieu)
                        {
                            foreach (var it in lstImportVao)
                            {
                                var findchild = it.fileImportDetails.Where(m => m.ID == item).FirstOrDefault();
                                if (findchild != null)
                                {
                                    findchild.SoHieu = hiddenValue;
                                    findchild.DVT = hiddenValue2;
                                    findchild.Ten = hiddenValue3;
                                    //Update vào database
                                    var query = @"UPDATE tbimportdetail SET  SoHieu=?,DVT=?,Ten=? where ID=? ";
                                    var parameters = new OleDbParameter[]
                                    {
                                                new OleDbParameter("?", hiddenValue),
                                                new OleDbParameter("?",  Helpers.ConvertUnicodeToVni(hiddenValue2)) ,
                                                new OleDbParameter("?", Helpers.ConvertUnicodeToVni(hiddenValue3)),
                                                new OleDbParameter("?", findchild.ID)
                                    };
                                    var rowsAffected = ExecuteQueryResult(query, parameters);
                                }
                            }
                        }
                        lstrowSohieu = new List<int>();
                    }

                }
            }
        }

        private void btnMatdinhnganhang_Click(object sender, EventArgs e)
        {
            frmMatdinhNganHang frmMatdinhNganHang = new frmMatdinhNganHang();
            frmMatdinhNganHang.Show();
        }

        private void btnImportChungtunganhang_Click(object sender, EventArgs e)
        {
            if (xtraTabControl2.SelectedTabPageIndex == 2)
            {
                foreach (var item in lstNganhan)
                {
                    var query = @"INSERT INTO tbNganhang (SHDon, NgayGD, DienGiai, TongTien,TongTien2, TKNo,TKCo,Status) VALUES (?, ?, ?, ?, ?,?,?,?)";
                    var parameters = new OleDbParameter[]
                       {
            new OleDbParameter("?", item.Maso),
            new OleDbParameter("?", item.NgayGD),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(item.Diengiai)),
            new OleDbParameter("?", item.ThanhTien),
               new OleDbParameter("?", item.ThanhTien2),
            new OleDbParameter("?", item.TKNo),
            new OleDbParameter("?", item.TKCo),
              new OleDbParameter("?","0"),
                       };
                    int rowsAffected = ExecuteQueryResult(query, parameters);
                }
                this.Close();
            }

        }

        private void btnChontknganhang_Click(object sender, EventArgs e)
        {
            frmTaikhoan frmTaikhoan = new frmTaikhoan();
            frmTaikhoan.frmMain = this;
            frmTaikhoan.ShowDialog();
            lblTKNganhang.Text = tknh;
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            frmTaiCoQuanThue frmTaiCoQuanThue = new frmTaiCoQuanThue();
            username = txtuser.Text;
            pasword = txtpass.Text;
            type = chkDauvao.Checked ? 1 : 2;
            dtFrom = dtTungay.DateTime;
            dtTo = dtDenngay.DateTime;
            frmTaiCoQuanThue.frmMain = this;
            frmTaiCoQuanThue.Show();
        }

        //private void btnScanCmera_Click(object sender, EventArgs e)
        //{
        //    frmCamera camera = new frmCamera();
        //    camera.ShowDialog();

        //}

        public static string NormalizeVietnameseString(string input)
        {
            //Bỏ đi ký tự đặc biệt đầu chữ
            input = RemoveLeadingSpecialCharacters(input);
            //Bỏ đi tab
            input= input.Replace("\t", ""); // Thay thế ký tự tab bằng chuỗi rỗng
            input = input.Normalize(NormalizationForm.FormC);

            if (string.IsNullOrEmpty(input))
                return input;
            return input;

            // Bước 1: Chuẩn hóa dấu huyền/sắc/hỏi/ngã/nặng tách rời
            var replacements = new Dictionary<string, string>
    {
        // Dấu huyền (`) tách rời (U+0300)
        {"à", "à"}, {"ằ", "ằ"}, {"ầ", "ầ"}, {"è", "è"}, {"ề", "ề"},
        {"ì", "ì"}, {"ò", "ò"}, {"ồ", "ồ"}, {"ờ", "ờ"}, {"ù", "ù"}, {"ừ", "ừ"}, {"ỳ", "ỳ"},
        
        // Dấu sắc (') tách rời (U+0301)
        {"á", "á"}, {"ắ", "ắ"}, {"ấ", "ấ"}, {"é", "é"}, {"ế", "ế"},
        {"í", "í"}, {"ó", "ó"}, {"ố", "ố"}, {"ớ", "ớ"}, {"ú", "ú"}, {"ứ", "ứ"}, {"ý", "ý"},
        
        // Dấu hỏi (?) tách rời (U+0309)
        {"ả", "ả"}, {"ẳ", "ẳ"}, {"ẩ", "ẩ"}, {"ẻ", "ẻ"}, {"ể", "ể"},
        {"ỉ", "ỉ"}, {"ỏ", "ỏ"}, {"ổ", "ổ"}, {"ở", "ở"}, {"ủ", "ủ"}, {"ử", "ử"}, {"ỷ", "ỷ"},
        
        // Dấu ngã (~) tách rời (U+0303)
        {"ã", "ã"}, {"ẵ", "ẵ"}, {"ẫ", "ẫ"}, {"ẽ", "ẽ"}, {"ễ", "ễ"},
        {"ĩ", "ĩ"}, {"õ", "õ"}, {"ỗ", "ỗ"}, {"ỡ", "ỡ"}, {"ũ", "ũ"}, {"ữ", "ữ"}, {"ỹ", "ỹ"},
        
        // Dấu nặng (.) tách rời (U+0323)
        {"ạ", "ạ"}, {"ặ", "ặ"}, {"ậ", "ậ"}, {"ẹ", "ẹ"}, {"ệ", "ệ"},
        {"ị", "ị"}, {"ọ", "ọ"}, {"ộ", "ộ"}, {"ợ", "ợ"}, {"ụ", "ụ"}, {"ự", "ự"}, {"ỵ", "ỵ"}
    };

            // Bước 2: Thay thế tất cả các trường hợp
            foreach (var replacement in replacements)
            {
                input = input.Replace(replacement.Key, replacement.Value);
            }

            return input;
        }
        #endregion
    }
}
