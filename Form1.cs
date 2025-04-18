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
namespace SaovietTax
{
    public partial class frmMain : DevExpress.XtraEditors.XtraForm
    {
        #region  Khai báo
        string savedPath = "";
        string dbPath = "";
        public static int Id = 1;
        bool isSetuppath = false;
        bool isetupDbpath = false;
        string password, connectionString;
        public string MasterMST = "3502264379";
        public DataTable result { get; set; }
        private BindingList<FileImport> people = new BindingList<FileImport>();
        System.Windows.Forms.BindingSource bindingSource = new System.Windows.Forms.BindingSource();
        public class FileImportDetail
        {
            public string Ten { get; set; }
            public int ParentId { get; set; }
            public string SoHieu { get; set; }
            public double Soluong { get; set; }
            public double Dongia { get; set; }
            public string DVT { get; set; }
            public string MaCT { get; set; }
            public string TKNo { get; set; }

            public FileImportDetail(string ten, int parentId, string soHieu, double soluong, double dongia, string dVT, string maCT, string tkNo)
            {
                Ten = ten;
                ParentId = parentId;
                SoHieu = soHieu;
                Soluong = soluong;
                Dongia = dongia;
                DVT = dVT;
                MaCT = maCT;
                TKNo = tkNo;
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
            public int Vat { get; set; }
            public int Type { get; set; }

            public string SoHieuTP { get; set; }
            public List<FileImportDetail> fileImportDetails;
            public FileImport(string path, string shdon, string khhdon, DateTime nlap, string ten, string noidung, string tkno, string tkco, int tkthue, string mst, double tongTien, int vat, int type, string tenTP)
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
            }

        }
        #endregion# 
        #region loadData
        public frmMain()
        {
            InitializeComponent();
        }

        private void LoadDataGridview()
        {
            bindingSource.DataSource = people;
            gridControl1.DataSource = bindingSource;
            GridView gridView = gridControl1.MainView as GridView;

            if (gridView != null)
            {
                // Kích hoạt kiểu dáng hàng chẵn và lẻ
                gridView.OptionsView.EnableAppearanceEvenRow = true;
                gridView.OptionsView.EnableAppearanceOddRow = true;

                // Thiết lập màu sắc cho hàng chẵn
                gridView.Appearance.EvenRow.BackColor = System.Drawing.Color.LightCyan; // Màu nền cho hàng chẵn
                gridView.Appearance.EvenRow.ForeColor = System.Drawing.Color.Black; // Màu chữ cho hàng chẵn

                // Thiết lập màu sắc cho hàng lẻ
                gridView.Appearance.OddRow.BackColor = System.Drawing.Color.White; // Màu nền cho hàng lẻ
                gridView.Appearance.OddRow.ForeColor = System.Drawing.Color.Black; // Màu chữ cho hàng lẻ

                gridView.CellValueChanged += GridView_CellValueChanged;

            }
        }
        private void GridView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            // Lấy thông tin về hàng và cột của ô đã thay đổi
            int rowHandle = e.RowHandle;
            string columnName = e.Column.FieldName; // Tên cột
            if (columnName != "TKNo" && columnName != "TKCo")
                return;
            object newValue = e.Value; // Giá trị mới

            string query = "SELECT * FROM HeThongTK WHERE SoHieu = ?";

            if (!string.IsNullOrEmpty(newValue.ToString()))
            {
                // Tạo mảng tham số với giá trị cho câu lệnh SQL
                OleDbParameter[] parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", newValue),
                };
                var kq = ExecuteQuery(query, parameters);
                if (kq.Rows.Count == 0 && !newValue.ToString().Contains("|"))
                {
                    GridView gridView = gridControl1.MainView as GridView;
                    gridView.SetRowCellValue(rowHandle, e.Column, "");
                    XtraMessageBox.Show("Số tài khoản không tồn tại trong hệ thống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        static bool TableExists(OleDbConnection connection, string tableName)
        {
            try
            {
                // Kiểm tra sự tồn tại của bảng
                DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
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
            TKNo TEXT
        );";

            using (OleDbCommand command = new OleDbCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
        private void InitDB()
        {
            // Đường dẫn đến cơ sở dữ liệu Access và mật khẩu
            //dbPath = @"C:\S.T.E 25\S.T.E 25\DATA\KT2025.mdb"; // Thay đổi đường dẫn này
            dbPath = ConfigurationManager.AppSettings["dbpath"];
            //dbPath = "sadsa";


            string filePath = ConfigurationManager.AppSettings["dbpath"];
            if (string.IsNullOrEmpty(filePath))
                return;
            // Đọc toàn bộ nội dung tệp
            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            //connectionString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            // connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database";
            //connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\S.T.E 25\S.T.E 25\DATA\importData.accdb;Persist Security Info=False";
            MessageBox.Show(connectionString);
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                MessageBox.Show("Kết nối database thành công");
                string tableName = "tbimport";
                string tableNamedetail = "tbimportdetail";
                string tableDinhdanh = "tbDinhdanhtaikhoan";
                // Kiểm tra xem bảng đã tồn tại hay không
                if (!TableExists(connection, tableDinhdanh))
                {
                    // Tạo bảng nếu chưa tồn tại
                    CreateTableDinhDanh(connection, tableDinhdanh);
                    Console.WriteLine($"Bảng '{tableDinhdanh}' đã được tạo thành công.");
                }
                else
                {
                    Console.WriteLine($"Bảng '{tableName}' đã tồn tại.");
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
                    Console.WriteLine($"Bảng '{tableNamedetail}' đã tồn tại.");
                }
            }

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
            if (!string.IsNullOrEmpty(savedPath))
            {
                txtPath.Text = savedPath;
                txtPath.Enabled = false;
                isSetuppath = true;
            }
            else
            {
                XtraMessageBox.Show("Vui lòng thiết lập đường dẫn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                isSetuppath = false;
                return;
            }
            dbPath = ConfigurationManager.AppSettings["dbpath"];
            if (!string.IsNullOrEmpty(dbPath))
            {
                txtdbPath.Text = dbPath;
                txtdbPath.Enabled = false;
                isetupDbpath = true;
            }
            else
            {
                XtraMessageBox.Show("Vui lòng thiết lập đường dẫn cơ sở dữ liệu!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                isetupDbpath = false;
                return;
            }
            //txtUsername.Text = ConfigurationManager.AppSettings["username"];
            //txtPassword.Text = ConfigurationManager.AppSettings["password"];
            txtuser.Text = ConfigurationManager.AppSettings["username"];
            txtpass.Text = ConfigurationManager.AppSettings["password"];
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
                    DataTable schemaTable = reader.GetSchemaTable();
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
        private void frmMain_Load(object sender, EventArgs e)
        {
           
            InitData();
            InitDB();
            CheckDB();
            ControlsSetup();
            LoadDataDinhDanh();
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
        private void LoadXmlFiles(string path)
        {
            progressBarControl1.EditValue = 0;
            Application.DoEvents();
            people = new BindingList<FileImport>();
            path = path + "\\HDDauVao";
            int fromMonth = int.Parse(comboBoxEdit1.SelectedItem.ToString()); // Thay đổi theo tháng bắt đầu (ví dụ: 3 cho tháng 3)
            int toMonth = int.Parse(comboBoxEdit2.SelectedItem.ToString());   // Thay đổi theo tháng kết thúc (ví dụ: 7 cho tháng 7)
            // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
            var files = Directory.EnumerateFiles(path, "*.xml", SearchOption.AllDirectories)
                .Where(file => IsFileInMonthRange(file, path, fromMonth, toMonth));
            int countXml = files.Count();
            Dictionary<string, string> lstHodpn = new Dictionary<string, string>();
            //Lấy danh sách hóa đơn để kiểm tra cho excel

            // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
            var excelFiles = Directory.EnumerateFiles(path, "*.xlsx", SearchOption.AllDirectories)
                .Where(file => IsFileInMonthRange(file, path, fromMonth, toMonth)).ToList(); // Kiểm tra xem file có nằm trong khoảng tháng
            int rowCount = 0;

            //Kiểm tra xem có bao nhieu dòng dữ liệu trong Excel
            for (int j = 0; j < excelFiles.Count; j++)
            {
                using (var workbook = new XLWorkbook(excelFiles[0]))
                {
                    // Lấy worksheet đầu tiên
                    var worksheet = workbook.Worksheet(1); // Hoặc bạn có thể dùng tên worksheet như worksheet = workbook.Worksheet("Sheet1");
                                                           // Lấy giá trị của ô A6

                    var currentCell = worksheet.Cell("A6"); // Bắt đầu từ ô A6

                    // Kiểm tra các ô bắt đầu từ A6 cho đến khi gặp ô trống
                    while (!currentCell.IsEmpty())
                    {
                        rowCount++; // Tăng số dòng
                        currentCell = currentCell.Worksheet.Row(currentCell.Address.RowNumber + 1).Cell("A"); // Chuyển xuống ô bên dưới
                    }

                }
            }

            int countExcel = 0;
            if (rowCount > 0)
                countExcel = rowCount - 1;

            totalCount = countXml + countExcel;
            lblSofiles.Text = totalCount.ToString();
            //foreach (string file in files)
            //{
            //    progressPercentage = (filesLoaded * 100) / totalCount;
            //    filesLoaded += 1;
            //    progressBarControl1.EditValue = progressPercentage;
            //} 
            foreach (string file in files)
            {

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
                var kq = ExecuteQuery(query, parameters);
                if (kq.Rows.Count > 0)
                {
                    continue;
                }
                if (people.Any(m => m.SHDon.Contains(SHDon) && m.KHHDon == KHHDon))
                {
                    continue;
                }

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


                XmlNode nBanNode = ndhDonNode.SelectSingleNode("NBan");
                if (nBanNode != null)
                {
                    ten = nBanNode.SelectSingleNode("Ten")?.InnerText;
                    mst = nBanNode.SelectSingleNode("MST")?.InnerText;
                    if (mst == MasterMST)
                    {
                        XmlNode nMuaNode = ndhDonNode.SelectSingleNode("NMua");
                        if (nBanNode != null)
                        {
                            ten = nMuaNode.SelectSingleNode("Ten")?.InnerText;
                            mst = nMuaNode.SelectSingleNode("MST")?.InnerText;
                        }
                    }
                }

                if (nTSuat != null)
                {
                    for (int i = 0; i < nTSuat.Count; i++)
                    {
                        XmlNode item = nTSuat[i];
                        if (item.InnerText != "KKKNT" && item.InnerText != "KCT")
                            Vat = int.Parse(item.InnerText.Replace("%", ""));
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
                                Thanhtien = double.Parse(nThTien[i].InnerText);
                            else
                                Thanhtien = 0;
                        }
                    }
                    else
                    {
                        XmlNode TgTTTBSo = root.SelectSingleNode("//TToan//TgTTTBSo");
                        Thanhtien = double.Parse(TgTTTBSo.InnerText);
                    }

                }
                else
                {
                    //Kiểm tra tiếp
                    XmlNode TgTTTBSo = root.SelectSingleNode("//TToan//TgTTTBSo");
                    Thanhtien = double.Parse(TgTTTBSo.InnerText);
                }

                //Kiểm tra thêm mới khách hàng
                string querykh = @" SELECT TOP 1 *  FROM KhachHang As kh
WHERE kh.MST = ?"; // Sử dụng ? thay cho @mst trong OleDb
                DataTable result = ExecuteQuery(querykh, new OleDbParameter("?", mst));
                if (result.Rows.Count == 0)
                {
                    string diachi = nBanNode.SelectSingleNode("DChi")?.InnerText;
                    var Sohieu = GetLastFourDigits(mst);
                    ten = Helpers.ConvertUnicodeToVni(ten);
                    diachi = Helpers.ConvertUnicodeToVni(diachi);
                    query = @" SELECT TOP 1 *  FROM KhachHang As kh
WHERE kh.SoHieu = ?";
                    DataTable result2 = ExecuteQuery(query, new OleDbParameter("?", Sohieu));
                    if (result2.Rows.Count > 0)
                        Sohieu = "0" + Sohieu;
                    if (string.IsNullOrEmpty(diachi))
                    {
                        diachi = Helpers.ConvertUnicodeToVni("Bô sung địa chỉ");
                    }
                    InitCustomer(2, Sohieu, ten, diachi, mst);
                }

                query = @" SELECT TOP 1 *  FROM KhachHang AS kh  
INNER JOIN HoaDon AS hd ON kh.Maso = hd.MaKhachHang    
WHERE kh.MST = ?  
ORDER BY hd.MaSo DESC"; // Sử dụng ? thay cho @mst trong OleDb
                result = ExecuteQuery(query, new OleDbParameter("?", mst));
                if (result.Rows.Count > 10)
                {
                    SoHD = result.Rows[0]["SoHD"].ToString();

                    query = @"Select top 2 * from ChungTu 
where SoHieu = ?
ORDER BY  MaSo DESC";
                    result = ExecuteQuery(query, new OleDbParameter("?", SoHD));
                    var index = 0;
                    if (result.Rows.Count > 0)
                    {
                        foreach (DataRow row in result.Rows)
                        {
                            if (index == 0)
                            {
                                TkThue = int.Parse(row["MaTKNo"].ToString());  // Giả sử có cột "MaSo"; 
                            }
                            if (index == 1)
                            {
                                TkNo = int.Parse(row["MaTKNo"].ToString());  // Giả sử có cột "MaSo"; 
                                TkCo = int.Parse(row["MaTKCo"].ToString());  // Giả sử có cột "MaSo"; 
                                diengiai = Helpers.ConvertVniToUnicode(row["DienGiai"].ToString());
                            }
                            // Lấy giá trị từ cột cụ thể trong hàng hiện tại

                            index += 1;
                        }
                    }
                    // Tra cứu từ bảng HeThongTK
                    query = @"Select   * from HeThongTK where MaTC = ?";
                    result = ExecuteQuery(query, new OleDbParameter("?", TkNo));
                    if (result.Rows.Count > 0)
                    {
                        TkNo = int.Parse(result.Rows[0]["SoHieu"].ToString());

                        query = @"Select   * from HeThongTK where MaTC = ?";
                        if (TkCo > 0)
                        {
                            result = ExecuteQuery(query, new OleDbParameter("?", TkCo));
                            TkCo = int.Parse(result.Rows[0]["SoHieu"].ToString());
                        }


                        query = @"Select   * from HeThongTK where MaTC = ?";
                        if (TkThue > 0)
                        {
                            result = ExecuteQuery(query, new OleDbParameter("?", TkThue));
                            TkThue = int.Parse(result.Rows[0]["SoHieu"].ToString());
                        }


                    }
                    else
                    {
                        //TkNo = 0;
                        //TkCo = 1111;
                        //TkThue = 1331;
                    }
                }
                else
                {

                }
                if (TkThue == 0)
                {

                }
                //Add detail
                var hhdVuList = xmlDoc.SelectNodes("//HHDVu");
                //Mật định tài khoản 
                //Kiểm tra Đã tồn tại số hóa đơn và số hiệu
                if (!people.Any(m => m.SHDon.Contains(SHDon) && m.KHHDon == KHHDon))
                {
                    people.Add(new FileImport(file, SHDon, KHHDon, NLap, ten, diengiai, TkNo.ToString(), TkCo.ToString(), TkThue, mst, Thanhtien, Vat, 1, ""));
                }
                for (int i = 0; i < hhdVuList.Count; i++)
                {
                    try
                    {
                        string tkno = "";
                        string mct = "";
                        if (hhdVuList[i].SelectSingleNode("DVTinh") != null && !string.IsNullOrEmpty(hhdVuList[i].SelectSingleNode("DVTinh").ToString()))
                        {
                            var THHDVu = hhdVuList[i].SelectSingleNode("THHDVu").InnerText;
                            var DVTinh = hhdVuList[i].SelectSingleNode("DVTinh").InnerText;

                            var SLuong = hhdVuList[i].SelectSingleNode("SLuong").InnerText;
                            var DGia = "";
                            if (hhdVuList[i].SelectSingleNode("DGia") != null)
                                DGia = hhdVuList[i].SelectSingleNode("DGia").InnerText;
                            else
                                DGia = "0";
                            string newName = Helpers.ConvertUnicodeToVni(NormalizeVietnameseString(THHDVu.Trim()));
                            //Kiểm tra trong database xem có sản phẩm chưa, nếu chưa có thì thêm mới
                            query = @"SELECT * FROM Vattu 
WHERE LCase(TenVattu) = LCase(?) AND LCase(DonVi) = LCase(?)";

                            //int rs = (int)ExecuteQuery(query, new OleDbParameter("?", "SAdsd")).Rows[0][0];
                            var getdata = ExecuteQuery(query, new OleDbParameter("?", newName.ToLower()), new OleDbParameter("?", Helpers.ConvertUnicodeToVni(DVTinh).ToLower()));
                            //Kiểm tra thêm trong list
                            var checkold = people.LastOrDefault().fileImportDetails.Where(m => m.Ten == newName && m.DVT == Helpers.ConvertUnicodeToVni(DVTinh)).FirstOrDefault();
                            string sohieu = "";
                            if (getdata.Rows.Count == 0)
                            {
                                if (checkold == null)
                                    sohieu = GenerateResultString(NormalizeVietnameseString(THHDVu.Trim()));
                                else
                                    sohieu = checkold.SoHieu;
                            }
                            else
                                sohieu = getdata.Rows[0]["SoHieu"].ToString();

                            //Gán giá trị cho các giá trị ""
                            DGia = !string.IsNullOrEmpty(DGia) ? DGia : "0";
                            SLuong = !string.IsNullOrEmpty(SLuong) ? SLuong : "0";
                            //Thiết lập MÃ ctrinh2 và tkno cho detail 
                            FileImportDetail fileImportDetail = new FileImportDetail(newName, people.LastOrDefault().ID, sohieu, double.Parse(SLuong), double.Parse(DGia), Helpers.ConvertUnicodeToVni(DVTinh), mct, tkno);
                            people.LastOrDefault().fileImportDetails.Add(fileImportDetail);
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
                                    FileImportDetail fileImportDetail = new FileImportDetail(THHDVu, people.LastOrDefault().ID, "711", 1, double.Parse(ThTien), "Exception", "", "");
                                    people.LastOrDefault().TKNo = "3311";
                                    people.LastOrDefault().TKCo = "711";
                                    people.LastOrDefault().TkThue = 1331;
                                    people.LastOrDefault().Noidung = "Chiếc khấu thương mại";
                                    people.LastOrDefault().fileImportDetails.Add(fileImportDetail);
                                }
                                else
                                {
                                    FileImportDetail fileImportDetail = new FileImportDetail(THHDVu, people.LastOrDefault().ID, "711", 0, double.Parse(ThTien), "Exception", "", "");
                                    people.LastOrDefault().fileImportDetails.Add(fileImportDetail);
                                }

                            }
                            else
                            {
                                var ThTien = hhdVuList[i].SelectSingleNode("ThTien")?.InnerText;
                                if (ThTien == null)
                                    ThTien = hhdVuList[i].SelectSingleNode("THTien")?.InnerText;
                                if (ThTien != null && double.Parse(ThTien) > 0)
                                {
                                    FileImportDetail fileImportDetail = new FileImportDetail(THHDVu, people.LastOrDefault().ID, "6422", 0, double.Parse(ThTien), "Exception", "", "");
                                    people.LastOrDefault().fileImportDetails.Add(fileImportDetail);
                                }

                            }

                        }
                    }
                    catch (Exception ex)
                    {

                    }
                    //Kiểm tra nếu ko có con thì tk cha sẽ là 6240
                    if (people.LastOrDefault().fileImportDetails.Count == 0)
                    {
                        people.LastOrDefault().TKNo = "6422";
                    }
                }

            }
            //Trường hợp không đủ info
            //Th1 có 1 sản phẩm và ko có đơn vị tính
            foreach (var item in people)
            {
                if (item.fileImportDetails.Count == 1 && string.IsNullOrEmpty(item.fileImportDetails[0].DVT))
                {
                    item.TKNo = "6422";
                    item.TKCo = "1111";
                    item.TkThue = 1331;
                }
            }
            //Lấy danh sách định danh
            string querydinhdanh = @" SELECT *  FROM tbDinhdanhtaikhoan"; // Sử dụng ? thay cho @mst trong OleDb

            result = ExecuteQuery(querydinhdanh, new OleDbParameter("?", ""));
            //Kiểm tra lại lại mã với Định danh
            foreach (var item in people)
            {
              
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
                            foreach (string condition in conditions)
                            {
                                string[] parts = Regex.Split(condition, @"([><=%]+)"); // Vẫn giữ % để linh hoạt nếu cần
                                if (parts.Length == 3)
                                {
                                    string key = parts[0];
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
                                item.TKNo = row["TKNo"].ToString();
                                item.TKCo = row["TKCo"].ToString();
                                item.TkThue = int.Parse(row["TkThue"].ToString());
                                if (string.IsNullOrEmpty(item.Noidung))
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
            }
            //Nếu vẫn chưa có thì dùng ưu tiên
             querydinhdanh = @" SELECT *  FROM tbDinhdanhtaikhoan where KeyValue like '%Ưu tiên%'"; // Sử dụng ? thay cho @mst trong OleDb

            result = ExecuteQuery(querydinhdanh, new OleDbParameter("?", ""));
            foreach (var item in people)
            {
               
                if (result.Rows.Count > 0)
                {
                    foreach (DataRow row in result.Rows)
                    {
                        if (string.IsNullOrEmpty(item.TKNo) || item.TKNo == "0")
                            item.TKNo = row["TKNo"].ToString();
                        if (string.IsNullOrEmpty(item.TKCo) || item.TKCo == "0")
                            item.TKCo = row["TKCo"].ToString();
                        if (item.TkThue == 0)
                            item.TkThue = int.Parse(row["TkThue"].ToString());
                        //Cho truong hop 331 711
                        if (item.TKNo == "3311")
                        {
                            item.TKNo = "711";
                            item.TKCo = "3311";
                        }
                    }
                }

            }
            //Điền lại diễn giải
            foreach (var item in people)
            {
                if (item.fileImportDetails.Count > 0)
                {
                    if (string.IsNullOrEmpty(item.Noidung))
                        item.Noidung = Helpers.ConvertVniToUnicode(item.fileImportDetails.FirstOrDefault().Ten);
                }
            }
            progressBarControl1.EditValue = 100;
        }
        private void btnChonthang_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(savedPath))
            {
                XtraMessageBox.Show("Vui lòng thiết lập đường dẫn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            LoadXmlFiles(savedPath);
            LoadExcel(savedPath);
            LoadDataGridview();
        }
        public void LoadExcel(string filePath)
        {
            //       var excelFiles = Directory.EnumerateFiles(filePath, "*.xlsx", SearchOption.AllDirectories)
            //.Where(file => !file.Contains("HDChonLoc")).ToList();  // Loại trừ đường dẫn chứa "HDChonLoc"
            filePath = filePath + "\\HDDauVao";
            int fromMonth = int.Parse(comboBoxEdit1.SelectedItem.ToString()); // Thay đổi theo tháng bắt đầu (ví dụ: 3 cho tháng 3)
            int toMonth = int.Parse(comboBoxEdit2.SelectedItem.ToString());   // Thay đổi theo tháng kết thúc (ví dụ: 7 cho tháng 7)

            // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
            var excelFiles = Directory.EnumerateFiles(filePath, "*.xlsx", SearchOption.AllDirectories)
                .Where(file => IsFileInMonthRange(file, filePath, fromMonth, toMonth)).ToList(); // Kiểm tra xem file có nằm trong khoảng tháng


            if (excelFiles.Count() == 0)
                return;
            // Kiểm tra xem file có tồn tại không
            for (int j = 0; j < excelFiles.Count; j++)
            {
                using (var workbook = new XLWorkbook(excelFiles[0]))
                {
                    // Lấy worksheet đầu tiên
                    var worksheet = workbook.Worksheet(1); // Hoặc bạn có thể dùng tên worksheet như worksheet = workbook.Worksheet("Sheet1");
                                                           // Lấy giá trị của ô A6
                    int rowCount = 0;
                    var currentCell = worksheet.Cell("A6"); // Bắt đầu từ ô A6

                    // Kiểm tra các ô bắt đầu từ A6 cho đến khi gặp ô trống
                    while (!currentCell.IsEmpty())
                    {
                        rowCount++; // Tăng số dòng
                        currentCell = currentCell.Worksheet.Row(currentCell.Address.RowNumber + 1).Cell("A"); // Chuyển xuống ô bên dưới

                    }
                    string SHDon = "";
                    string KHHDon = "";
                    string ten = "";
                    string mst = "";
                    string ten2 = "";
                    string mst2 = "";
                    string SoHD = "";
                    string diachi = "";
                    int TkCo = 0;
                    int TkNo = 0;
                    int TkThue = 0;
                    int Vat = 0;
                    string Mst = "";
                    double Thanhtien = 0;
                    DateTime NLap = new DateTime();
                    string diengiai = "";
                    int total = rowCount + 7;
                    for (int i = 7; i < (total - 1); i++)
                    {
                        progressPercentage = (filesLoaded * 100) / totalCount;
                        filesLoaded += 1;
                        progressBarControl1.EditValue = progressPercentage;
                        Application.DoEvents();

                        diengiai = "";
                        mst = worksheet.Cell(i, 6).Value.ToString().Trim();
                        NLap = DateTime.Parse(worksheet.Cell(i, 5).Value.ToString().Trim());
                        ten = worksheet.Cell(i, 7).Value.ToString();
                        SHDon = worksheet.Cell(i, 4).Value.ToString().Trim();
                        KHHDon = worksheet.Cell(i, 3).Value.ToString();
                        string query = "SELECT * FROM HoaDon WHERE KyHieu = ? AND SoHD LIKE ?";
                        string trangthaihoadon = worksheet.Cell(i, 16).Value.ToString();
                        if (trangthaihoadon.Contains("điều chỉnh"))
                        {
                            diengiai = "(*) Hóa đơn điều chỉnh";
                        }
                        // Tạo mảng tham số với giá trị cho câu lệnh SQL
                        OleDbParameter[] parameters = new OleDbParameter[]
                        {
            new OleDbParameter("KyHieu", KHHDon),          // Sử dụng chỉ số mà không cần tên
            new OleDbParameter("SoHD", "%" + SHDon + "%")  // Thêm ký tự % cho LIKE
                        };
                        var kq = ExecuteQuery(query, parameters);
                        if (kq.Rows.Count > 0)
                        {
                            continue;
                        }
                        //Kiem tra trong tbimport
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

                        double TienSauVAT = 0;
                        //Kiểm tra xem có phải trường hợp ko có thuế
                        //Lấy mst
                        if (worksheet.Cell(i, 6).Value.ToString() != "")
                        {
                            mst = worksheet.Cell(i, 6).Value.ToString();
                        }
                            if (worksheet.Cell(i, 9).Value.ToString() != "")
                        {
                            Thanhtien = double.Parse(worksheet.Cell(i, 9).Value.ToString().Replace(",", ""));
                            if (Thanhtien < 0)
                            {
                                diengiai = "(*) Hóa đơn điều chỉnh âm";
                            }
                            TienSauVAT = double.Parse(worksheet.Cell(i, 10).Value.ToString().Replace(",", ""));
                            if (TienSauVAT > 0)
                                Vat = int.Parse(Math.Round((TienSauVAT / Thanhtien * 100)).ToString());
                            else
                                Vat = 0;
                        }
                        else
                        {
                            Thanhtien = double.Parse(worksheet.Cell(i, 13).Value.ToString().Replace(",", ""));

                            Vat = 0;
                        }


                        //Kiểm tra thêm mới khách hàng
                        query = @" SELECT TOP 1 *  FROM KhachHang As kh
WHERE kh.MST = ?"; // Sử dụng ? thay cho @mst trong OleDb
                        DataTable result = ExecuteQuery(query, new OleDbParameter("?", mst));
                        if (result.Rows.Count == 0)
                        {
                            diachi = worksheet.Cell(i, 8).Value.ToString();
                            var Sohieu = GetLastFourDigits(mst);
                            ten = Helpers.ConvertUnicodeToVni(ten);
                            diachi = Helpers.ConvertUnicodeToVni(diachi);
                            //Kiểm tra sohieu có trùng nữa ko
                            query = @" SELECT TOP 1 *  FROM KhachHang As kh
WHERE kh.SoHieu = ?";
                            DataTable result2 = ExecuteQuery(query, new OleDbParameter("?", Sohieu));
                            if (result2.Rows.Count > 0)
                                Sohieu = "0" + Sohieu;
                            if (string.IsNullOrEmpty(diachi))
                            {
                                diachi = Helpers.ConvertUnicodeToVni("Bô sung địa chỉ");
                            }
                            InitCustomer(2, Sohieu, ten, diachi, mst);
                        }

                        //Kiểm tra đã có hóa đơn trước đó chưa 
                        TkNo = 6422;
                        TkCo = 1111;
                        TkThue = 1331;

                        if (!people.Any(m => m.SHDon.Contains(SHDon) && m.KHHDon == KHHDon))
                        {
                            people.Add(new FileImport(excelFiles[j], SHDon, KHHDon, NLap, ten, diengiai, TkNo.ToString(), TkCo.ToString(), TkThue, mst, Thanhtien, Vat, 2, ""));
                        }
                        //Load nội dung theo định danh
                        //Kiểm tra lại lại mã với Định danh
                        string querydinhdanh = @" SELECT *  FROM tbDinhdanhtaikhoan where KeyValue like '%MST%'"; // Sử dụng ? thay cho @mst trong OleDb

                        result = ExecuteQuery(querydinhdanh, new OleDbParameter("?", ""));
                        foreach (var item in people)
                        {
                            //Lấy danh sách định danh
                          
                            foreach (DataRow row in result.Rows)
                            {
                                string[] conditions = row["KeyValue"].ToString().Split('&');
                                string name = Helpers.ConvertUnicodeToVni((string)row["KeyValue"]);
                                int hasdata = 0;
                                foreach (string condition in conditions)
                                {
                                    string[] parts = Regex.Split(condition, @"([><=%]+)"); // Vẫn giữ % để linh hoạt nếu cần
                                    if (parts.Length == 3)
                                    {
                                        string key = parts[0];
                                        string operatorStr = parts[1];
                                        string valueStr = parts[2];
                                       
                                        if (key == "MST")
                                        {
                                            if (item.Mst == valueStr)
                                                hasdata += 1;
                                        }
                                    }

                                }
                                if (hasdata == conditions.Count())
                                {

                                    if (string.IsNullOrEmpty(item.Noidung))
                                        item.Noidung = row["Type"].ToString();
                                }
                            }
                        }
                    }

                }
            }

        }
        private void btnDuongdanthumuc_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Chọn thư mục bạn muốn lưu.";
                // folderBrowserDialog.rootFolder = Environment.SpecialFolder.MyComputer; // Thay đổi thư mục gốc nếu cần

                DialogResult result = folderBrowserDialog.ShowDialog();

                if (result == DialogResult.OK)
                {
                    string selectedPath = folderBrowserDialog.SelectedPath;

                    // Lưu đường dẫn thư mục vào App.config
                    Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    config.AppSettings.Settings["LastFilePath"].Value = selectedPath;
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");

                    savedPath = selectedPath;
                    txtPath.Text = savedPath;
                }
                else if (result == DialogResult.Cancel)
                {
                    // MessageBox.Show("Không có thư mục nào được chọn.");
                }
            }
        }

        private void btnSetupdbpath_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Cài đặt thuộc tính cho hộp thoại
            openFileDialog.Filter = "Access Database Files (*.mdb)|*.mdb|All Files (*.*)|*.*";
            openFileDialog.Title = "Chọn tệp MDB";

            // Hiển thị hộp thoại và kiểm tra nếu người dùng chọn tệp
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Lưu đường dẫn tệp vào TextBox
                txtdbPath.Text = openFileDialog.FileName;
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings["dbpath"].Value = txtdbPath.Text;
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
        }
        public static ChromeDriver Driver { get; private set; }
        #endregion

        #region Xu ly co quan thue
        int DoTask = 0;
        int Endtask = 0;
        private void btnTaicoquanthue_Click(object sender, EventArgs e)
        {
            int type = 0;
            if(chkDauvao.Checked && !chkDaura.Checked)
            {
                type =1;
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
        private void Taihoadon(int type)
        {
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
                    downloadPath = savedPath + "\\HDDauVao";
                if (type == 2)
                    downloadPath = savedPath + "\\HDDauRa";
                options.AddUserProfilePreference("download.default_directory", downloadPath);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("disable-popup-blocking", "true");
                options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
                options.AddUserProfilePreference("safebrowsing.enabled", false); // Tắt Safe Browsing hoàn toàn
                var driverPath = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                Driver = new ChromeDriver(driverPath, options);

                //
                try
                {
                    Driver.Navigate().GoToUrl("https://hoadondientu.gdt.gov.vn");
                    IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
                    js.ExecuteScript("window.scrollTo(0, 0);");
                    Thread.Sleep(1000);
                    var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
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
                    string username = ConfigurationManager.AppSettings["username"];
                    string password = ConfigurationManager.AppSettings["password"];
                    usernameField.SendKeys(username); // Thay your_username bằng tên đăng nhập thực tế
                    passwordField.SendKeys(password);
                    new Actions(Driver)
    .KeyDown(Keys.Tab).KeyUp(Keys.Tab)  // Tab lần 1
    .Pause(TimeSpan.FromMilliseconds(100))  // Đợi ngắn
    .KeyDown(Keys.Tab).KeyUp(Keys.Tab)  // Tab lần 2
    .Perform();
                    wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(200));
                    //chờ khi nao dang nhap xong
                    //                var button = wait.Until(d =>
                    //d.FindElement(By.CssSelector("button.ant-btn-icon-only i[aria-label='icon: user']"))
                    // .FindElement(By.XPath("./parent::button")));
                    wait.Until(d =>
                    d.FindElements(By.XPath("//div[contains(@class,'home-header-menu-item')]//span[text()='Đăng nhập']")).Count == 0);
                    DoTask = int.Parse(comboBoxEdit1.SelectedItem.ToString());
                    Endtask = int.Parse(comboBoxEdit2.SelectedItem.ToString());

                    if (type == 1)
                        Xulysaudangnhap();
                    if (type == 2)
                        Xulysaudangnhap2();

                }
                catch (Exception ex)
                {
                    Driver.Close();
                    // MessageBox.Show($"Lỗi: {ex.Message}");
                }
            }
        }
        int globaltype = 0;
        private void Xulysaudangnhap2()
        {
            if (DoTask > Endtask)
            {
                Driver.Quit(); // Đóng WebDriver
                return;
            }
            Thread.Sleep(1000);
            if (Driver == null)
            {
                MessageBox.Show("Vui lòng mở trình duyệt trước!");
                return;
            }
            Thread.Sleep(1000);
            var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
            string targetUrl = "https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don";
            Driver.Navigate().GoToUrl(targetUrl);
            Thread.Sleep(1000);
            // Tìm input với class 'ant-calendar-input' và placeholder 'Chọn thời điểm'
            var allInputs = Driver.FindElements(By.CssSelector("input.ant-calendar-picker-input"));
            Thread.Sleep(100);
            allInputs[0].Click();
            IWebElement monthSelect = Driver.FindElement(
By.CssSelector("a.ant-calendar-month-select[title='Chọn tháng']"));
            monthSelect.Click();
            IWebElement monthItem = Driver.FindElement(
By.XPath("//a[contains(@class,'ant-calendar-month-panel-month') and text()='Thg 0" + DoTask.ToString() + "']"));
            monthItem.Click();

            //
            var elements = Driver.FindElements(By.CssSelector("div.ant-calendar-date"));

            // Lọc các phần tử có text là "1"
            var targetElement = elements.FirstOrDefault(div => div.Text.Trim() == "1");
            targetElement.Click();
            new Actions(Driver)
.SendKeys(Keys.Enter) // Tab lần 2
.Perform();
            var button = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn') and .//span[text()='Tìm kiếm']])[1]")));
            button.Click();

            //Chờ table xuất hiện
            wait.Until(d => d.FindElements(By.CssSelector("tr.ant-table-row")).Count > 0);
            IReadOnlyCollection<IWebElement> rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
            int rowCount = rows.Count;
            //Chọn 50 rows
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));

            //Chọn 50 rows
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
            var divElement = wait.Until(d => d.FindElements(By.XPath("//div[@class='ant-select-selection-selected-value' and @title='15']")));

            // Kiểm tra nếu phần tử được tìm thấy và nhấp vào nó
            if (divElement != null && divElement[0].Displayed)
            {
                divElement[0].Click();
                Console.WriteLine("Đã nhấp vào phần tử.");
            }
            var dropdownMenu = wait.Until(d => d.FindElement(By.ClassName("ant-select-dropdown-menu")));

            // Tìm phần tử <li> có nội dung là "50" và nhấp vào nó
            var option50 = wait.Until(d => dropdownMenu.FindElements(By.XPath(".//li[text()='50']")));

            // Nhấp vào phần tử "50"
            if (option50 != null)
            {
                option50[0].Click();
            }
            //Click download XML   
            Thread.Sleep(1000);
            //
            bool isPhantrang = false;
            while (isPhantrang == false)
            {
                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(5));
                bool hasrow = false;
                // Đợi cho đến khi có ít nhất 1 dòng xuất hiện
                try
                {
                    wait.Until(d => d.FindElements(By.CssSelector("tr.ant-table-row")).Count > 0);
                    hasrow = true;
                }
                catch(Exception ex)
                {
                    hasrow = false;
                }
                if (hasrow)
                {
                     rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
                     rowCount = rows.Count;

                    Console.WriteLine($"Số dòng trong bảng: {rowCount}");


                    int currentRow = 1;
                    bool hasMoreRows = true;
                    List<string> lstHas = new List<string>();
                    int hasdata = 0;
                    while ((currentRow) <= rowCount)
                    {
                        try
                        {
                            // Tìm dòng hiện tại
                            var row = wait.Until(d =>
                                d.FindElement(By.XPath($"(//tbody[@class='ant-table-tbody']/tr[contains(@class,'ant-table-row')])[{currentRow}]")));
                            var cellC25TYY = row.FindElement(By.XPath("./td[3]/span")).Text; // C25TYY
                            var cell22252 = row.FindElement(By.XPath("./td[4]")).Text; // 22252

                            string query = "SELECT * FROM HoaDon WHERE KyHieu = ? AND SoHD LIKE ?";


                            // Tạo mảng tham số với giá trị cho câu lệnh SQL
                            OleDbParameter[] parameters = new OleDbParameter[]
                            {
            new OleDbParameter("KyHieu", cellC25TYY),          // Sử dụng chỉ số mà không cần tên
            new OleDbParameter("SoHD", "%" + cell22252 + "%")  // Thêm ký tự % cho LIKE
                            };


                            // Click vào dòng
                            row.Click();
                            button = wait.Until(d =>
                             d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[13]")));
                            button.Click();
                            // Xử lý sau khi click (đợi tải, đóng popup,...)
                            Thread.Sleep(50); // Đợi 1 giây giữa các lần click
                            string fp = "";
                            if (currentRow == 15)
                            {
                                int aas = 10;
                            }
                            if (currentRow == 1)
                                fp = savedPath + "\\HDDauRa\\" + "invoice.zip";
                            else
                                fp = savedPath + "\\HDDauRa\\" + "invoice (" + (currentRow - 1 - hasdata) + ").zip";

                            try
                            {
                                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(5));
                                wait.Until(d => File.Exists(fp));
                                lstHas.Add(fp);
                            }
                            catch(Exception ex)
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
                            currentRow++; // Vẫn tiếp tục với dòng tiếp theo
                        }
                    }
                    if (lstHas.Count > 0)
                    {
                        //var getlastlist = lstHas.LastOrDefault();
                        //wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
                        //wait.Until(d => File.Exists(getlastlist));
                        GiaiNenhoadon(2);
                    }
                    //Xử lý phần trang 
                    var buttonElement = Driver.FindElements(By.ClassName("ant-btn-primary"));

                    // Kiểm tra xem button có bị vô hiệu hóa không
                    bool isDisabled = !buttonElement[3].Enabled;
                    if (isDisabled == false)
                    {
                        buttonElement[3].Click();
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        isPhantrang = true;
                    }
                }

                Xulymaytinhtien2(wait);
                DoTask += 1;

                Xulysaudangnhap2();
            }
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
        private void Xulysaudangnhap()
        {
            if (DoTask > Endtask)
            {
                Driver.Quit(); // Đóng WebDriver
                return;
            }
            Thread.Sleep(1000);
            if (Driver == null)
            {
                MessageBox.Show("Vui lòng mở trình duyệt trước!");
                return;
            }

            try
            {
                var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
                var notificationButton = wait.Until(d => d.FindElement(By.XPath("//i[@aria-label='icon: bell']/parent::button")));

                string targetUrl = "https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don";
                Thread.Sleep(5000);
                // Cách 1: Chuyển trang đơn giản
                Driver.Navigate().GoToUrl(targetUrl);
                Thread.Sleep(5000);
                var tab = wait.Until(d => d.FindElement(
                    By.XPath("//div[@role='tab' and .//span[contains(text(),'Tra cứu hóa đơn điện tử mua vào')]]")
                ));
                tab.Click();


                Thread.Sleep(1000);
                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));

                Thread.Sleep(1000);

                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
                // Tìm input với class 'ant-calendar-input' và placeholder 'Chọn thời điểm'
                var allInputs = Driver.FindElements(By.CssSelector("input.ant-calendar-picker-input"));
                Thread.Sleep(1000);
                allInputs[2].Click();
                IWebElement monthSelect = Driver.FindElement(
By.CssSelector("a.ant-calendar-month-select[title='Chọn tháng']"));
                monthSelect.Click();
                IWebElement monthItem = Driver.FindElement(
By.XPath("//a[contains(@class,'ant-calendar-month-panel-month') and text()='Thg 0" + DoTask.ToString() + "']"));
                monthItem.Click();

                //
                var elements = Driver.FindElements(By.CssSelector("div.ant-calendar-date"));

                // Lọc các phần tử có text là "1"
                var targetElement = elements.FirstOrDefault(div => div.Text.Trim() == "1");
                targetElement.Click();
                new Actions(Driver)
.SendKeys(Keys.Enter) // Tab lần 2
.Perform();
                var button = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn') and .//span[text()='Tìm kiếm']])[2]")));
                button.Click();

                //Chờ loading ẩn
                waitLoading(wait);

                //wait.Until(d => d.FindElements(By.CssSelector("tr.ant-table-row")).Count > 0);
                //IReadOnlyCollection<IWebElement> rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
                //int rowCount = rows.Count;
                //Chọn 50 rows
                wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
                

                var divElement = wait.Until(d => d.FindElements(By.XPath("//div[@class='ant-select-selection-selected-value' and @title='15']")));
                // Nếu tìm thấy ít nhất một phần tử
                if (divElement[1] != null)
                {
                    // Thực hiện cuộn đến phần tử với thời gian
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
                var dropdownMenu = wait.Until(d => d.FindElement(By.ClassName("ant-select-dropdown-menu")));

                // Tìm phần tử <li> có nội dung là "50" và nhấp vào nó
                var option50 = wait.Until(d => dropdownMenu.FindElements(By.XPath(".//li[text()='50']")));

                // Nhấp vào phần tử "50"
                if (option50 != null)
                {
                    option50[0].Click();
                }
                //Click download XML
                //chờ loading tiếp
                waitLoading(wait);

                // Cách 1: Target vào thẻ <i> có aria-label
                //   d.FindElement(By.CssSelector("button.ant-btn-icon-only i[aria-label='icon: user']")
                Thread.Sleep(1000);
                bool isPhantrang = false;
                try
                {
                    while (isPhantrang == false)
                    {
                        wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));

                        // Đợi cho đến khi có ít nhất 1 dòng xuất hiện
                        wait.Until(d => d.FindElements(By.CssSelector("tr.ant-table-row")).Count > 0);
                        IReadOnlyCollection<IWebElement> rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
                        var rowCount = rows.Count;

                        Console.WriteLine($"Số dòng trong bảng: {rowCount}");


                        int currentRow = 1;
                        bool hasMoreRows = true;
                        List<string> lstHas = new List<string>();
                        int hasdata = 0;
                        while ((currentRow) <= rowCount)
                        {
                            try
                            {
                                // Tìm dòng hiện tại
                                var row = wait.Until(d =>
                                    d.FindElement(By.XPath($"(//tbody[@class='ant-table-tbody']/tr[contains(@class,'ant-table-row')])[{currentRow}]")));
                                var cellC25TYY = row.FindElement(By.XPath("./td[3]/span")).Text; // C25TYY
                                var cell22252 = row.FindElement(By.XPath("./td[4]")).Text; // 22252

                                string query = "SELECT * FROM HoaDon WHERE KyHieu = ? AND SoHD LIKE ?";


                                // Tạo mảng tham số với giá trị cho câu lệnh SQL
                                OleDbParameter[] parameters = new OleDbParameter[]
                                {
            new OleDbParameter("KyHieu", cellC25TYY),          // Sử dụng chỉ số mà không cần tên
            new OleDbParameter("SoHD", "%" + cell22252 + "%")  // Thêm ký tự % cho LIKE
                                };
                                //var kq = ExecuteQuery(query, parameters);
                                //var a = people;
                                //var check = a.Any(m => m.KHHDon == cellC25TYY && m.SHDon.Contains(cell22252));
                                //if (check || kq.Rows.Count != 0)
                                //{
                                //    currentRow++;
                                //    hasdata++;
                                //    continue;
                                //}

                                // Click vào dòng
                                row.Click();
                                button = wait.Until(d =>
                                 d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[19]")));
                                button.Click();
                                // Xử lý sau khi click (đợi tải, đóng popup,...)
                                waitLoading(wait);
                                string fp = "";
                                if (currentRow == 1)
                                    fp = savedPath + "\\HDDauVao\\" + "invoice.zip";
                                else
                                    fp = savedPath + "\\HDDauVao\\" + "invoice (" + (currentRow - 1 - hasdata) + ").zip";

                                try
                                {
                                    wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                                    wait.Until(d => File.Exists(fp));
                                    lstHas.Add(fp);
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
                                currentRow++; // Vẫn tiếp tục với dòng tiếp theo
                            }
                        }
                        if (lstHas.Count > 0)
                        {
                            //var getlastlist = lstHas.LastOrDefault();
                            //wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
                            //wait.Until(d => File.Exists(getlastlist));
                            GiaiNenhoadon(1);
                        }
                        //Xử lý phần trang 
                        var buttonElement = Driver.FindElements(By.ClassName("ant-btn-primary"));

                        // Kiểm tra xem button có bị vô hiệu hóa không
                        bool isDisabled = !buttonElement[7].Enabled;
                        if (isDisabled == false)
                        {
                            buttonElement[7].Click();
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
                    throw ex;
                }
                Thread.Sleep(1000);
                Cucthuekhngnhanma(wait); 
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}");
            }
        }
        private void GiaiNenhoadon(int type)
        {
            string typepath = "";
            if (type == 1)
                typepath = "\\HDDauVao";
            if (type == 2)
                typepath = "\\HDDauRa";
            string pah = savedPath + typepath;

            string[] zipFiles = Directory.GetFiles(pah, "*.zip");
            var i = 0;
            foreach (var zipFile in zipFiles)
            {
                string filename = "";
                if (i == 0)
                {
                    filename = "invoice.zip";
                    ExtractZip(pah, filename);
                }
                else
                {
                    filename = "invoice (" + i + ").zip";
                    ExtractZip(pah, filename);
                }
                i++;
            }

        }
        private void ExtractZip(string path, string filename)
        {
            string downloadPath = path;
            string zipFileName = filename; // Thay bằng tên file thực tế


            string zipFilePath = System.IO.Path.Combine(downloadPath, zipFileName);
            string extractPath = System.IO.Path.Combine(downloadPath, System.IO.Path.GetFileNameWithoutExtension(zipFileName));
            // Tạo thư mục giải nén nếu chưa tồn tại
            Directory.CreateDirectory(extractPath);

            // Giải nén file
            try
            {
                ZipFile.ExtractToDirectory(zipFilePath, extractPath);
                Console.WriteLine($"Đã giải nén thành công vào: {extractPath}");
                File.Delete(zipFilePath);
                //Vào thư mục mới tạo
                string invoiceFilePath = System.IO.Path.Combine(extractPath, "invoice.xml");
                string htmlFilepath = System.IO.Path.Combine(extractPath, "invoice.html");
                var newname = Getnewname(invoiceFilePath, 1);
                var newnamehtml = Getnewname(invoiceFilePath, 2);
                var getsplit = (newname.Split('_'))[2];
                var getsplithtml = (newnamehtml.Split('_'))[2];
                newname = getsplit + "\\" + newname;
                newnamehtml = getsplithtml + "\\" + newnamehtml;
                string newFilePath = System.IO.Path.Combine(path, newname);
                string newFilePathhtml = System.IO.Path.Combine(path, newnamehtml);
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
                        if (!File.Exists(newFilePathhtml))
                        {
                            File.Move(htmlFilepath, newFilePathhtml);
                        }
                       
                        Directory.Delete(extractPath, true); // true để xóa cả nội dung bên trong
                    }
                    catch (Exception ex)
                    {
                        var ms = ex.Message;
                    }
                }
                else
                {
                    Console.WriteLine("Tệp invoice.xml không tồn tại trong thư mục giải nén.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi giải nén: {ex.Message}");
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
            wait.Until(d => d.FindElements(By.CssSelector("tr.ant-table-row")).Count > 0);
            Thread.Sleep(1000);
            IReadOnlyCollection<IWebElement> rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
            int rowCount = rows.Count;

            Console.WriteLine($"Số dòng trong bảng: {rowCount}");



            int currentRow = 1;
            bool hasMoreRows = true;
            List<string> lstHas = new List<string>();
            int hasdata = 0;
            while ((currentRow) <= rowCount)
            {
                try
                {
                    // Tìm dòng hiện tại
                    var row = wait.Until(d =>
                        d.FindElement(By.XPath($"(//tbody[@class='ant-table-tbody']/tr[contains(@class,'ant-table-row')])[{currentRow}]")));
                    var cellC25TYY = row.FindElement(By.XPath("./td[3]/span")).Text; // C25TYY
                    var cell22252 = row.FindElement(By.XPath("./td[4]")).Text; // 22252

                    string query = "SELECT * FROM HoaDon WHERE KyHieu = ? AND SoHD LIKE ?";


                    // Tạo mảng tham số với giá trị cho câu lệnh SQL
                    OleDbParameter[] parameters = new OleDbParameter[]
                    {
            new OleDbParameter("KyHieu", cellC25TYY),          // Sử dụng chỉ số mà không cần tên
            new OleDbParameter("SoHD", "%" + cell22252 + "%")  // Thêm ký tự % cho LIKE
                    };
                  
                    // Click vào dòng
                    row.Click();
                   var button = wait.Until(d =>
                     d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[19]")));
                    button.Click();
                    // Xử lý sau khi click (đợi tải, đóng popup,...)
                    Thread.Sleep(50); // Đợi 1 giây giữa các lần click
                    string fp = "";
                    if (currentRow == 15)
                    {
                        int aas = 10;
                    }
                    if (currentRow == 1)
                        fp = savedPath + "\\HDDauVao\\" + "invoice.zip";
                    else
                        fp = savedPath + "\\HDDauVao\\" + "invoice (" + (currentRow - 1 - hasdata) + ").zip";

                    wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
                    wait.Until(d => File.Exists(fp));
                    lstHas.Add(fp);
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
                    currentRow++; // Vẫn tiếp tục với dòng tiếp theo
                }
            }
            if (lstHas.Count == 0)
                return;
            var getlastlist = lstHas.LastOrDefault();
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
            wait.Until(d => File.Exists(getlastlist));

            var pp = savedPath + "\\HDDauVao";
            var pp2 = savedPath + "\\HDDauVao\\" + DoTask;
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


            GiaiNenhoadon(2);
            //  LoadXmlFiles(savedPath);

            //End Xử lý hóa đơn từ máy tính tiền
        }
        private void Xulymaytinhtien(WebDriverWait wait)
        {
            //Xử lý hóa đơn từ máy tính tiền
            //By.Id("ttxly")
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
            // Tìm input với class 'ant-calendar-input' và placeholder 'Chọn thời điểm'
            //            var allInputs = Driver.FindElements(By.CssSelector("input.ant-calendar-picker-input"));
            //            Thread.Sleep(1000);
            //            allInputs[2].Click();
            //            IWebElement monthSelect = Driver.FindElement(
            //By.CssSelector("a.ant-calendar-month-select[title='Chọn tháng']"));
            //            monthSelect.Click();
            //            IWebElement monthItem = Driver.FindElement(
            //By.XPath("//a[contains(@class,'ant-calendar-month-panel-month') and text()='Thg 0" + DoTask.ToString() + "']"));
            //            monthItem.Click();

            //            //
            //            var elements = Driver.FindElements(By.CssSelector("div.ant-calendar-date"));

            //            // Lọc các phần tử có text là "1"
            //            var targetElement = elements.FirstOrDefault(div => div.Text.Trim() == "1");
            //            targetElement.Click();
            //            new Actions(Driver)
            //.KeyDown(Keys.Enter) // Tab lần 2
            //.Perform();

            //Tìm tab tính tiền
            var tabElement = wait.Until(d => d.FindElements(By.XPath("//div[@role='tab']"))
               .FirstOrDefault(e => e.Text.Trim() == "Hóa đơn có mã khởi tạo từ máy tính tiền"));
            if (tabElement != null)
            {
                tabElement.Click();
                Console.WriteLine("Đã nhấp vào tab.");
            }
            var button = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn') and .//span[text()='Tìm kiếm']])[2]")));
            button.Click();
            //chờ loading

            waitLoading(wait);
            wait.Until(d => d.FindElements(By.CssSelector("tr.ant-table-row")).Count > 0);
            Thread.Sleep(1000);
            IReadOnlyCollection<IWebElement> rows = Driver.FindElements(By.CssSelector("tr.ant-table-row"));
            int rowCount = rows.Count;

            Console.WriteLine($"Số dòng trong bảng: {rowCount}");

           

            int currentRow = 1;
            bool hasMoreRows = true;
            List<string> lstHas = new List<string>();
            int hasdata = 0;
            while ((currentRow) <= rowCount)
            {
                try
                {
                    // Tìm dòng hiện tại
                    var row = wait.Until(d =>
                        d.FindElement(By.XPath($"(//tbody[@class='ant-table-tbody']/tr[contains(@class,'ant-table-row')])[{currentRow}]")));
                    var cellC25TYY = row.FindElement(By.XPath("./td[3]/span")).Text; // C25TYY
                    var cell22252 = row.FindElement(By.XPath("./td[4]")).Text; // 22252

                    string query = "SELECT * FROM HoaDon WHERE KyHieu = ? AND SoHD LIKE ?";


                    // Tạo mảng tham số với giá trị cho câu lệnh SQL
                    OleDbParameter[] parameters = new OleDbParameter[]
                    {
            new OleDbParameter("KyHieu", cellC25TYY),          // Sử dụng chỉ số mà không cần tên
            new OleDbParameter("SoHD", "%" + cell22252 + "%")  // Thêm ký tự % cho LIKE
                    };
                    //var kq = ExecuteQuery(query, parameters);
                    //var a = people;
                    //var check = a.Any(m => m.KHHDon == cellC25TYY && m.SHDon.Contains(cell22252));
                    //if (check || kq.Rows.Count != 0)
                    //{
                    //    //Xóa Excel
                    //    var pp1 = "C:\\S.T.E 25\\S.T.E 25\\Hoadon\\HDDauVao";
                    //    var pp21 = "C:\\S.T.E 25\\S.T.E 25\\Hoadon\\HDDauVao\\" + DoTask;
                    //    // Lấy tất cả các file XML từ các thư mục tháng từ fromMonth đến toMonth
                    //    string[] excelFiles1 = Directory.GetFiles(pp1, "*.xlsx");
                    //    string fileName1 = Path.GetFileName(excelFiles1[0]); // Lấy tên file
                    //    string destFilePath1 = Path.Combine(pp21, fileName1); // Tạo đường dẫn đích

                    //    // Di chuyển file
                    //    try
                    //    {
                    //        File.Move(excelFiles1[0], destFilePath1);
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        File.Delete(excelFiles1[0]);
                    //    }
                    //    currentRow++;
                    //    hasdata++;
                    //    continue;
                    //}

                    // Click vào dòng
                    row.Click();
                    button = wait.Until(d =>
                     d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[19]")));
                    button.Click();
                    // Xử lý sau khi click (đợi tải, đóng popup,...)
                    waitLoading(wait);
                    string fp = "";
                    if (currentRow == 15)
                    {
                        int aas = 10;
                    }
                    if (currentRow == 1)
                        fp = savedPath +"\\HDDauVao\\" + "invoice.zip";
                    else
                        fp = savedPath +"\\HDDauVao\\" + "invoice (" + (currentRow - 1 - hasdata) + ").zip";

                    wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
                    wait.Until(d => File.Exists(fp));
                    lstHas.Add(fp);
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
                    currentRow++; // Vẫn tiếp tục với dòng tiếp theo
                }
            }
            if (lstHas.Count == 0)
                return;
            var getlastlist = lstHas.LastOrDefault();
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(120));
            wait.Until(d => File.Exists(getlastlist));

            var pp = savedPath + "\\HDDauVao";
            var pp2 = savedPath + "\\HDDauVao\\" + DoTask;
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


            GiaiNenhoadon(1);
            //  LoadXmlFiles(savedPath);

            //End Xử lý hóa đơn từ máy tính tiền
            DoTask += 1;
            Xulysaudangnhap();
        }
        private void Cucthuekhngnhanma(WebDriverWait wait)
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
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
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
            button.Click();
            // 
            waitLoading(wait);
            button = wait.Until(d =>
                d.FindElement(By.XPath("(//button[contains(@class, 'ant-btn-icon-only')])[18]")));


            ((IJavaScriptExecutor)Driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth'});", button);

            // Hover rồi mới click
            new Actions(Driver)
                .MoveToElement(button)
                .Pause(TimeSpan.FromSeconds(1))
                .Click()
                .Perform();

            //Tải file excel
            Xulymaytinhtien(wait);
        }
        private void simpleButton6_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtUsername.Text) || string.IsNullOrEmpty(txtPassword.Text))
            {
                XtraMessageBox.Show("Vui lòng nhập thông tin!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["username"].Value = txtuser.Text;
            config.AppSettings.Settings["password"].Value = txtpass.Text;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
            XtraMessageBox.Show("Cập nhật tài khoản thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
        #endregion
        #region Database Excute, query
        public void InitCustomer(int Maphanloai, string Sohieu, string Ten, string Diachi, string Mst)
        {
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
        private DataTable ExecuteQuery(string query, params OleDbParameter[] parameters)
        {
            DataTable dataTable = new DataTable();

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
            DataTable dataTable = new DataTable();

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

        }

        private void gridControl1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void gridControl1_Click(object sender, EventArgs e)
        {
            GridView gridView = gridControl1.MainView as GridView;
            var hitInfo = gridView.CalcHitInfo(gridView.GridControl.PointToClient(MousePosition));


            // Kiểm tra nếu nhấp vào một ô
            if (hitInfo.InRowCell)
            {
                int columnIndex = hitInfo.Column.VisibleIndex; // Chỉ số cột
                if (columnIndex != 0)
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

        public string hiddenValue { get; set; }
        private void gridControl1_KeyUp(object sender, KeyEventArgs e)
        {
            GridView gridView = gridControl1.MainView as GridView;
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
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
                if (currentValue.ToString().Contains("154"))
                {
                    frmCongtrinh frmCongtrinh = new frmCongtrinh();
                    frmCongtrinh.frmMain = this;
                    frmCongtrinh.ShowDialog();
                    if (currentValue.ToString().Contains("|"))
                        currentValue = currentValue.ToString().Split('|')[0];
                    gridView.SetRowCellValue(currentRowHandle, "TKNo", currentValue + "|" + hiddenValue);
                    return;
                }
                // Di chuyển xuống hàng
                int nextRowHandle = currentRowHandle + 1;

                // Kiểm tra xem hàng tiếp theo có tồn tại không
                if (nextRowHandle < gridView.DataRowCount)
                {
                    // Gán giá trị cho cột trong hàng tiếp theo
                    gridView.SetRowCellValue(nextRowHandle, gridView.FocusedColumn, currentValue);
                    if (gridView.FocusedColumn == gridView.Columns["TKNo"])
                    {
                        currentValue = gridView.GetRowCellValue(currentRowHandle, gridView.Columns["TKCo"]).ToString();
                        gridView.SetRowCellValue(nextRowHandle, gridView.Columns["TKCo"], currentValue);
                        currentValue = gridView.GetRowCellValue(currentRowHandle, gridView.Columns["Noidung"]).ToString();
                        gridView.SetRowCellValue(nextRowHandle, gridView.Columns["Noidung"], currentValue);
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
        private void btnimport_Click(object sender, EventArgs e)
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
                    if (item.TKNo == "6422" || item.TKNo == "6421")
                        item.TkThue = 1331;
                    if (item.TKNo == "152")
                        item.TkThue = 1331;
                    if (item.TKNo == "5111" || item.TKNo == "5112" || item.TKNo == "5113")
                        item.TkThue = 33311;
                }

                string query = @"
            INSERT INTO tbImport (SHDon, KHHDon, NLap, Ten, Noidung, TKCo, TKNo, TkThue, Mst, Status, Ngaytao, TongTien, Vat, SohieuTP)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

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
            new OleDbParameter("?", item.SoHieuTP.ToString())
                };

                int a = ExecuteQueryResult(query, parameters);

                if (a > 0)
                {
                    if (item.TKNo != "6422" && item.TKNo != "64221" && !item.TKNo.Contains("154|"))
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

                            query = @"
                        INSERT INTO tbimportdetail (ParentId, SoHieu, SoLuong, DonGia, DVT, Ten,MaCT,TKNo)
                        VALUES (?, ?, ?, ?, ?, ?,?,?)";

                            parameters = new OleDbParameter[]
                            {
                        new OleDbParameter("?", parentID),
                        new OleDbParameter("?", it.SoHieu),
                        new OleDbParameter("?", it.Soluong),
                        new OleDbParameter("?", it.Dongia),
                        new OleDbParameter("?", it.DVT),
                        new OleDbParameter("?", it.Ten),
                        new OleDbParameter("?", it.MaCT),
                        new OleDbParameter("?", it.TKNo)
                            };

                            int resl = ExecuteQueryResult(query, parameters);
                            InsertHangHoa(it.DVT, it.SoHieu, it.Ten);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Thêm dữ liệu thất bại.");
                }

                try
                {
                    var htmlPath = Path.Combine(savedPath, "HDDauVao");
                    var month = "\\" + item.NLap.Month;
                    htmlPath += month;

                    var htmlFiles = Directory.EnumerateFiles(htmlPath, "*.html", SearchOption.AllDirectories);
                    foreach (var it in htmlFiles)
                    {
                        File.Move(it, it.Replace("HDDauVao", "HDChonLoc"));
                    }

                    try
                    {
                        File.Move(item.Path, item.Path.Replace("HDDauVao", "HDChonLoc"));
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
        bool isClick = false;
        public void InsertHangHoa(string DVTinh, string sohieu, string newName)
        {
            sohieu = sohieu.ToUpper();
            if (DVTinh == "Exception")
            {
                return;
            }

            string query = "";
            if (DVTinh == "kWh")
            {
                return;
            }

            // Insert thêm vô database
            // Trước khi insert vật tư, kiểm tra phân loại vật tư trước
            string km = Helpers.ConvertUnicodeToVni("khuyến mãi").ToLower();
            string searchTerm = "%" + km + "%";

            string querydinhdanh = @"SELECT * FROM PhanLoaiVattu WHERE TenPhanLoai LIKE ?"; // Sử dụng ? thay cho @mst trong OleDb
            var resultkm = ExecuteQuery(querydinhdanh, new OleDbParameter("?", searchTerm));

            // Nếu chưa có thì thêm mới
            if (resultkm.Rows.Count == 0)
            {
                query = @"
            INSERT INTO PhanLoaiVattu (SoHieu, TenPhanLoai, Cap, MaTK)
            VALUES (?, ?, ?, ?)";

                var parameterss = new OleDbParameter[]
                {
            new OleDbParameter("?", "HKM"),
            new OleDbParameter("?", "Haøng khuyeán maõi"),
            new OleDbParameter("?", 1),
            new OleDbParameter("?", 39)
                };

                int rr = ExecuteQueryResult(query, parameterss);
            }

            query = @"SELECT * FROM Vattu 
              WHERE LCase(TenVattu) = LCase(?) AND LCase(DonVi) = LCase(?)";

            var getdata = ExecuteQuery(query,
                new OleDbParameter("?", newName.ToLower()),
                new OleDbParameter("?", DVTinh.ToLower()));

            if (getdata.Rows.Count == 0)
            {
                query = @"
            INSERT INTO Vattu (MaPhanLoai, SoHieu, TenVattu, DonVi)
            VALUES (?, ?, ?, ?)";

                int maphanloai = 0;
                if (newName.Contains(km))
                {
                    maphanloai = int.Parse(resultkm.Rows[0]["MaSo"].ToString());
                }
                else
                {
                    // Nếu chưa có nhóm tạm thì tạo
                    string nt = Helpers.ConvertUnicodeToVni("Nhóm tạm").ToLower();
                    string searchTemp = "%" + nt + "%";

                    string quryNhomtam = @"SELECT * FROM PhanLoaiVattu WHERE TenPhanLoai LIKE ?";
                    var resultnt = ExecuteQuery(querydinhdanh, new OleDbParameter("?", searchTemp));

                    if (resultnt.Rows.Count == 0)
                    {
                        // Tạo nhóm tạm
                        query = @"
                    INSERT INTO PhanLoaiVattu (SoHieu, TenPhanLoai, Cap, MaTK)
                    VALUES (?, ?, ?, ?)";

                        var parameterss = new OleDbParameter[]
                        {
                    new OleDbParameter("?", "NT"),
                    new OleDbParameter("?", nt),
                    new OleDbParameter("?", 1),
                    new OleDbParameter("?", 39)
                        };

                        int rr = ExecuteQueryResult(query, parameterss);
                        quryNhomtam = @"SELECT * FROM PhanLoaiVattu WHERE TenPhanLoai LIKE ?";
                        resultnt = ExecuteQuery(querydinhdanh, new OleDbParameter("?", searchTemp));

                        if (resultnt.Rows.Count > 0)
                        {
                            maphanloai = int.Parse(resultnt.Rows[0]["MaSo"].ToString());
                        }
                    }
                    else
                    {
                        maphanloai = int.Parse(resultnt.Rows[0]["MaSo"].ToString());
                    }
                }

                var parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", maphanloai),
            new OleDbParameter("?", sohieu.ToUpper()),
            new OleDbParameter("?", newName),
            new OleDbParameter("?", DVTinh)
                };

                // Thực thi truy vấn và lấy kết quả
                int a = ExecuteQueryResult(query, parameters);
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

        private void btnLuudinhdanh_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtTukhoa.Text) || string.IsNullOrEmpty(txtTKNo.Text) || string.IsNullOrEmpty(txtTKCo.Text) || string.IsNullOrEmpty(txtTKThue.Text))
            {
                XtraMessageBox.Show("Vui lòng nhập thông tin!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string query = @"
        INSERT INTO tbDinhdanhtaikhoan (KeyValue,TKNo,TKCo,TKThue,Type)
        VALUES (?,?,?,?,?)";
            OleDbParameter[] parameters = new OleDbParameter[]
{
        new OleDbParameter("?",txtTukhoa.Text),
           new OleDbParameter("?",txtTKNo.Text),
                 new OleDbParameter("?",txtTKCo.Text),
             new OleDbParameter("?",txtTKThue.Text),
              new OleDbParameter("?",txtDiengiai.Text)
};

            // Thực thi truy vấn và lấy kết quả
            int a = ExecuteQueryResult(query, parameters);
            LoadDataDinhDanh();
        }
        private void LoadDataDinhDanh()
        {
            if (string.IsNullOrEmpty(savedPath) || string.IsNullOrEmpty(dbPath))
                return;
            string querykh = @" SELECT *  FROM tbDinhdanhtaikhoan"; // Sử dụng ? thay cho @mst trong OleDb

            result = ExecuteQuery(querykh, new OleDbParameter("?", ""));
            gcDinhdanh.DataSource = result;
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
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            config.AppSettings.Settings["username"].Value = txtuser.Text;
            // config.AppSettings.Settings["password"].Value = txtpass.Text;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void txtpass_TextChanged_1(object sender, EventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            //config.AppSettings.Settings["username"].Value = txtuser.Text;
            config.AppSettings.Settings["password"].Value = txtpass.Text;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        public static string NormalizeVietnameseString(string input)
        {

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
