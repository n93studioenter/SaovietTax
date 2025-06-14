using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System.Reflection;
using System.Threading;
using Keys = OpenQA.Selenium.Keys;
using System.Diagnostics;
using System.IO;
using Windows.UI.ViewManagement;
using System.IO.Compression;
using System.Xml;
using DevExpress.CodeParser;
using XmlNode = System.Xml.XmlNode;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using System.Xml.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using DevExpress.Utils;
using SaovietTax.DTO;
using Newtonsoft.Json;
using System.Text.Json;
using Windows.Media.Protection.PlayReady;
using DevExpress.XtraSpreadsheet;
using DevExpress.Spreadsheet;
using System.Data.OleDb;
using DocumentFormat.OpenXml.VariantTypes;
namespace SaovietTax
{
    public partial class frmTaiCoQuanThue : DevExpress.XtraEditors.XtraForm
    {
        public frmTaiCoQuanThue()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.Manual; // Cài đặt vị trí thủ công

        }
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // Đặt vị trí ở góc trên bên trái
            this.Location = new Point(0, 0);

            // Hoặc để ở góc phải
            // this.Location = new Point(Screen.PrimaryScreen.WorkingArea.Width - this.Width, 0);
        }
        public frmMain frmMain;
        public static ChromeDriver Driver { get; private set; }
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
        private string Readcapcha()

        {
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "captcha.txt"; // Đảm bảo tệp ở cùng thư mục với chương trình

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
        public void Xulydaura1(string tokken)
        {
            using (var client = new HttpClient())
            {
                // Đặt URL
                string formattedDate1 = frmMain.dtFrom.ToString("dd/MM/yyyyTHH:mm:ss");
                string formattedDate2 = frmMain.dtTo.ToString("dd/MM/yyyyTHH:mm:ss");
                string url = $@"https://hoadondientu.gdt.gov.vn:30000/query/invoices/sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=50&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2}";

                // Thêm Bearer token vào Header
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                try
                {
                    // Gửi yêu cầu GET
                    HttpResponseMessage response = client.GetAsync(url).Result;

                    // Đảm bảo phản hồi thành công
                    response.EnsureSuccessStatusCode();

                    // Đọc nội dung phản hồi
                    string responseBody = response.Content.ReadAsStringAsync().Result;
                    InvoiceRa2 rootObject;
                    try
                    {
                        rootObject = JsonConvert.DeserializeObject<InvoiceRa2>(responseBody);
                    }
                    catch (Exception ex)
                    {
                        rootObject = JsonConvert.DeserializeObject<InvoiceRa2>(responseBody);
                    }

                    XulyRa2ML(rootObject, tokken,4);
                    while (!string.IsNullOrEmpty(rootObject.state))
                    {
                        url = $@"https://hoadondientu.gdt.gov.vn:30000/query/invoices/sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=50&state={rootObject.state}&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2}";
                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        try
                        {
                            response = client.GetAsync(url).Result;

                            // Đảm bảo phản hồi thành công
                            response.EnsureSuccessStatusCode();

                            // Đọc nội dung phản hồi
                            responseBody = response.Content.ReadAsStringAsync().Result;
                            rootObject = JsonConvert.DeserializeObject<InvoiceRa2>(responseBody);
                            XulyRa2ML(rootObject, tokken,4);
                        }
                        catch (Exception ex)
                        {
                            // Xử lý lỗi nếu cần
                        }
                    }

                    Console.WriteLine("Response Body:");
                    Console.WriteLine(responseBody);
                   
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Request error: {e.Message}");
                }
            }
        }
        public void Xulydaura2(string tokken)
        {
            using (var client = new HttpClient())
            {
                // Đặt URL
                string formattedDate1 = frmMain.dtFrom.ToString("dd/MM/yyyyTHH:mm:ss");
                string formattedDate2 = frmMain.dtTo.ToString("dd/MM/yyyyTHH:mm:ss");
                string url = $@"https://hoadondientu.gdt.gov.vn:30000/sco-query/invoices/sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=50&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2}";
                // Thêm Bearer token vào Header
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                try
                {
                    // Gửi yêu cầu GET
                    HttpResponseMessage response = client.GetAsync(url).Result;

                    // Đảm bảo phản hồi thành công
                    response.EnsureSuccessStatusCode();

                    // Đọc nội dung phản hồi
                    string responseBody = response.Content.ReadAsStringAsync().Result;
                    InvoiceRa2 rootObject;
                    try
                    {
                        rootObject = JsonConvert.DeserializeObject<InvoiceRa2>(responseBody);
                    }
                    catch (Exception ex)
                    {
                        rootObject = JsonConvert.DeserializeObject<InvoiceRa2>(responseBody);
                    }

                    XulyRa2ML(rootObject, tokken,5);
                    while (!string.IsNullOrEmpty(rootObject.state))
                    {
                        url = $@"https://hoadondientu.gdt.gov.vn:30000/sco-query/invoices/sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=50&state={rootObject.state}&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2}";
                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        try
                        {
                            response = client.GetAsync(url).Result;

                            // Đảm bảo phản hồi thành công
                            response.EnsureSuccessStatusCode();

                            // Đọc nội dung phản hồi
                            responseBody = response.Content.ReadAsStringAsync().Result;
                            rootObject = JsonConvert.DeserializeObject<InvoiceRa2>(responseBody);
                            XulyRa2ML(rootObject, tokken, 5);
                        }
                        catch (Exception ex)
                        {
                            // Xử lý lỗi nếu cần
                        }
                    }

                    Console.WriteLine("Response Body:");
                    Console.WriteLine(responseBody);
                    frmMain.tokken = tokken;
                    Driver.Close();
                    this.Close();
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Request error: {e.Message}");
                }
            }
        }
        public void Xulydauvao1(string tokken, int type)
        {
            using (var client = new HttpClient())
            {
                // Đặt URL
                string formattedDate1 = frmMain.dtFrom.ToString("dd/MM/yyyyTHH:mm:ss");
                string formattedDate2 = frmMain.dtTo.ToString("dd/MM/yyyyTHH:mm:ss");
                string url = "";
                int ttxly = 0;

                if (type == 6)
                {
                    ttxly = 5;
                }
                if (type == 8)
                {
                    ttxly = 6;
                }

                url = @"https://hoadondientu.gdt.gov.vn:30000/query/invoices/purchase?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=50&search=tdlap=ge=" + formattedDate1 + ";tdlap=le=" + formattedDate2 + ";ttxly==" + ttxly;

                // Thêm Bearer token vào Header
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                try
                {
                    // Gửi yêu cầu GET
                    HttpResponseMessage response = client.GetAsync(url).Result;

                    // Đảm bảo phản hồi thành công
                    response.EnsureSuccessStatusCode();

                    // Đọc nội dung phản hồi
                    string responseBody = response.Content.ReadAsStringAsync().Result;
                    RootObject rootObject = JsonConvert.DeserializeObject<RootObject>(responseBody);

                    XulyDataXML(rootObject, tokken, type);

                    while (!string.IsNullOrEmpty(rootObject.state))
                    {
                        url = @"https://hoadondientu.gdt.gov.vn:30000/query/invoices/purchase?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=50&state=" + rootObject.state + "&search=tdlap=ge=" + formattedDate1 + ";tdlap=le=" + formattedDate2 + ";ttxly==5";
                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                        response = client.GetAsync(url).Result;

                        // Đảm bảo phản hồi thành công
                        response.EnsureSuccessStatusCode();

                        // Đọc nội dung phản hồi
                        responseBody = response.Content.ReadAsStringAsync().Result;
                        rootObject = JsonConvert.DeserializeObject<RootObject>(responseBody);
                    }

                    Console.WriteLine("Response Body:");
                    Console.WriteLine(responseBody);
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Request error: {e.Message}");
                }
            }
        }
        public async Task Xulydauvaomaytinhtien(string tokken, int type)
        {
            using (var client = new HttpClient())
            {
                // Đặt URL
                string formattedDate1 = frmMain.dtFrom.ToString("dd/MM/yyyyTHH:mm:ss");
                string formattedDate2 = frmMain.dtTo.ToString("dd/MM/yyyyTHH:mm:ss");
                string url = "";
                int ttxly = 0;
                if (type == 10)
                {
                    ttxly = 8;
                }
                url = @"https://hoadondientu.gdt.gov.vn:30000/sco-query/invoices/purchase?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=50&search=tdlap=ge=" + formattedDate1 + ";tdlap=le=" + formattedDate2 + ";ttxly==" + ttxly;
                // Thêm Bearer token vào Header
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                try
                {
                    // Gửi yêu cầu GET
                    HttpResponseMessage response=null;

                    try
                    {
                        response = await client.GetAsync(url);
                        if (response.StatusCode == System.Net.HttpStatusCode.InternalServerError)
                        {
                            XtraMessageBox.Show("Lỗi máy chủ, vui lòng thử lại sau.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Driver.Close();
                            this.Close();   
                        }
                    }
                    catch(Exception ex)
                    {
                       
                    }

                    // Đảm bảo phản hồi thành công
                    response.EnsureSuccessStatusCode();

                    // Đọc nội dung phản hồi
                    string responseBody = await response.Content.ReadAsStringAsync();
                    RootObject rootObject;
                    try
                    {
                        rootObject = JsonConvert.DeserializeObject<RootObject>(responseBody);
                    }
                    catch (Exception ex)
                    {
                        rootObject = JsonConvert.DeserializeObject<RootObject>(responseBody);
                    }

                    XulyDataXML(rootObject, tokken, type);
                    while (!string.IsNullOrEmpty(rootObject.state))
                    {
                        url = @"https://hoadondientu.gdt.gov.vn:30000/sco-query/invoices/purchase?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=50&state=" + rootObject.state + "&search=tdlap=ge=" + formattedDate1 + ";tdlap=le=" + formattedDate2 + ";ttxly==5";
                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        try
                        {
                            response = await client.GetAsync(url);

                            // Đảm bảo phản hồi thành công
                            response.EnsureSuccessStatusCode();

                            // Đọc nội dung phản hồi
                            responseBody = await response.Content.ReadAsStringAsync();
                            rootObject = JsonConvert.DeserializeObject<RootObject>(responseBody);
                            XulyDataXML(rootObject, tokken, type);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    Console.WriteLine("Response Body:");
                    Console.WriteLine(responseBody); 
                    Driver.Close();
                    this.Close();
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Request error: {e.Message}");
                }
            }
        }
        string dbPath = "";
        private DataTable ExecuteQuery(string query, params OleDbParameter[] parameters)
        {
            DataTable dataTable = new DataTable();
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
            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
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
        public void GetdetailXML2(string nbmst, string khhdon, string shdon, string tokken)
        {
            string url = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/detail?nbmst={nbmst}&khhdon={khhdon}&shdon={shdon}&khmshdon=1";

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                try
                {
                    // Gửi yêu cầu GET đồng bộ
                    Thread.Sleep(400);
                    HttpResponseMessage response = client.GetAsync(url).GetAwaiter().GetResult();

                    // Đảm bảo phản hồi thành công
                    response.EnsureSuccessStatusCode();

                    // Đọc nội dung phản hồi đồng bộ
                    string responseBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    var rootObject = JsonConvert.DeserializeObject<Invoice>(responseBody);

                    // Tìm ID Cha mới nhất
                    string query = "SELECT * FROM tbimport WHERE SHDon=? AND KHHDon=? AND Mst=?";
                    var parameterss = new OleDbParameter[]
                    {
                new OleDbParameter("?", shdon),
                new OleDbParameter("?", khhdon),
                new OleDbParameter("?", nbmst)
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
                            catch(Exception ex)
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
        string password, connectionString;
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
        public  void XulyDataXML(RootObject rootObject,string tokken,int invoceType)
        {
            foreach(var item in rootObject.datas)
            {
                InsertTbImport(item, invoceType);
                // GetdetailXML(item.nbmst, item.khhdon, item.shdon.ToString(),tokken);
            }
        }
        public void XulyRa2ML(InvoiceRa2 rootObject, string tokken,int invoceType)
        {
            foreach (var item in rootObject.datas)
            {
                InsertTbImport2(item,tokken, invoceType);
               
            }
        }
       
        private bool CheckExistKH(string mst, string ten)
        {
            if (!string.IsNullOrEmpty(mst))
            {
                if (existingKhachHang.AsEnumerable().Any(row => row.Field<string>("MST") == mst))
                {
                    return true;
                }
            }
            else
            {
                if (existingKhachHang.AsEnumerable().Any(row => Helpers.RemoveVietnameseDiacritics(row.Field<string>("Ten").ToLower()) == Helpers.RemoveVietnameseDiacritics(ten.ToLower())))
                {
                    return true;
                }
            }

            return false;
        }
        DataTable existingKhachHang = new DataTable();
        DataTable existingTbImport = new DataTable();
        string csohieu = "";
        public void InitCustomer(int Maphanloai, string Sohieu, string Ten, string Diachi, string Mst)
        {
            int randNumber = 0;
            Random random = new Random();

            //Xử lý địa chỉ
            string diachiKHVni = !string.IsNullOrEmpty(Diachi) ? Helpers.ConvertUnicodeToVni(Diachi) : Helpers.ConvertUnicodeToVni("Bổ sung địa chỉ");

            if (string.IsNullOrEmpty(Mst))
            {
                Sohieu = Helpers.RemoveVietnameseDiacritics(Ten.Split(' ').LastOrDefault());
                //Sohieu = CapitalizeFirstLetter(Sohieu);
                if (Sohieu.Length >= 3)
                    Sohieu = Sohieu.ToUpper().Substring(0, 3);
                 randNumber = random.Next(101,999);
                Sohieu = Sohieu + randNumber.ToString();
                csohieu = Sohieu;
                Mst = "00"; 
            }
            else
            {
                Sohieu = Helpers.GetLastFourDigits(Mst.Replace("-", ""));

                string tenKHVni = Helpers.ConvertUnicodeToVni(Ten);
                 
                //Xử lý khi số hiệu bị trùng
                if (existingKhachHang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                {
                    Sohieu = "0" + Sohieu;
                }
                if (existingKhachHang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                {
                    Sohieu = "00" + Sohieu;
                }
            }
             
            string query = @"
        INSERT INTO KhachHang (MaPhanLoai,SoHieu,Ten,DiaChi,MST)
        VALUES (?,?,?,?,?)";


            // Khai báo mảng tham số với đủ 10 tham số
            OleDbParameter[] parameters = new OleDbParameter[]
            {
        new OleDbParameter("?", Maphanloai),
          new OleDbParameter("?", Sohieu),
        new OleDbParameter("?", Ten),
        new OleDbParameter("?", diachiKHVni),
        new OleDbParameter("?", Mst)
            };

            // Thực thi truy vấn và lấy kết quả
            int a = ExecuteQueryResult(query, parameters);
        }
        public void InsertTbImport(Data item,int invoceType)
        {
            if (existingTbImport.AsEnumerable().Any(row => row.Field<string>("SHDon").ToString() == item.shdon.ToString() && row.Field<string>("KHHDon").ToString() == item.khhdon.ToString()))
            {
                return;
            }

            int type = frmMain.type;
            string query = @"
            INSERT INTO tbImport (SHDon, KHHDon, NLap, Ten, Noidung,TKNo,TKCo, TkThue, Mst, Status, Ngaytao, TongTien, Vat, SohieuTP,TPhi,TgTCThue,TgTThue,Type,InvoiceType,IsHaschild)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?,?,?,?)";

            string newTen = Helpers.ConvertUnicodeToVni(item.nbten);
            string newNoidung = "";
            //Lấy tài khoản từ mất định
            string tkno = "";
            string tkco = "";
            string tkthue = "";
            string querykh = @" SELECT *  FROM tbDinhdanhtaikhoan"; // Sử dụng ? thay cho @mst trong OleDb
            if (CheckExistKH(item.nmmst, newTen) == false)
            {
                int maphanloai = 0;
                maphanloai = type == 1 ? 2 : 3; //1 là mua, 2 là bán 
                InitCustomer(maphanloai, item.khhdon, newTen, item.nmdchi, item.nmmst);
            }
            var result = ExecuteQuery(querykh, new OleDbParameter("?", ""));



            if (result.Rows.Count > 0)
            {
                foreach (DataRow row in result.Rows)
                {
                    if (type == 1)
                    {
                        if (row["KeyValue"].ToString().Contains("Ưu tiên vào"))
                        {
                            tkno = row["TKNo"].ToString();
                            tkco = row["TKCo"].ToString();
                            tkthue = row["TKThue"].ToString();
                            break;
                        }
                    }
                    if (type == 2)
                    {
                        if (row["KeyValue"].ToString().Contains("Ưu tiên ra"))
                        {
                            tkno = row["TKNo"].ToString();
                            tkco = row["TKCo"].ToString();
                            tkthue = row["TKThue"].ToString();
                            break;
                        }
                    }

                }
            }

            string tgtkcthue = "";
            if(item.tgtcthue != null)
            {
                tgtkcthue = item.tgtcthue.ToString();
            }
            else
            {
                if (item.tgtkcthue != null)
                {
                    tgtkcthue = item.tgtkcthue.ToString();
                }
                else
                {
                    if (item.tgtkcthue != null)
                    {
                        tgtkcthue = item.tgtkcthue.ToString();
                    }
                    else
                    {
                        tgtkcthue = "0";
                    }
                }
            }
            DateTime nl = new DateTime();
            if(item.shdon == 39519)
            {
                int a = 10;
            }
            if (invoceType == 8)
            {
                if (item.ntao.Month != DateTime.Now.Month)
                {

                    nl = item.ntao;
                }
                else
                {
                    nl = item.tdlap.AddDays(1);
                }
            }
            else
            {
                nl = item.ntao;
            }
            string getMST = "";
            if (item.nmmst != null)
            {
                getMST = item.nmmst;
            }
            else
            {
                //Lấy ma số thuế từ số hiệu khách hàng
                querykh = @" SELECT *  FROM KhachHang where Ten=?"; // Sử dụng ? thay cho @mst trong OleDb
                result = ExecuteQuery(querykh, new OleDbParameter("?", newTen));
                if (result.Rows.Count > 0)
                {
                    getMST = result.Rows[0]["SoHieu"].ToString();
                }
            }
            string dateTimeString = nl.ToString();
            DateTime dateTime = DateTime.Parse(dateTimeString);
            string formattedDate = dateTime.ToShortDateString();
            OleDbParameter[] parameters = new OleDbParameter[]
                {
                new OleDbParameter("?", item.shdon),
                new OleDbParameter("?", item.khhdon),
                new OleDbParameter("?", formattedDate),
                new OleDbParameter("?", newTen),
                new OleDbParameter("?", newNoidung),
                new OleDbParameter("?",tkno),
                new OleDbParameter("?",tkco),
                new OleDbParameter("?",tkthue),
                new OleDbParameter("?", item.nbmst),
                new OleDbParameter("?", "0"),
                new OleDbParameter("?", DateTime.Now.ToShortDateString()),
                new OleDbParameter("?", item.tgtttbso),
                new OleDbParameter("?", "0"),
                new OleDbParameter("?",""),
                new OleDbParameter("?", "0"),
                new OleDbParameter("?",tgtkcthue!=null?tgtkcthue:item.tgtttbso.ToString()),
                new OleDbParameter("?",item.tgtthue!=null? item.tgtthue.ToString():"0"),
                new OleDbParameter("?", type),
                new OleDbParameter("?", invoceType),
                 new OleDbParameter("?","1")
                };

            try
            {
                int a = ExecuteQueryResult(query, parameters);
            }
            catch (Exception ex)
            {

            }

        }
        public void InsertTbImport2(InvoiceRa2List item, string tokken, int invoceType)
        {
            //Kiểm tra tồn tại trước khi thêm mới
            if (existingTbImport.AsEnumerable().Any(row => row.Field<string>("SHDon").ToString() == item.shdon.ToString() && row.Field<string>("KHHDon").ToString() == item.khhdon.ToString()))
            {
                return;
            }
            int type = frmMain.type;
            string query = @"
            INSERT INTO tbImport (SHDon, KHHDon, NLap, Ten, Noidung,TKNo,TKCo, TkThue, Mst, Status, Ngaytao, TongTien, Vat, SohieuTP,TPhi,TgTCThue,TgTThue,Type,InvoiceType,IsHaschild)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?,?,?,?)";

            string newTen = "";
            if (!string.IsNullOrEmpty(item.nmten))
            {
                newTen = Helpers.ConvertUnicodeToVni(item.nmten);
            }
            else
            {
                if (!string.IsNullOrEmpty(item.nmtnmua))
                {
                    newTen = Helpers.ConvertUnicodeToVni(item.nmtnmua);
                }
            }
            //Insert khach hàng
            if (CheckExistKH(item.nmmst, newTen) == false)
            {
                int maphanloai = 0;
                maphanloai = type == 1 ? 2 : 3; //1 là mua, 2 là bán 
                InitCustomer(maphanloai, item.khhdon, newTen, item.nmdchi, item.nmmst);
            }
            string newNoidung = Helpers.ConvertUnicodeToVni("");
            //Lấy tài khoản từ mất định
            string tkno = "";
            string tkco = "";
            string tkthue = "";
            string querykh = @" SELECT *  FROM tbDinhdanhtaikhoan"; // Sử dụng ? thay cho @mst trong OleDb

            var result = ExecuteQuery(querykh, new OleDbParameter("?", ""));
            if (result.Rows.Count > 0)
            {
                foreach (DataRow row in result.Rows)
                {
                    if (type == 1)
                    {
                        if (row["KeyValue"].ToString().Contains("Ưu tiên vào"))
                        {
                            tkno = row["TKNo"].ToString();
                            tkco = row["TKCo"].ToString();
                            tkthue = row["TKThue"].ToString();
                            break;
                        }
                    }
                    if (type == 2)
                    {
                        if (row["KeyValue"].ToString().Contains("Ưu tiên ra"))
                        {
                            tkno = row["TKNo"].ToString();
                            tkco = row["TKCo"].ToString();
                            tkthue = row["TKThue"].ToString();
                            break;
                        }
                    }

                }
            }

            string tgtkcthue = "";
            string dateTimeString = item.ntao;
            DateTime dateTime = DateTime.Parse(dateTimeString);
            string formattedDate = dateTime.ToShortDateString();
            string getMST = "";
            if (item.nmmst != null)
            {
                getMST = item.nmmst;
            }
            else
            {
                //Lấy ma số thuế từ số hiệu khách hàng
                querykh = @" SELECT *  FROM KhachHang where Ten=?"; // Sử dụng ? thay cho @mst trong OleDb
                result = ExecuteQuery(querykh, new OleDbParameter("?", newTen));
                if (result.Rows.Count > 0)
                {
                    getMST = result.Rows[0]["SoHieu"].ToString();
                }
            }
            Console.WriteLine(formattedDate); // Kết quả: 2025-05-22
            OleDbParameter[] parameters = new OleDbParameter[]
            {
                new OleDbParameter("?", item.shdon),
                new OleDbParameter("?", item.khhdon),
                new OleDbParameter("?", formattedDate),
                new OleDbParameter("?", newTen),
                new OleDbParameter("?", newNoidung),
                new OleDbParameter("?",tkno),
                new OleDbParameter("?",tkco),
                new OleDbParameter("?",tkthue),
                new OleDbParameter("?",getMST),
                new OleDbParameter("?", "0"),
                new OleDbParameter("?", DateTime.Now.ToShortDateString()),
                new OleDbParameter("?", item.tgtttbso),
                new OleDbParameter("?", "0"),
                new OleDbParameter("?",""),
                new OleDbParameter("?", "0"),
                new OleDbParameter("?", item.tgtcthue!=null?item.tgtcthue.ToString():"0"),
                new OleDbParameter("?", item.tgtthue.ToString()),
                new OleDbParameter("?", type),
                new OleDbParameter("?", invoceType),
               new OleDbParameter("?", "1")
            };

            try
            {
                int a = ExecuteQueryResult(query, parameters);
                //Xu lý khách hàng

                //Update lại MST cho tbimport 


                //  GetdetailXML2(item.nbmst, item.khhdon, item.shdon.ToString(), tokken);
            }
            catch (Exception ex)
            {
                var a = ex.Message;
            }

        }
        private void frmTaiCoQuanThue_Load(object sender, EventArgs e)
        {
            var query = @"SELECT * FROM PhanLoaiVattu ORDER BY TenPhanLoai";
            var dt = ExecuteQuery(query, null);
            query = "SELECT * FROM KhachHang"; // Giả sử bạn muốn lấy tất cả dữ liệu từ bảng KhachHang
            existingKhachHang = ExecuteQuery(query);
            //existingTbImport
            query = "SELECT * FROM tbimport"; // Giả sử bạn muốn lấy tất cả dữ liệu từ bảng KhachHang
            existingTbImport = ExecuteQuery(query);
            // XulyFileExcel();
            // return;
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
                if (frmMain.type == 1)
                {
                    downloadPath = frmMain.savedPath + "\\HDVao"; 
                }
                if (frmMain.type == 2)
                {
                    downloadPath = frmMain.savedPath + "\\HDRa"; 
                }
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
                    var usernameField = Driver.FindElement(By.Id("username"));
                    var passwordField = Driver.FindElement(By.Id("password"));
                    string username = frmMain.username;
                    string password = frmMain.pasword;
                    usernameField.SendKeys(username);
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
                    catch (Exception ex)
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


                 


                    var cookies = Driver.Manage().Cookies.AllCookies.Where(m=>m.Name== "jwt");

                    foreach (var cookie in cookies)
                    {
                        Console.WriteLine($"Name: {cookie.Name}, Value: {cookie.Value}");
                        //Lưu tokken
                         query = "UPDATE tbRegister SET tokken=? ";
                        var parametersss = new OleDbParameter[]
                        { 
                            new OleDbParameter("?", cookie.Value),
                        };

                        frmMain.tokken = cookie.Value;
                        ExecuteQueryResult(query, parametersss);
                        if(frmMain.type==1)
                        {
                            Xulydauvao1(cookie.Value, 6);
                            Xulydauvao1(cookie.Value, 8);
                            Xulydauvaomaytinhtien(cookie.Value, 10);
                        }
                        if (frmMain.type == 2)
                        {
                            Xulydaura1(cookie.Value);
                            Xulydaura2(cookie.Value);
                        }
                     
                    }
 
                }
                catch (Exception ex)
                { 
                    Driver.Close();
                    // MessageBox.Show($"Lỗi: {ex.Message}");
                }
            }
             
            // Thêm cột với tiêu đề 
           
        }

        private void spreadsheetControl1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue ==13)
            {
                //var cell = spreadsheetControl1.ActiveCell;
                //string value = spreadsheetControl1.ActiveWorksheet.GetCellValue(0,0).ToString();

                //// In ra console hoặc hiển thị thông báo
                //Console.WriteLine($"Enter tại {cell.GetReferenceA1()}, giá trị: {value}");

                //// (Tùy chọn) Ngăn di chuyển xuống ô khác
                //e.Handled = true;
                e.Handled = true;
                frmTaikhoan frmTaikhoan = new frmTaikhoan();
                frmTaikhoan.Show();

            }
        }
    }
}