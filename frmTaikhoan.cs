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
using System.Data.OleDb;
using System.Reflection;
using System.IO;

namespace SaovietTax
{
	public partial class frmTaikhoan: DevExpress.XtraEditors.XtraForm
	{
        public frmTaikhoan()
		{
            InitializeComponent();
		}
        private List<Item> CreateData()
        {
            return new List<Item>
    {
        new Item { Id = 1, Name = "Root 1", ParentId = 0 },
        new Item { Id = 2, Name = "Child 1.1", ParentId = 1 },
        new Item { Id = 3, Name = "Child 1.1.1", ParentId = 2 }, // Con của Child 1.1
        new Item { Id = 4, Name = "Child 1.1.2", ParentId = 2 }, // Con của Child 1.1
        new Item { Id = 5, Name = "Child 1.2", ParentId = 1 },
        new Item { Id = 6, Name = "Root 2", ParentId = 0 },
        new Item { Id = 7, Name = "Child 2.1", ParentId = 6 },
        new Item { Id = 8, Name = "Child 2.1.1", ParentId = 7 } // Con của Child 2.1
    };
        }
        public class Item
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public int ParentId { get; set; } // Để xác định cấu trúc cây
        }
        private void frmTaikhoan_Load(object sender, EventArgs e)
        {
            var data = CreateData();
            treeList1.DataSource = data;
            treeList1.ParentFieldName = "ParentId"; // Thiết lập mối quan hệ cha-con
            treeList1.KeyFieldName = "Id"; // Thuộc tính khóa
            LoadList();
        }
        private void LoadList()
        {
            string query = $"SELECT SoHieu  " +
                  "FROM HeThongTK " +
                  "WHERE LEFT(SoHieu, 1) <> '#' AND Loai = ? " +
                  "AND Cap > 0 AND MaNT <= 0 " +
                  "GROUP BY HeThongTK.SoHieu, HeThongTK.MaNT " +
                  "ORDER BY HeThongTK.SoHieu, HeThongTK.MaNT";
                var parameterss = new OleDbParameter[]
             {
                    new OleDbParameter("?","1"), 
                };
                var kq = ExecuteQuery(query, parameterss);
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
            string connectionString = "";
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
        private int ExecuteQueryResult(string query, params OleDbParameter[] parameters)
        {
            string connectionString = "";
            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
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
    }
}