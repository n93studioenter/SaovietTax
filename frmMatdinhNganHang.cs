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
using static DevExpress.Xpo.Helpers.AssociatedCollectionCriteriaHelper;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Views.Grid;
namespace SaovietTax
{
    public partial class frmMatdinhNganHang : DevExpress.XtraEditors.XtraForm
    {
        public frmMatdinhNganHang()
        {
            InitializeComponent();
        }

        private void frmMatdinhNganHang_Load(object sender, EventArgs e)
        {
            LoadData(); 
        }
        private void LoadData()
        {
            string querykh = @" SELECT *  FROM tbDinhdanhNganhang"; // Sử dụng ? thay cho @mst trong OleDb

            var result = ExecuteQuery(querykh, new OleDbParameter("?", ""));
            gcDinhdanh.DataSource = result;
        }
        private void btnLuudinhdanh_Click(object sender, EventArgs e)
        {
            var query = @"INSERT INTO tbDinhdanhNganhang (Noidung,TK) VALUES (?, ?)";
            var parameters = new OleDbParameter[]
            {
            new OleDbParameter("?", txtDiengiai.Text), 
            new OleDbParameter("?", textEdit1.Text)
            };
            int rowsAffected = ExecuteQueryResult(query, parameters);
            LoadData();
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

        private void gridView2_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridView gridView = gcDinhdanh.MainView as GridView;
            // Lấy thông tin về hàng và cột của ô đã thay đổi
            int rowHandle = e.RowHandle;
            string columnName = e.Column.FieldName; // Tên cột
            //Lấy current data row
            int ID = int.Parse(gridView.GetRowCellValue(rowHandle, gridView.Columns["ID"]).ToString());
            string Noidung = gridView.GetRowCellValue(rowHandle, gridView.Columns["Noidung"]).ToString();
            string TK = gridView.GetRowCellValue(rowHandle, gridView.Columns["TK"]).ToString();  
            string sql = "UPDATE tbDinhdanhNganhang SET Noidung = ?, TK = ?  WHERE ID = ?";
            OleDbParameter[] parameters = new OleDbParameter[]
 { 
           new OleDbParameter("?",Noidung),
                 new OleDbParameter("?",TK), 
                new OleDbParameter("?",ID)
 };
            int resl = ExecuteQueryResult(sql, parameters);
        }

        private void gridView2_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            var getNAmecol = e.Column;
            if (getNAmecol.ToString() == "Xóa")
            {
                if (XtraMessageBox.Show("Bạn có chắc chắn muốn xóa hàng này?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    var getID = gridView2.GetRowCellValue(e.RowHandle, "ID");
                    string sql = "DELETE FROM tbDinhdanhNganhang WHERE ID = ?";
                    OleDbParameter[] parameters = new OleDbParameter[]
                {
        new OleDbParameter("?", getID),
                };
                    int resl = ExecuteQueryResult(sql, parameters);
                    LoadData();
                }
               
            }
        }
    }
}