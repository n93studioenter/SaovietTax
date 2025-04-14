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
using System.Configuration;
using DevExpress.Internal.WinApi.Windows.UI.Notifications;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Columns;

namespace SaovietTax
{
	public partial class frmDinhdanh: DevExpress.XtraEditors.XtraForm
	{
        public frmDinhdanh()
		{
            InitializeComponent();
		}
        string dbPath = "";
        public DataTable result { get; set; }
        private DataTable ExecuteQuery(string query, params OleDbParameter[] parameters)
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

                    using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                    {
                        dataAdapter.Fill(dataTable);
                    }
                }
            }

            return dataTable; // Trả về DataTable chứa dữ liệu
        }
        private void frmDinhdanh_Load(object sender, EventArgs e)
        {
            dbPath = ConfigurationManager.AppSettings["dbpath"];
            InitDB();
            LoadDataDinhDanh();
        }
        private void LoadDataDinhDanh()
        {
            if ( string.IsNullOrEmpty(dbPath))
                return;
            string querykh = @" SELECT *  FROM tbDinhdanhtaikhoan"; // Sử dụng ? thay cho @mst trong OleDb

            result = ExecuteQuery(querykh, new OleDbParameter("?", ""));
            gcDinhdanh.DataSource = result;
            GridView gridView = gcDinhdanh.MainView as GridView;

            // Tạo cột xóa
            
            gridView.CustomUnboundColumnData += gridView_CustomUnboundColumnData;
            gridView.RowCellClick += gridView_RowCellClick;

        }
        private void gridView_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
            if (e.Column.FieldName == "colDelete" )
            {
                e.Value = "Xóa";
            }
        }
        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e.Column.FieldName == "colDelete" )
            {
                var rowHandle = e.RowHandle;
                GridView gridView = gcDinhdanh.MainView as GridView;
                if (gridView.GetRowCellValue(rowHandle, "ID") == null)
                    return;
                // Ví dụ: Lấy giá trị của một cột có tên "Name" từ hàng hiện tại
                string nameValue = gridView.GetRowCellValue(rowHandle, "ID").ToString();
                if (XtraMessageBox.Show("Bạn có chắc chắn muốn xóa hàng này?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    string sql = "DELETE FROM tbDinhdanhtaikhoan WHERE ID = @AccountID";
                    OleDbParameter[] parameters = new OleDbParameter[]
                {
        new OleDbParameter("?", nameValue),
                };
                    int resl = ExecuteQueryResult(sql, parameters);
                    LoadDataDinhDanh();
                } 
            }
        }
        private void InitDB()
        {
            // Đường dẫn đến cơ sở dữ liệu Access và mật khẩu
            //dbPath = @"C:\S.T.E 25\S.T.E 25\DATA\KT2025.mdb"; // Thay đổi đường dẫn này
            dbPath = ConfigurationManager.AppSettings["dbpath"]; 
            string filePath = ConfigurationManager.AppSettings["dbpath"];
            if (string.IsNullOrEmpty(filePath))
                return;
            // Đọc toàn bộ nội dung tệp
            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            //connectionString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            // connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database";
            //connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\S.T.E 25\S.T.E 25\DATA\importData.accdb;Persist Security Info=False";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Có lỗi xảy ra: {ex.Message}");
            }

        }
        string password, connectionString;
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
    }
}