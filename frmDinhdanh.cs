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
using System.Reflection;
using System.IO;

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
            comboBoxEdit1.Properties.Buttons[0].Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Combo;

            comboBoxEdit1.Properties.Items.AddRange(new string[]
         {
            "Thấp",
            "Cao" 
         });
            comboBoxEdit1.SelectedIndex = 0;
            comboBoxEdit1.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;

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
            gridView.CellValueChanged += GridView_CellValueChanged;
        }
        private void GridView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridView gridView = gcDinhdanh.MainView as GridView;
            // Lấy thông tin về hàng và cột của ô đã thay đổi
            int rowHandle = e.RowHandle;
            string columnName = e.Column.FieldName; // Tên cột
            //Lấy current data row
            int ID = int.Parse(gridView.GetRowCellValue(rowHandle, gridView.Columns["ID"]).ToString());
            string Type = gridView.GetRowCellValue(rowHandle, gridView.Columns["Type"]).ToString();
            string KeyValue = gridView.GetRowCellValue(rowHandle, gridView.Columns["KeyValue"]).ToString();
            string TKNo = gridView.GetRowCellValue(rowHandle, gridView.Columns["TKNo"]).ToString();
            string TKCo = gridView.GetRowCellValue(rowHandle, gridView.Columns["TKCo"]).ToString();
            string TKThue = gridView.GetRowCellValue(rowHandle, gridView.Columns["TKThue"]).ToString();
            string sql = "UPDATE tbDinhdanhtaikhoan SET Type = ?, KeyValue = ?, TKNo = ?, TKCo = ?, TKThue = ? WHERE ID = ?";
            OleDbParameter[] parameters = new OleDbParameter[]
{
        new OleDbParameter("?",Type),
           new OleDbParameter("?",KeyValue),
                 new OleDbParameter("?",TKNo),
             new OleDbParameter("?",TKCo),
              new OleDbParameter("?",TKThue),
                new OleDbParameter("?",ID)
};
            int resl = ExecuteQueryResult(sql, parameters);
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