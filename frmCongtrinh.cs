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
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.OleDb;
using System.Configuration;
using DevExpress.XtraGrid.Views.Grid;
using System.Reflection;
using System.IO;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Text.RegularExpressions;
using DevExpress.XtraEditors.Mask.Design;

namespace SaovietTax
{
    public partial class frmCongtrinh : DevExpress.XtraEditors.XtraForm
    {
        public frmCongtrinh()
        {
            InitializeComponent();
        }
        private void LoadData()
        {
            txtId.Text = "0";
            string query = @" SELECT *  FROM TP154 ";
            var kq = ExecuteQuery(query, null);
            for (int i = 0; i < kq.Rows.Count; i++)
            {
                kq.Rows[i]["TenVattu"] = Helpers.ConvertVniToUnicode(kq.Rows[i]["TenVattu"].ToString());
                kq.Rows[i]["DonVi"] = Helpers.ConvertVniToUnicode(kq.Rows[i]["DonVi"].ToString());
                kq.Rows[i]["GhiChu"] = Helpers.ConvertVniToUnicode(kq.Rows[i]["GhiChu"].ToString());
            }
            gridControl1.DataSource = kq;
        }
        private void frmCongtrinh_Load(object sender, EventArgs e)
        {
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
        public frmMain frmMain;
        private void gridControl1_DoubleClick(object sender, EventArgs e)
        {
            GridView gridView = gridControl1.MainView as GridView;
            var hitInfo = gridView.CalcHitInfo(gridView.GridControl.PointToClient(MousePosition));


            // Kiểm tra nếu nhấp vào một ô
            if (hitInfo.InRowCell)
            {
                int columnIndex = hitInfo.Column.VisibleIndex; // Chỉ số cột
                
                // Lấy giá trị trong ô đã nhấp
                var hiddenValue = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["SoHieu"]);
                frmMain.hiddenValue = hiddenValue.ToString();
                this.Close();
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (txtId.Text == "0")
            {
                double matk = 0;
                double tk = 0;
                string query = @" SELECT *  FROM HeTHongTK where SoHieu= ? ";
                var parameterss = new OleDbParameter[]
                {
                new OleDbParameter("?",txtTaikhoan.Text)
                   };
                var kq = ExecuteQuery(query, parameterss);
                if(kq.Rows.Count==0)
                {
                    XtraMessageBox.Show("Mã tài khoản không hợp lệ");
                    return;
                }
                else
                {
                    matk = int.Parse(kq.Rows[0]["MaTC"].ToString());
                }
                    query = @"
        INSERT INTO TP154 (MaPhanLoai,SoHieu,TenVattu,DonVi,GhiChu,DK,MaTK)
        VALUES (?,?,?,?,?,?,?)";


                // Khai báo mảng tham số với đủ 10 tham số
                OleDbParameter[] parameters = new OleDbParameter[]
                {
        new OleDbParameter("?", 1),
          new OleDbParameter("?", txtSohieu.Text),
        new OleDbParameter("?", txtTenvattu.Text),
        new OleDbParameter("?", txtDonvi.Text), 
          new OleDbParameter("?", txtGhichu.Text),
             new OleDbParameter("?",tk),
               new OleDbParameter("?",matk),
                };

                // Thực thi truy vấn và lấy kết quả
                int a = ExecuteQueryResult(query, parameters);
                LoadData();
            }
            else
            {
                double matk = 0;
                double tk = 0;
                string query = @" SELECT *  FROM HeTHongTK where SoHieu= ? ";
                var parameterss = new OleDbParameter[]
                {
                new OleDbParameter("?",txtTaikhoan.Text)
                   };
                var kq = ExecuteQuery(query, parameterss);
                if (kq.Rows.Count == 0)
                {
                    XtraMessageBox.Show("Mã tài khoản không hợp lệ");
                    return;
                }
                else
                {
                    matk = int.Parse(kq.Rows[0]["MaTC"].ToString());
                }
                query = @"
        Update TP154 set SoHieu=?,TenVattu=?,DonVi=?,GhiChu=?,MaTK=? where MaSo=?";
                  
                // Khai báo mảng tham số với đủ 10 tham số
                OleDbParameter[] parameters = new OleDbParameter[]
                { 
          new OleDbParameter("?", txtSohieu.Text),
        new OleDbParameter("?", txtTenvattu.Text),
        new OleDbParameter("?", txtDonvi.Text),
          new OleDbParameter("?", txtGhichu.Text), 
               new OleDbParameter("?",matk),
                 new OleDbParameter("?",txtId.Text),
                };

                // Thực thi truy vấn và lấy kết quả
                int a = ExecuteQueryResult(query, parameters);
                LoadData();
            }
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        bool isAdd = true;
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            isAdd = true;
            txtId.Text = "0";
            txtSohieu.Text = "";
            txtTenvattu.Text = "";
            txtDonvi.Text = "";
            txtDonvi.Text = "";
            txtGhichu.Text = "";
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            string query = @" Delete from TP154  where MaSo=?";

            // Khai báo mảng tham số với đủ 10 tham số
            OleDbParameter[] parameters = new OleDbParameter[]
            { 
                 new OleDbParameter("?",txtId.Text),
            };
            int a = ExecuteQueryResult(query, parameters);
            LoadData();
        }

        private void gridView1_RowClick(object sender, RowClickEventArgs e)
        {
            // Lấy chỉ số hàng đã click
            int rowHandle = e.RowHandle;

            // Lấy dữ liệu từ hàng
            var value = gridView1.GetRowCellValue(rowHandle, "SoHieu").ToString();
            txtSohieu.Text = value;
            txtTenvattu.Text = gridView1.GetRowCellValue(rowHandle, "TenVattu").ToString();
            txtDonvi.Text = gridView1.GetRowCellValue(rowHandle, "DonVi").ToString();
            txtGhichu.Text = gridView1.GetRowCellValue(rowHandle, "GhiChu").ToString();
            txtId.Text = gridView1.GetRowCellValue(rowHandle, "MaSo").ToString();
            var matc = gridView1.GetRowCellValue(rowHandle, "MaTK").ToString();
            string query = @" SELECT *  FROM HeTHongTK where MaTC= ? ";
            var parameterss = new OleDbParameter[]
            {
                new OleDbParameter("?",matc) 
               };
            var kq = ExecuteQuery(query, parameterss);
            txtTaikhoan.Text = kq.Rows[0]["SoHieu"].ToString();
            // Thực hiện hành động với dữ liệu
            isAdd = false;
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