﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using System.Data.OleDb;
using System.Reflection;
using System.IO;
using static SaovietTax.frmKhachhang;

namespace SaovietTax
{
	public partial class frmKhachhang: DevExpress.XtraEditors.XtraForm
	{
        public frmKhachhang()
		{
            InitializeComponent();
		}
        public class Khachhang
        {
            public int MaSo { get; set; }
            public int MaPhanLoai { get; set; }
            public string SoHieu { get; set; }
            public string Ten { get; set; }
            public string Mst { get; set; }
            public string DiaChi { get; set; }  
        }
        public Khachhang dtoVatTu { get; set; }
        public frmMain frmMain;
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
        public class Item
        {
            public string Name { get; set; }
            public int Id { get; set; }

            public override string ToString()
            {
                return Name; // Hiển thị tên trong ComboBox
            }
        }
        private void frmKhachhang_Load(object sender, EventArgs e)
        {
            string query = @"SELECT * FROM PhanLoaiKhachHang ORDER BY TenPhanLoai";
            var dt = ExecuteQuery(query, null);
            if (dt != null && dt.Rows.Count > 0)
            {
                comboBoxEdit1.Properties.Items.Clear(); // Xóa các mục cũ

                foreach (DataRow row in dt.Rows)
                {
                    // Thêm từng mục vào ComboBoxEdit
                    comboBoxEdit1.Properties.Items.Add(new Item
                    {
                        Name = Helpers.ConvertVniToUnicode(row["SoHieu"].ToString()) + " - " + Helpers.ConvertVniToUnicode(row["TenPhanLoai"].ToString()),
                        Id = Convert.ToInt32(row["MaSo"])
                    });
                }

                comboBoxEdit1.Properties.NullText = "Chọn Tài khoản";
                comboBoxEdit1.Properties.TextEditStyle = TextEditStyles.DisableTextEditor; // Ngăn người dùng nhập trực tiếp
                if (comboBoxEdit1.Properties.Items.Count > 0)
                {
                    comboBoxEdit1.SelectedIndex = 0; // Chọn phần tử đầu tiên
                    var selectedItem = comboBoxEdit1.SelectedItem as Item;

                    LoadData(selectedItem.Id);
                }
            }
            else
            {
                comboBoxEdit1.Properties.Items.Clear(); // Xóa dữ liệu cũ
                comboBoxEdit1.Properties.NullText = "Không có tài khoản nào";
            }
            //
            //Load data vat tu
            txtSohieu.Text = dtoVatTu.SoHieu;
            txtTenvattu.Text = dtoVatTu.Ten; 
            //Kiểm tra xem là sp moi hay cũ
            string queryCheckVatTu = @"SELECT * FROM KhachHang WHERE  SoHieu = ? ";
            var parameterss = new OleDbParameter[]
            {
                new OleDbParameter("?",dtoVatTu.SoHieu), 
               };
            var kq = ExecuteQuery(queryCheckVatTu, parameterss);
            if (kq.Rows.Count == 0)
            {
                txtMaSo.Text = "0";
            }
            else
            {
                txtMaSo.Text = kq.Rows[0]["MaSo"].ToString();
                txtTenvattu.Text = Helpers.ConvertVniToUnicode(kq.Rows[0]["Ten"].ToString());
                txtSohieu.Text = kq.Rows[0]["SoHieu"].ToString();
                txtDonvi.Text= kq.Rows[0]["MST"].ToString();
                txtGhichu.Text = Helpers.ConvertVniToUnicode(kq.Rows[0]["DiaChi"].ToString());
                int mapl = int.Parse(kq.Rows[0]["MaPhanLoai"].ToString());


                //comboBoxEdit1.SelectedItem=
                foreach (Item item in comboBoxEdit1.Properties.Items)
                {
                    if (item.Id == mapl)
                    {
                        comboBoxEdit1.EditValue = item; // Chọn mục theo ID
                        break; // Thoát khỏi vòng lặp
                    }
                }
            }

            DevExpress.XtraGrid.Views.Grid.GridView view = gridControl1.MainView as DevExpress.XtraGrid.Views.Grid.GridView;
            for (int i = 0; i < view.RowCount; i++)
            {
                // Lấy giá trị của cột STT
                if (view.GetRowCellValue(i, "SoHieu").ToString() == txtSohieu.Text)
                {
                    this.BeginInvoke((MethodInvoker)delegate
                    {
                        if (gridView1.RowCount > i) // Kiểm tra số lượng dòng
                        {
                            gridView1.FocusedRowHandle = i; // Đặt focus
                            gridView1.MakeRowVisible(i); // Cuộn đến dòng
                            gridView1.SelectRow(i); // Chọn dòng
                        }
                    });
                    return;
                }
            }
        }

        private void comboBoxEdit1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxEdit1.SelectedItem != null)
            {
                // Lấy phần tử được chọn
                var selectedItem = comboBoxEdit1.SelectedItem as Item;

                if (selectedItem != null)
                {
                    int selectedId = selectedItem.Id; // Lấy giá trị Id 
                    frmMain.currentselectId = comboBoxEdit1.SelectedIndex;
                    LoadData(selectedId);
                }
            }
        }
        private void LoadData(int Maso)
        {
            string query = @" SELECT *  FROM KhachHang where MaPhanLoai= ? ";
            var parameterss = new OleDbParameter[]
            {
                new OleDbParameter("?",Maso)
               };
            var kq = ExecuteQuery(query, parameterss);
            foreach (DataRow item in kq.Rows)
            {
                item["Ten"] = Helpers.ConvertVniToUnicode(item["Ten"].ToString());
                item["SoHieu"] = Helpers.ConvertVniToUnicode(item["SoHieu"].ToString());
            }
            gridControl1.DataSource = kq;

        }

        private void gridView1_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            int rowHandle = e.RowHandle;
            int ID = int.Parse(gridView1.GetRowCellValue(rowHandle, "MaSo").ToString());
            // Lấy dữ liệu từ hàng
            var value = gridView1.GetRowCellValue(rowHandle, "SoHieu").ToString();
            // tbKhachhang.AsEnumerable().Where(row => row.Field<string>("MST") == mst).FirstOrDefault();
            var getKhachhang =frmMain.tbKhachhang.AsEnumerable().Where(row => row.Field<Int32>("MaSo") == ID).FirstOrDefault();
            // Lấy chỉ số hàng đã click
            
            txtSohieu.Text = getKhachhang["SoHieu"].ToString();
            txtTenvattu.Text = Helpers.ConvertVniToUnicode(getKhachhang["Ten"].ToString());
            txtMaSo.Text = getKhachhang["MST"].ToString();
            txtGhichu.Text = Helpers.ConvertVniToUnicode(getKhachhang["DiaChi"].ToString());
            txtMaSo.Text = getKhachhang["MaSo"].ToString();
        }
        private void RefreshData()
        {
            txtMaSo.Text = "";
            txtSohieu.Text = "";
            txtTenvattu.Text = "";
            txtGhichu.Text = "";
            txtDonvi.Text = ""; 
        }
        private void btnGhi_Click(object sender, EventArgs e)
        {
            int selectedId = 0;
            var selectedItem = comboBoxEdit1.SelectedItem as Item;
            if (selectedItem.Id == 0)
            {
                XtraMessageBox.Show("Vui lòng chọn danh mục");
                return;
            }
            if (selectedItem != null)
            {
                selectedId = selectedItem.Id; // Lấy giá trị Id  
            }

            // Xác định xem đây là thêm mới hay cập nhật
            bool isInsert = txtMaSo.Text == "";
            string query;
            OleDbParameter[] parameters;

            if (isInsert)
            {
                query = @"INSERT INTO KhachHang (MaPhanLoai, SoHieu, Ten, DiaChi, MST) VALUES (?, ?, ?, ?, ?)";
                parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", selectedId),
            new OleDbParameter("?", txtSohieu.Text),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(txtTenvattu.Text)),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(txtGhichu.Text)),
            new OleDbParameter("?", txtDonvi.Text)
                };
            }
            else
            {
                query = @"UPDATE KhachHang SET MaPhanLoai=?, SoHieu=?, Ten=?, DiaChi=?, MST=? WHERE MaSo=?";
                parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", selectedId),
            new OleDbParameter("?", txtSohieu.Text),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(txtTenvattu.Text)),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(txtGhichu.Text)),
            new OleDbParameter("?", txtDonvi.Text),
            new OleDbParameter("?", txtMaSo.Text)
                };
            }

            // Thực hiện truy vấn
            int rowsAffected = ExecuteQueryResult(query, parameters);
            query = "SELECT * FROM KhachHang"; // Giả sử bạn muốn lấy tất cả dữ liệu từ bảng KhachHang
            frmMain.tbKhachhang = ExecuteQuery(query);
            // [Optional] Xử lý kết quả trả về (ví dụ: thông báo thành công/thất bại)
            if (rowsAffected > 0)
            {

                //var hiddenValue = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["SoHieu"]);
                //var hiddenValue2 = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["DonVi"]);
                //var hiddenValue3 = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["TenVattu"]);
                frmMain.hiddenValue = txtSohieu.Text;
                frmMain.hiddenValue2 = txtDonvi.Text;
                frmMain.hiddenValue3 = txtTenvattu.Text;
                //this.Close();

                LoadData(selectedItem.Id);
              //  RefreshData();
                DevExpress.XtraGrid.Views.Grid.GridView view = gridControl1.MainView as DevExpress.XtraGrid.Views.Grid.GridView; // Lấy GridView
                for (int i = 0; i < view.RowCount; i++)
                {
                    // Lấy giá trị của cột STT
                    if (view.GetRowCellValue(i, "SoHieu").ToString() == txtSohieu.Text)
                    {
                        view.FocusedRowHandle = i; // Chọn dòng
                        view.SelectRow(i); // Chọn dòng
                        view.MakeRowVisible(i); // Cuộn tới dòng đã chọn
                        return; // Thoát sau khi tìm thấy
                    }
                }
            }
            else
            {
                MessageBox.Show(isInsert ? "Thêm mới thất bại!" : "Cập nhật thất bại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            RefreshData();
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void gridControl1_DoubleClick(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = gridControl1.MainView as DevExpress.XtraGrid.Views.Grid.GridView;
            var hitInfo = gridView.CalcHitInfo(gridView.GridControl.PointToClient(MousePosition));


            // Kiểm tra nếu nhấp vào một ô
            if (hitInfo.InRowCell)
            {
                int columnIndex = hitInfo.Column.VisibleIndex; // Chỉ số cột

                // Lấy giá trị trong ô đã nhấp
                var hiddenValue = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["Ten"]);
                var hiddenValue2 = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["SoHieu"]);
                var hiddenValue3 = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["MST"]);
                frmMain.hiddenValue = hiddenValue.ToString();
                frmMain.hiddenValue2 = hiddenValue2.ToString();
                frmMain.hiddenValue3 = hiddenValue3.ToString(); 

                this.Close();
            }
        }
    }
}