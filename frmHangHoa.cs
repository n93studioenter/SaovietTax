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
using System.Text.RegularExpressions;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors.Mask.Design;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using DevExpress.Xpo.DB.Helpers;
using System.Web.UI.WebControls;
using Windows.UI.Xaml.Controls;
using DevExpress.Utils;
using SaovietTax.DTO;

namespace SaovietTax
{
	public partial class frmHangHoa: DevExpress.XtraEditors.XtraForm
	{
      
        public VatTu dtoVatTu { get; set; }
        public frmHangHoa()
		{
            InitializeComponent();
            dtoVatTu = new VatTu();
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
                 

            }
        }
        private void LoadData(int Maso)
        {
            string query = "";
            if (Maso != 0)
            {  
                gridControl1.DataSource = frmMain.lstvt.Where(m=>m.MaPhanLoai==Maso);
                GridStripRow(gridView1);
            }
            else
            {
                //query = @" SELECT *  FROM Vattu  "; 
                //var kq = ExecuteQuery(query, null);
                //foreach (DataRow item in kq.Rows)
                //{
                //    item["TenVattu"] = Helpers.ConvertVniToUnicode(item["TenVattu"].ToString());
                //    item["DonVi"] = Helpers.ConvertVniToUnicode(item["DonVi"].ToString());
                //}
                gridControl1.DataSource = frmMain.lstvt;

                GridStripRow(gridView1);
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // Kiểm tra phím tắt (ví dụ: Ctrl + N)
            if (keyData == (Keys.Control | Keys.G))
            {
                btnGhi.PerformClick(); // Gọi sự kiện nhấn nút
                return true; // Đã xử lý phím
            }
            return base.ProcessCmdKey(ref msg, keyData); // Chuyển tiếp cho xử lý tiếp
        }
        private bool firstload=true; 
        private void frmHangHoa_Load(object sender, EventArgs e)
        {
             
           
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

        private void comboBoxEdit1_SelectedIndexChanged(object sender, EventArgs e)
        {
           if (comboBoxEdit1.SelectedItem != null && !firstload)
            {
                // Lấy phần tử được chọn
                var selectedItem = comboBoxEdit1.SelectedItem as Item;

                if (selectedItem != null)
                {
                    int selectedId = selectedItem.Id; // Lấy giá trị Id 
                    //frmMain.currentselectId = comboBoxEdit1.SelectedIndex;
                    LoadData(selectedId);
                }
            }
        }
        public frmMain frmMain;
        public bool isChange = false;
        private void gridControl1_DoubleClick(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = gridControl1.MainView as DevExpress.XtraGrid.Views.Grid.GridView;
            var hitInfo = gridView.CalcHitInfo(gridView.GridControl.PointToClient(MousePosition));


            // Kiểm tra nếu nhấp vào một ô
            if (hitInfo.InRowCell)
            {
                int columnIndex = hitInfo.Column.VisibleIndex; // Chỉ số cột

                // Lấy giá trị trong ô đã nhấp
                var hiddenValue = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["SoHieu"]);
                var hiddenValue2= gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["DonVi"]);
                var hiddenValue3 = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["TenVattu"]);
                frmMain.hiddenValue = hiddenValue.ToString();
                frmMain.hiddenValue2= hiddenValue2.ToString();
                frmMain.hiddenValue3 = hiddenValue3.ToString();
                isChange = true;
                //

                this.Close();
            }
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
            bool isInsert = txtMaSo.Text == "0";
            string query;
            OleDbParameter[] parameters;

            if (isInsert)
            {
                query = @"INSERT INTO Vattu (MaPhanLoai, SoHieu, TenVattu, DonVi, GhiChu) VALUES (?, ?, ?, ?, ?)";
                parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", selectedId),
            new OleDbParameter("?", txtSohieu.Text),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(txtTenvattu.Text)),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(txtDonvi.Text)),
            new OleDbParameter("?", string.IsNullOrEmpty(txtGhichu.Text)?"...":txtGhichu.Text)
                };
            }
            else
            {
                query = @"UPDATE Vattu SET MaPhanLoai=?, SoHieu=?, TenVattu=?, DonVi=?, GhiChu=? WHERE MaSo=?";
                parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", selectedId),
            new OleDbParameter("?", txtSohieu.Text),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(txtTenvattu.Text)),
            new OleDbParameter("?", Helpers.ConvertUnicodeToVni(txtDonvi.Text)),
            new OleDbParameter("?", txtGhichu.Text),
            new OleDbParameter("?", txtMaSo.Text)
                };
            }

            // Thực hiện truy vấn
            int rowsAffected = ExecuteQueryResult(query, parameters);

            // [Optional] Xử lý kết quả trả về (ví dụ: thông báo thành công/thất bại)
            if (rowsAffected > 0)
            {

                //var hiddenValue = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["SoHieu"]);
                //var hiddenValue2 = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["DonVi"]);
                //var hiddenValue3 = gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["TenVattu"]);
                frmMain.hiddenValue = txtSohieu.Text;
                frmMain.hiddenValue2 = txtDonvi.Text;
                frmMain.hiddenValue3 = txtTenvattu.Text;
                isChange = true;
                this.Close();

                LoadData(selectedItem.Id);
                // RefreshData();
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


        private void btnThoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void RefreshData()
        {
            txtMaSo.Text = "0";
            txtSohieu.Text = "";
            txtTenvattu.Text = "";
            txtDonvi.Text = "";
            txtGhichu.Text = "";
            gridControl2.DataSource = null;
        }
        private void btnThem_Click(object sender, EventArgs e)
        {
            RefreshData();
        }
        private class TonKho
        {
            public int MaSo { get; set; }
            public double SoLuong { get; set; }
            public double DonGia { get; set; }
            public double ThanhTien { get; set; }
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
            txtMaSo.Text = gridView1.GetRowCellValue(rowHandle, "MaSo").ToString();

            string query = @" SELECT *  FROM TonKho where MaVatTu= ? ";
            var parameterss = new OleDbParameter[]
            {
                new OleDbParameter("?",txtMaSo.Text)
               };
            var kq = ExecuteQuery(query, parameterss);
            List<TonKho> lstTonkho = new List<TonKho>();
            if (kq.Rows.Count > 0)
            {
                try
                {
                    TonKho tk = new TonKho();
                    int cnt = 12;
                    while (kq.Rows[0]["Luong_" + cnt].ToString() == "0")
                    {
                        cnt += 1;
                    }
                    tk.SoLuong = kq.Rows[0]["Luong_" + cnt] != null ? double.Parse(kq.Rows[0]["Luong_" + cnt].ToString()) : 0;
                    tk.ThanhTien = kq.Rows[0]["Tien_" + cnt] != null ? double.Parse(kq.Rows[0]["Tien_" + cnt].ToString()) : 0;
                    if (tk.SoLuong!=0 && tk.ThanhTien!=0)
                    {
                        tk.DonGia = Math.Round(double.Parse(kq.Rows[0]["Tien_" + cnt].ToString()) / tk.SoLuong, 2);
                        lstTonkho.Add(tk);
                    }
                }
                catch(Exception ex)
                {

                }
              
            }
            gridControl2.DataSource = lstTonkho;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
           
            DialogResult result = XtraMessageBox.Show(
        "Bạn có chắc chắn muốn xóa vật tư này?",
        "Xác Nhận",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                var query = @"delete from Vattu where MaSo=?";
                var parameters = new OleDbParameter[]
               {
            new OleDbParameter("?", txtMaSo.Text)
               };
                int rowsAffected = ExecuteQueryResult(query, parameters);
                var selectedItem = comboBoxEdit1.SelectedItem as Item;

                if (selectedItem != null)
                {
                    int selectedId = selectedItem.Id; // Lấy giá trị Id 
                    LoadData(selectedId);
                }
            }
            
        }

        private void gridView1_DataSourceChanged(object sender, EventArgs e)
        {
           
        }

        private void frmHangHoa_Load_1(object sender, EventArgs e)
        {
            gridView1.OptionsFind.AlwaysVisible = true; // Kích hoạt thanh tìm kiếm

            var query = @"SELECT * FROM PhanLoaiVattu ORDER BY TenPhanLoai";
            var dt = ExecuteQuery(query, null);
            if (dt != null && dt.Rows.Count > 0)
            {
                comboBoxEdit1.Properties.Items.Clear(); // Xóa các mục cũ
                comboBoxEdit1.Properties.Items.Add(new Item { Name = "Tất cả", Id = 0 });
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
                int idsl = 0;
                if (comboBoxEdit1.Properties.Items.Count > 0)
                {
                    foreach (Item item in comboBoxEdit1.Properties.Items)
                    {
                        if (item.Id == frmMain.currentselectId)
                        {
                            idsl = comboBoxEdit1.Properties.Items.IndexOf(item);
                            break;
                        }
                    }
                    comboBoxEdit1.SelectedIndex = idsl; // Chọn phần tử đầu tiên
                    var selectedItem = comboBoxEdit1.SelectedItem as Item;
                    firstload = false;
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
            txtTenvattu.Text = dtoVatTu.TenVattu;
            txtDonvi.Text = dtoVatTu.DonVi;
            //Kiểm tra xem là sp moi hay cũ
            string queryCheckVatTu = @"SELECT * FROM Vattu WHERE LCase(SoHieu) = LCase(?) AND LCase(DonVi) = LCase(?)";
            var parameterss = new OleDbParameter[]
            {
                new OleDbParameter("?",dtoVatTu.SoHieu.ToLower()),
                 new OleDbParameter("?",Helpers.ConvertUnicodeToVni(dtoVatTu.DonVi.ToLower()))
               };
            var kq = ExecuteQuery(queryCheckVatTu, parameterss);
            if (kq.Rows.Count == 0)
            {
                txtMaSo.Text = "0";
            }
            else
            {
                txtMaSo.Text = kq.Rows[0]["MaSo"].ToString();
                txtGhichu.Text = kq.Rows[0]["GhiChu"].ToString();
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
    }
}