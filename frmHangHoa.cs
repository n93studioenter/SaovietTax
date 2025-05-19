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

namespace SaovietTax
{
	public partial class frmHangHoa: DevExpress.XtraEditors.XtraForm
	{
        public class VatTu
        {
            public int MaSo { get; set; }
            public int MaPhanLoai { get; set; }
            public string SoHieu { get; set; }
            public string TenVattu { get; set; }
            public string DonVi { get; set; }
        }
        public VatTu dtoVatTu { get; set; }
        public frmHangHoa()
		{
            InitializeComponent();
            dtoVatTu = new VatTu();

        }
        private void LoadData(int Maso)
        {
            string query = @" SELECT *  FROM Vattu where MaPhanLoai= ? ";
            var parameterss = new OleDbParameter[]
            {
                new OleDbParameter("?",Maso)
               };
            var kq = ExecuteQuery(query, parameterss);
            foreach(DataRow item in kq.Rows)
            {
                item["TenVattu"] = Helpers.ConvertVniToUnicode(item["TenVattu"].ToString());
                item["DonVi"] = Helpers.ConvertVniToUnicode(item["DonVi"].ToString());
            }
            gridControl1.DataSource = kq;
           
        }
        private void frmHangHoa_Load(object sender, EventArgs e)
        {
                string query = @"SELECT * FROM PhanLoaiVattu ORDER BY TenPhanLoai"; 
            var dt = ExecuteQuery(query, null);
            if (dt != null && dt.Rows.Count > 0)
            {
                comboBoxEdit1.Properties.Items.Clear(); // Xóa các mục cũ

                foreach (DataRow row in dt.Rows)
                {
                    // Thêm từng mục vào ComboBoxEdit
                    comboBoxEdit1.Properties.Items.Add(new Item
                    {
                        Name = row["SoHieu"].ToString() +" - "+ Helpers.ConvertVniToUnicode(row["TenPhanLoai"].ToString()),
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
            txtTenvattu.Text = dtoVatTu.TenVattu;
            txtDonvi.Text = dtoVatTu.DonVi;
            //Kiểm tra xem là sp moi hay cũ
            string queryCheckVatTu = @"SELECT * FROM Vattu WHERE LCase(SoHieu) = LCase(?) AND LCase(DonVi) = LCase(?)";
          var  parameterss = new OleDbParameter[]
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
           if (comboBoxEdit1.SelectedItem != null)
            {
                // Lấy phần tử được chọn
                var selectedItem = comboBoxEdit1.SelectedItem as Item;

                if (selectedItem != null)
                {
                    int selectedId = selectedItem.Id; // Lấy giá trị Id 
                    LoadData(selectedId);
                }
            }
        }
        public frmMain frmMain;
        public bool isChange = false;
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
                var hiddenValue2= gridView.GetRowCellValue(hitInfo.RowHandle, gridView.Columns["DonVi"]);
                frmMain.hiddenValue = hiddenValue.ToString();
                frmMain.hiddenValue2= hiddenValue2.ToString();
                isChange = true;
                this.Close();
            }
        }

        private void btnGhi_Click(object sender, EventArgs e)
        {
            int selectedId = 0;
            var selectedItem = comboBoxEdit1.SelectedItem as Item;

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
                LoadData(selectedItem.Id);
                RefreshData();
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
        }
        private void btnThem_Click(object sender, EventArgs e)
        {
            RefreshData();
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
    }
}