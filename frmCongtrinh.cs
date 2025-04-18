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
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.OleDb;
using System.Configuration;
using DevExpress.XtraGrid.Views.Grid;

namespace SaovietTax
{
    public partial class frmCongtrinh : DevExpress.XtraEditors.XtraForm
    {
        public frmCongtrinh()
        {
            InitializeComponent();
        }

        private void frmCongtrinh_Load(object sender, EventArgs e)
        {
            string query = @" SELECT *  FROM TP154 ";
            var kq = ExecuteQuery(query, null);
            gridControl1.DataSource = kq;
        }
        private DataTable ExecuteQuery(string query, params OleDbParameter[] parameters)
        {
            DataTable dataTable = new DataTable();
            string dbPath = ConfigurationManager.AppSettings["dbpath"];
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

        }
    }
}