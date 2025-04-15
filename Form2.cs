using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaovietTax
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        string dbPath = "";
        public static int Id = 1;
        bool isSetuppath = false;
        bool isetupDbpath = false;
        string password, connectionString;

        private void button1_Click(object sender, EventArgs e)
        {
            frmMain frmm = new frmMain();
            frmm.Show();

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string dbPath = "D:\\Tao moi1\\Data\\Phat Dat Vung Tau 25.mdb";
            //dbPath = "sadsa";


            // Đọc toàn bộ nội dung tệp
            string password = "1@35^7*9)1";
            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            //connectionString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            // connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database";
            //connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\S.T.E 25\S.T.E 25\DATA\importData.accdb;Persist Security Info=False";
            MessageBox.Show(connectionString);
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                MessageBox.Show("Kết nối database thành công");
            }
        }
    }
}
