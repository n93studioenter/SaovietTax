using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static SaovietTax.frmMain;

namespace SaovietTax
{
    public partial class SuggestionControl : UserControl
    {
        public event EventHandler<string> ItemSelected;

        public List<HeThongTK> Suggestions
        {
            set
            {
                listBox1.Items.Clear();
                foreach (var item in value)
                {
                    listBox1.Items.Add(item.SoHieu +"-"+Helpers.ConvertVniToUnicode(item.Ten)); // Hiển thị tên
                }
            }
        }

        public SuggestionControl()
        {
            InitializeComponent();
            listBox1.DoubleClick += ListBox1_DoubleClick;
        }

        private void ListBox1_DoubleClick(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                // Lấy SoHieu từ mục đã chọn
                string selectedItem = listBox1.SelectedItem.ToString();
                string soHieu = selectedItem.Split('-')[0].Trim(); // Tách SoHieu ra

                ItemSelected?.Invoke(this, soHieu); // Gửi SoHieu
                Hide();
            }
        }

        public void UpdateSuggestions(List<HeThongTK> newSuggestions)
        { 
            Suggestions = newSuggestions;

            // Hiển thị UserControl nếu có gợi ý
            if (listBox1.Items.Count > 0)
            {
                this.Show();
            }
            else
            {
                this.Hide();
            }
        }
    }
}
