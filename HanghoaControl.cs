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
    public partial class HanghoaControl: UserControl
    {
        public event EventHandler<string> ItemSelected;

        public List<HangHoaTK> Suggestions
        {
            set
            {
                listBox1.Items.Clear();
                foreach (var item in value)
                {
                    listBox1.Items.Add(item.SoHieu + "|" + Helpers.ConvertVniToUnicode(item.TenVattu)); // Hiển thị tên
                }
            }
        }

        public HanghoaControl()
        {
            InitializeComponent();
            listBox1.DrawMode = DrawMode.OwnerDrawFixed; // Hoặc OwnerDrawVariable

            listBox1.DoubleClick += ListBox1_DoubleClick;

        }

        private void ListBox1_DoubleClick(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                // Lấy SoHieu từ mục đã chọn
                string selectedItem = listBox1.SelectedItem.ToString();
                string soHieu = selectedItem.Split('|')[0].Trim(); // Tách SoHieu ra
                string tenVatTu = selectedItem.Split('|')[1].Trim(); // Tách tên vật tư ra
                ItemSelected?.Invoke(this, soHieu); // Gửi SoHieu
                Hide();
            }
        }

        public void UpdateSuggestions(List<HangHoaTK> newSuggestions)
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

        private void listBox1_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            string itemText = listBox1.Items[e.Index].ToString();
            string[] parts = itemText.Split('|');
            e.DrawBackground();

            // Vẽ phần đầu tiên với màu xanh
            e.Graphics.DrawString(parts[0], new Font(e.Font, FontStyle.Regular), Brushes.Blue, e.Bounds);

            // Vẽ phần thứ hai với màu đỏ
            if (parts.Length > 1)
            {
                float textWidth = e.Graphics.MeasureString(parts[0], e.Font).Width;
                e.Graphics.DrawString(parts[1], e.Font, Brushes.Black, e.Bounds.X + textWidth, e.Bounds.Y);
            }

            // Kết thúc vẽ
            e.DrawFocusRectangle();
        }
    }
}
