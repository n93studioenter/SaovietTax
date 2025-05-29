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
using AForge.Video.DirectShow;
using AForge.Video;

namespace SaovietTax
{
	public partial class frmCamera: DevExpress.XtraEditors.XtraForm
	{
        public frmCamera()
		{
            InitializeComponent();
		}
        private FilterInfoCollection videoDevices; // Danh sách camera
        private VideoCaptureDevice videoSource; // Camera hiện tại
        private void frmCamera_Load(object sender, EventArgs e)
        {
            // Chọn camera đầu tiên
            // Lấy danh sách camera
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            if (videoDevices.Count == 0)
            {
                MessageBox.Show("Không tìm thấy camera.");
                return;
            }

            // Chọn camera đầu tiên
            videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);
            videoSource.NewFrame += new NewFrameEventHandler(video_NewFrame);
            videoSource.Start(); // Bắt đầu camera
        }
        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            // Hiển thị hình ảnh từ camera vào PictureBox
            pictureBox1.Image = (Bitmap)eventArgs.Frame.Clone();
        }
    }
}