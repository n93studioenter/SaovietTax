using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SaovietTax.DTO
{
    public class VatTu
    {
        public int MaSo { get; set; }
        public int MaPhanLoai { get; set; }
        public string SoHieu { get; set; }
        public string GhiChu { get; set; }
        public string TenVattu { get; set; }
        public string DonVi { get; set; }
        public double Dongia { get; set; }
        public double SoLuong { get; set; }
        public double ThanhTien { get; set; }
    }
}
