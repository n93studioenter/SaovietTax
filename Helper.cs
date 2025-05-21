using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ClosedXML.Excel.XLPredefinedFormat;

namespace SaovietTax
{
    public static class Helpers
    {
      
        private static readonly Dictionary<char, string> UnicodeToVniMap = new Dictionary<char, string>
{
   
    {'À', "AØ"},{'\'', ""}, {'Á', "AÙ"}, {'Â', "AÂ"}, {'Ã', "AÕ"},
            {'È', "EØ"}, {'É', "EÙ"}, {'Ê', "EÂ"}, {'Ì', "Ì"},
            {'Í', "Í"}, {'Ò', "OØ"}, {'Ó', "OÙ"}, {'Ô', "OÂ"},
            {'Õ', "OÕ"}, {'Ù', "UØ"}, {'Ú', "UÙ"}, {'Ý', "YÙ"},
            {'à', "aø"}, {'á', "aù"}, {'â', "aâ"}, {'ã', "aõ"},
            {'è', "eø"}, {'é', "eù"}, {'ê', "eâ"}, {'ì', "ì"},
            {'í', "í"}, {'ò', "oø"}, {'ó', "où"}, {'ô', "oâ"},
            {'õ', "oõ"}, {'ù', "uø"}, {'ú', "uù"}, {'ý', "yù"},
            {'Ă', "AÊ"}, {'ă', "aê"}, {'Đ', "Ñ"}, {'đ', "ñ"},
            {'Ĩ', "Ó"}, {'ĩ', "ó"}, {'Ũ', "UÕ"}, {'ũ', "uõ"},
            {'Ơ', "Ô"}, {'ơ', "ô"}, {'Ư', "Ö"}, {'ư', "ö"},
            {'Ạ', "AÏ"}, {'ạ', "aï"}, {'Ả', "AÛ"}, {'ả', "aû"},
            {'Ấ', "AÁ"}, {'ấ', "aá"}, {'Ầ', "AÀ"}, {'ầ', "aà"},
            {'Ẩ', "AÅ"}, {'ẩ', "aå"}, {'Ẫ', "AÃ"}, {'ẫ', "aã"},
            {'Ậ', "AÄ"}, {'ậ', "aä"}, {'Ắ', "AÉ"}, {'ắ', "aé"},
            {'Ằ', "AÈ"}, {'ằ', "aè"}, {'Ẳ', "AÚ"}, {'ẳ', "aú"},
            {'Ẵ', "AÜ"}, {'ẵ', "aü"}, {'Ặ', "AË"}, {'ặ', "aë"},
            {'Ẹ', "EÏ"}, {'ẹ', "eï"}, {'Ẻ', "EÛ"}, {'ẻ', "eû"},
            {'Ẽ', "EÕ"}, {'ẽ', "eõ"}, {'Ế', "EÁ"}, {'ế', "eá"},
            {'Ề', "EÀ"}, {'ề', "eà"}, {'Ể', "EÅ"}, {'ể', "eå"},
            {'Ễ', "EÃ"}, {'ễ', "eã"}, {'Ệ', "EÄ"}, {'ệ', "eä"},
            {'Ỉ', "Æ"}, {'ỉ', "æ"}, {'Ị', "Ò"}, {'ị', "ò"},
            {'Ọ', "OÏ"}, {'ọ', "oï"}, {'Ỏ', "OÛ"}, {'ỏ', "oû"},
            {'Ố', "OÁ"}, {'ố', "oá"}, {'Ồ', "OÀ"}, {'ồ', "oà"},
            {'Ổ', "OÅ"}, {'ổ', "oå"}, {'Ỗ', "OÃ"}, {'ỗ', "oã"},
            {'Ộ', "OÄ"}, {'ộ', "oä"}, {'Ớ', "ÔÙ"}, {'ớ', "ôù"},
            {'Ờ', "ÔØ"}, {'ờ', "ôø"}, {'Ở', "ÔÛ"}, {'ở', "ôû"},
            {'Ỡ', "ÔÕ"}, {'ỡ', "ôõ"}, {'Ợ', "ÔÏ"}, {'ợ', "ôï"},
            {'Ụ', "UÏ"}, {'ụ', "uï"}, {'Ủ', "UÛ"}, {'ủ', "uû"},
            {'Ứ', "ÖÙ"}, {'ứ', "öù"}, {'Ừ', "ÖØ"}, {'ừ', "öø"},
            {'Ử', "ÖÛ"}, {'ử', "öû"}, {'Ữ', "ÖÕ"}, {'ữ', "öõ"},
            {'Ự', "ÖÏ"}, {'ự', "öï"}, {'Ỳ', "YØ"}, {'ỳ', "yø"},
            {'Ỵ', "Î"}, {'ỵ', "î"}, {'Ỷ', "YÛ"}, {'ỷ', "yû"},
            {'Ỹ', "YÕ"}, {'ỹ', "yõ"}
};

        public static string InsertZero(string text)
        {
            int lenght = int.Parse(text);
            return lenght.ToString("D8");
        }
        static string RemoveLeadingSpecialCharacters(string input)
        {
            // Sử dụng LINQ để lấy các ký tự không phải là ký tự đặc biệt
            return new string(input.SkipWhile(c => !char.IsLetterOrDigit(c)).ToArray());
        }
        public static string NormalizeVietnameseString(string input)
        {
            //Bỏ đi ký tự đặc biệt đầu chữ
            input = RemoveLeadingSpecialCharacters(input);
            //Bỏ đi tab
            input = input.Replace("\t", ""); // Thay thế ký tự tab bằng chuỗi rỗng
            input = input.Normalize(NormalizationForm.FormC);

            if (string.IsNullOrEmpty(input))
                return input;
            return input;
        }
        public static string ConvertUnicodeToVni(string input)
        {
            input = NormalizeVietnameseString(input);
            StringBuilder output = new StringBuilder();

            foreach (char c in input)
            {
                string character = c.ToString();
                Console.WriteLine($"Current character: {character}");

                if (UnicodeToVniMap.ContainsKey(c))
                {
                    output.Append(UnicodeToVniMap[c]); // Thay thế bằng ký tự VNI tương ứng
                }
                else
                {
                    output.Append(c); // Giữ nguyên ký tự nếu không có trong bảng ánh xạ
                }
            }

            return output.ToString();
        }

        private static readonly Dictionary<string, char> VniToUnicodeMap = new Dictionary<string, char>
    {
        // Chữ thường
        //{"aø", 'à'}, {"aù", 'á'}, {"aû", 'ả'}, {"aõ", 'ã'}, {"aï", 'ạ'},
        //{"aâ", 'â'}, {"aá", 'ấ'}, {"aà", 'ầ'}, {"aå", 'ẩ'}, {"aã", 'ẫ'}, {"aä", 'ậ'},
        //{"eø", 'è'}, {"eù", 'é'}, {"eû", 'ẻ'}, {"eõ", 'ẽ'}, {"eï", 'ẹ'},
        //{"eâ", 'ê'}, {"eá", 'ế'}, {"eà", 'ề'}, {"eå", 'ể'}, {"eã", 'ễ'}, {"eä", 'ệ'},
        //{"ì", 'ì'}, {"í", 'í'}, {"æ", 'ỉ'}, {"ó", 'ĩ'}, {"ò", 'ị'},
        //{"oø", 'ò'}, {"où", 'ó'}, {"oû", 'ỏ'}, {"oõ", 'õ'}, {"oï", 'ọ'},
        //{"oâ", 'ô'}, {"oá", 'ố'}, {"oà", 'ồ'}, {"oå", 'ổ'}, {"oã", 'ỗ'}, {"oä", 'ộ'},
        //{"ô", 'ơ'}, {"ôù", 'ớ'}, {"ôø", 'ờ'}, {"ôû", 'ở'}, {"ôõ", 'ỡ'}, {"ôï", 'ợ'},
        //{"uø", 'ù'}, {"uù", 'ú'}, {"uû", 'ủ'}, {"uõ", 'ũ'}, {"uï", 'ụ'},
        //{"ö", 'ư'}, {"öù", 'ứ'}, {"öø", 'ừ'}, {"öû", 'ử'}, {"öõ", 'ữ'}, {"öï", 'ự'},
        //{"yø", 'ỳ'}, {"yù", 'ý'}, {"yû", 'ỷ'}, {"yõ", 'ỹ'}, {"yï", 'ỵ'}, {"ñ", 'đ'},

        //// Chữ hoa
        //{"AØ", 'À'}, {"AÙ", 'Á'}, {"AÛ", 'Ả'}, {"AÕ", 'Ã'}, {"AÏ", 'Ạ'},
        //{"AÂ", 'Â'}, {"AÁ", 'Ấ'}, {"AÀ", 'Ầ'}, {"AÅ", 'Ẩ'}, {"AÃ", 'Ẫ'}, {"AÄ", 'Ậ'},
        //{"EØ", 'È'}, {"EÙ", 'É'}, {"EÛ", 'Ẻ'}, {"EÕ", 'Ẽ'}, {"EÏ", 'Ẹ'},
        //{"EÂ", 'Ê'}, {"EÁ", 'Ế'}, {"EÀ", 'Ề'}, {"EÅ", 'Ể'}, {"EÃ", 'Ễ'}, {"EÄ", 'Ệ'},
        //{"Ì", 'Ì'}, {"Í", 'Í'}, {"Æ", 'Ỉ'}, {"Ó", 'Ĩ'}, {"Ò", 'Ị'},
        //{"OØ", 'Ò'}, {"OÙ", 'Ó'}, {"OÛ", 'Ỏ'}, {"OÕ", 'Õ'}, {"OÏ", 'Ọ'},
        //{"OÂ", 'Ô'}, {"OÁ", 'Ố'}, {"OÀ", 'Ồ'}, {"OÅ", 'Ổ'}, {"OÃ", 'Ỗ'}, {"OÄ", 'Ộ'},
        //{"Ô", 'Ơ'}, {"ÔÙ", 'Ớ'}, {"ÔØ", 'Ờ'}, {"ÔÛ", 'Ở'}, {"ÔÕ", 'Ỡ'}, {"ÔÏ", 'Ợ'},
        //{"UØ", 'Ù'}, {"UÙ", 'Ú'}, {"UÛ", 'Ủ'}, {"UÕ", 'Ũ'}, {"UÏ", 'Ụ'},
        //{"Ö", 'Ư'}, {"ÖÙ", 'Ứ'}, {"ÖØ", 'Ừ'}, {"ÖÛ", 'Ử'}, {"ÖÕ", 'Ữ'}, {"ÖÏ", 'Ự'},
        //{"YØ", 'Ỳ'}, {"YÙ", 'Ý'}, {"YÛ", 'Ỷ'}, {"YÕ", 'Ỹ'}, {"YÏ", 'Ỵ'}, {"Ñ", 'Đ'}
        // Chữ thường
    {"aø", 'à'}, {"aù", 'á'}, {"aû", 'ả'}, {"aõ", 'ã'}, {"aï", 'ạ'},
    {"aâ", 'â'}, {"aá", 'ấ'}, {"aà", 'ầ'}, {"aå", 'ẩ'}, {"aã", 'ẫ'}, {"aä", 'ậ'},
    {"aé", 'ắ'}, {"aè", 'ằ'}, {"aú", 'ẳ'}, {"aü", 'ẵ'}, {"aë", 'ặ'},
    {"eø", 'è'}, {"eù", 'é'}, {"eû", 'ẻ'}, {"eõ", 'ẽ'}, {"eï", 'ẹ'},
    {"eâ", 'ê'}, {"eá", 'ế'}, {"eà", 'ề'}, {"eå", 'ể'}, {"eã", 'ễ'}, {"eä", 'ệ'},
    {"ì", 'ì'}, {"í", 'í'}, {"æ", 'ỉ'}, {"ó", 'ĩ'}, {"ò", 'ị'},
    {"oø", 'ò'}, {"où", 'ó'}, {"oû", 'ỏ'}, {"oõ", 'õ'}, {"oï", 'ọ'},
    {"oâ", 'ô'}, {"oá", 'ố'}, {"oà", 'ồ'}, {"oå", 'ổ'}, {"oã", 'ỗ'}, {"oä", 'ộ'},
    {"ô", 'ơ'}, {"ôù", 'ớ'}, {"ôø", 'ờ'}, {"ôû", 'ở'}, {"ôõ", 'ỡ'}, {"ôï", 'ợ'},
    {"uø", 'ù'}, {"uù", 'ú'}, {"uû", 'ủ'}, {"uõ", 'ũ'}, {"uï", 'ụ'},
    {"ö", 'ư'}, {"öù", 'ứ'}, {"öø", 'ừ'}, {"öû", 'ử'}, {"öõ", 'ữ'}, {"öï", 'ự'},
    {"yø", 'ỳ'}, {"yù", 'ý'}, {"yû", 'ỷ'}, {"yõ", 'ỹ'}, {"yï", 'ỵ'}, {"ñ", 'đ'},

    // Chữ hoa
    {"AØ", 'À'}, {"AÙ", 'Á'}, {"AÛ", 'Ả'}, {"AÕ", 'Ã'}, {"AÏ", 'Ạ'},
    {"AÂ", 'Â'}, {"AÁ", 'Ấ'}, {"AÀ", 'Ầ'}, {"AÅ", 'Ẩ'}, {"AÃ", 'Ẫ'}, {"AÄ", 'Ậ'},
    {"AÉ", 'Ắ'}, {"AÈ", 'Ằ'}, {"AÚ", 'Ẳ'}, {"AÜ", 'Ẵ'}, {"AË", 'Ặ'},
    {"EØ", 'È'}, {"EÙ", 'É'}, {"EÛ", 'Ẻ'}, {"EÕ", 'Ẽ'}, {"EÏ", 'Ẹ'},
    {"EÂ", 'Ê'}, {"EÁ", 'Ế'}, {"EÀ", 'Ề'}, {"EÅ", 'Ể'}, {"EÃ", 'Ễ'}, {"EÄ", 'Ệ'},
    {"Ì", 'Ì'}, {"Í", 'Í'}, {"Æ", 'Ỉ'}, {"Ó", 'Ĩ'}, {"Ò", 'Ị'},
    {"OØ", 'Ò'}, {"OÙ", 'Ó'}, {"OÛ", 'Ỏ'}, {"OÕ", 'Õ'}, {"OÏ", 'Ọ'},
    {"OÂ", 'Ô'}, {"OÁ", 'Ố'}, {"OÀ", 'Ồ'}, {"OÅ", 'Ổ'}, {"OÃ", 'Ỗ'}, {"OÄ", 'Ộ'},
    {"Ô", 'Ơ'}, {"ÔÙ", 'Ớ'}, {"ÔØ", 'Ờ'}, {"ÔÛ", 'Ở'}, {"ÔÕ", 'Ỡ'}, {"ÔÏ", 'Ợ'},
    {"UØ", 'Ù'}, {"UÙ", 'Ú'}, {"UÛ", 'Ủ'}, {"UÕ", 'Ũ'}, {"UÏ", 'Ụ'},
    {"Ö", 'Ư'}, {"ÖÙ", 'Ứ'}, {"ÖØ", 'Ừ'}, {"ÖÛ", 'Ử'}, {"ÖÕ", 'Ữ'}, {"ÖÏ", 'Ự'},
    {"YØ", 'Ỳ'}, {"YÙ", 'Ý'}, {"YÛ", 'Ỷ'}, {"YÕ", 'Ỹ'}, {"YÏ", 'Ỵ'}, {"Ñ", 'Đ'}
    };

        public static string ConvertVniToUnicode(string input)
        {
            input = NormalizeVietnameseString(input);
            StringBuilder output = new StringBuilder();
            int i = 0;

            while (i < input.Length)
            {
                // Kiểm tra các ký tự VNI có độ dài 2 (ví dụ: "aø", "aù", ...)
                if (i + 1 < input.Length && VniToUnicodeMap.ContainsKey(input.Substring(i, 2)))
                {
                    output.Append(VniToUnicodeMap[input.Substring(i, 2)]);
                    i += 2; // Di chuyển 2 ký tự
                }
                // Kiểm tra các ký tự VNI có độ dài 1 (ví dụ: "ì", "í", ...)
                else if (VniToUnicodeMap.ContainsKey(input[i].ToString()))
                {
                    output.Append(VniToUnicodeMap[input[i].ToString()]);
                    i += 1; // Di chuyển 1 ký tự
                }
                else
                {
                    // Giữ nguyên ký tự nếu không có trong bảng ánh xạ
                    output.Append(input[i]);
                    i += 1;
                }
            }

            return output.ToString();
        }





        public static string VniToUni(string str)
        {
            // Bảng ánh xạ VNI sang Unicode
            string[] VNI = {
    "aù", "aø", "aû", "aõ", "aï", "aâ", "aê", "aá", "aà", "aå", "aã", "aä", "aé", "aè", "aú", "aü", "aë",
    "AÙ", "AØ", "AÛ", "AÕ", "AÏ", "AÂ", "AÊ", "AÁ", "AÀ", "AÅ", "AÃ", "AÄ", "AÉ", "AÈ", "AÚ", "AÜ", "AË",
    "eù", "eø", "eû", "eõ", "eï", "eâ", "eá", "eà", "eå", "eã", "eä",
    "EÙ", "EØ", "EÛ", "EÕ", "EÏ", "EÂ", "EÁ", "EÀ", "EÅ", "EÃ", "EÄ",
    "í ", "ì ", "æ ", "ó ", "ò ",
    "Í ", "Ì ", "Æ ", "Ó ", "Ò ",
    "où", "oø", "oû", "oõ", "oï", "oâ", "ô", "oá", "oà", "oå", "oã", "oä", "ôù", "ôø", "ôû", "ôõ", "ôï",
    "OÙ", "OØ", "OÛ", "OÕ", "OÏ", "OÂ", "Ô ", "OÁ", "OÀ", "OÅ", "OÃ", "OÄ", "ÔÙ", "ÔØ", "ÔÛ", "ÔÕ", "ÔÏ",
    "uù", "uø", "uû", "uõ", "uï", "ö ", "öù", "öø", "öû", "öõ", "öï",
    "UÙ", "UØ", "UÛ", "UÕ", "UÏ", "Ö ", "ÖÙ", "ÖØ", "ÖÛ", "ÖÕ", "ÖÏ",
    "yù", "yø", "yû", "yõ", "î ",
    "YÙ", "YØ", "YÛ", "YÕ", "Î ",
    "ñ ", "Ñ "
};

            string[] UNI = {
    "E1", "E0", "1EA3", "E3", "1EA1", "E2", "103", "1EA5", "1EA7", "1EA9", "1EAB", "1EAD", "1EAF", "1EB1", "1EB3", "1EB5", "1EB7",
    "C1", "C0", "1EA2", "C3", "1EA0", "C2", "102", "1EA4", "1EA6", "1EA8", "1EAA", "1EAC", "1EAE", "1EB0", "1EB2", "1EB4", "1EB6",
    "E9", "E8", "1EBB", "1EBD", "1EB9", "EA", "1EBF", "1EC1", "1EC3", "1EC5", "1EC7",
    "C9", "C8", "1EBA", "1EBC", "1EB8", "CA", "1EBE", "1EC0", "1EC2", "1EC4", "1EC6",
    "ED", "EC", "1EC9", "129", "1ECB",
    "CD", "CC", "1EC8", "128", "1ECA",
    "F3", "F2", "1ECF", "F5", "1ECD", "F4", "1A1", "1ED1", "1ED3", "1ED5", "1ED7", "1ED9", "1EDB", "1EDD", "1EDF", "1EE1", "1EE3",
    "D3", "D2", "1ECE", "D5", "1ECC", "D4", "1A0", "1ED0", "1ED2", "1ED4", "1ED6", "1ED8", "1EDA", "1EDC", "1EDE", "1EE0", "1EE2",
    "FA", "F9", "1EE7", "169", "1EE5", "1B0", "1EE9", "1EEB", "1EED", "1EEF", "1EF1",
    "DA", "D9", "1EE6", "168", "1EE4", "1AF", "1EE8", "1EEA", "1EEC", "1EEE", "1EF0",
    "FD", "1EF3", "1EF7", "1EF9", "1EF5",
    "DD", "1EF2", "1EF6", "1EF8", "1EF4",
    "111", "110"
};
            StringBuilder sUni = new StringBuilder();

            for (int i = 0; i < str.Length; i++)
            {
                // Kiểm tra cặp ký tự (2 ký tự)
                if (i + 1 < str.Length)
                {
                    string pair = str.Substring(i, 2);
                    int index = Array.IndexOf(VNI, pair);
                    if (index >= 0)
                    {
                        sUni.Append(ConvertToUnicodeChar(UNI[index]));
                        i++; // Bỏ qua ký tự thứ hai trong cặp
                        continue;
                    }
                }

                // Kiểm tra ký tự đơn (1 ký tự)
                string single = str[i].ToString() + " ";
                int singleIndex = Array.IndexOf(VNI, single);
                if (singleIndex >= 0)
                {
                    sUni.Append(ConvertToUnicodeChar(UNI[singleIndex]));
                }
                else
                {
                    // Giữ nguyên ký tự nếu không có trong bảng ánh xạ
                    sUni.Append(str[i]);
                }
            }

            return sUni.ToString();
        }

        private static char ConvertToUnicodeChar(string hexCode)
        {
            int code = int.Parse(hexCode, System.Globalization.NumberStyles.HexNumber);
            return (char)code;
        }
    }
}
