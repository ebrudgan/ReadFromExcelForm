using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ReadFromExcelForm
{
    public static class StringExtensions
    {

        public static string ToLowerAndTurkishCharacter(this string text)
        {
            text = text.ToLower();
            text = Regex.Replace(text, @"\s", "");
            text = text.Replace("ö", "o");
            text = text.Replace("ü", "u");
            text = text.Replace("ı", "i");
            text = text.Replace("ğ", "g");
            text = text.Replace("ö", "o");
            text = text.Replace("ç", "c");
            return text;
        }
    }
}
