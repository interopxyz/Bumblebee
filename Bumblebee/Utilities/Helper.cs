using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Bumblebee
{
    public static class Helper
    {
        public static string GetCellAddress(int column, int row, bool absoluteColumn = false, bool absoluteRow = false)
        {
            int col = column+1;
            string colLetter = String.Empty;
            int mod;

            while (col > 0)
            {
                mod = (col - 1) % 26;
                colLetter = ((char)(65 + mod)).ToString() + colLetter;
                col = (int)((col - mod) / 26);
            }

            string ac = "";
            string ar = "";

            if (absoluteColumn) ac = "$";
            if (absoluteRow) ar = "$";

            string address = (ac+colLetter + ar+(row+1));
            return address;
        }

        public static Tuple<int,int> GetCellLocation(string address)
        {

            address.Replace("$", "");

            char[] arrC = address.ToCharArray();
            string Lint = "";
            string Lstr = "";
            int retVal = 0;

            for (int i = 0; i < arrC.Length; i++)
            {
                if (char.IsNumber(arrC[i]))
                {
                    Lint += arrC[i];
                }
                else
                {
                    Lstr += arrC[i];
                }
            }

            string col = Lstr.ToUpper();
            int k = col.Length - 1;

            for (int i = 0; i < k + 1; i++)
            {
                char colPiece = col[i];
                int t = (int)colPiece;
                int colNum = t - 64;
                retVal = retVal + colNum * (int)(Math.Pow(26, col.Length - (i + 1)));
            }

            return new Tuple<int, int>(retVal-1, Convert.ToInt32(Lint));
        }

    }
}
