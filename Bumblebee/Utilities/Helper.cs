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
        public static string GetCellAddress(int column, int row)
        {
            int col = column;
            string colLetter = String.Empty;
            int mod = 0;

            while (col > 0)
            {
                mod = (col - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                col = (int)((col - mod) / 26);
            }

            string address = (colLetter + row);
            return address;
        }

        public static Tuple<int,int> GetCellLocation(string address)
        {
            string columnLetter = Regex.Replace(address, @"[\d-]", string.Empty);
            columnLetter = columnLetter.ToUpper();

            string rowLetters = address.Remove(0, columnLetter.ToArray().Count());

            int sum = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                sum *= 26;
                sum += (columnLetter[i] - 'A' + 1);
            }

            return new Tuple<int, int>(Convert.ToInt32(rowLetters),sum);
        }

    }
}
