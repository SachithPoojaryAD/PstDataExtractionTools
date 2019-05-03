using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PstDataExtractionTools
{
    public static class Utility
    {
        public static string SplitOnCapitalLetters(this string inputString)
        {
            // starts with an empty string and accumulates the new string into 'result'
            // 'next' is the next character
            return inputString.Aggregate(string.Empty, (result, next) =>
            {
                if (char.IsUpper(next) && result.Length > 0)
                {
                    result += ' ';
                }
                return result + next;
            });
        }
    }
}
