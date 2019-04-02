using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PstDataExtractionTools
{
    public class InvalidFilePathException : Exception
    {
        public InvalidFilePathException()
        {
        }
        public InvalidFilePathException(string msg) : base(msg)
        {
        }
    }
}
