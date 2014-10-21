using System;
using System.Collections.Generic;
using System.Text;

namespace comaddin_xll_rtd_cs
{
    class Log
    {
        public static void WriteLine(string Line)
        {
            System.Diagnostics.Debug.WriteLine("!!! " + Line);
        }
    }
}
