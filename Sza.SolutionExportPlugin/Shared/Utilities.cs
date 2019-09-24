using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sza.SolutionExportPlugin.Shared
{
    public class Utilities
    {
        public static string Version(string number)
        {
            string[] split = number.Split(new char[] { '.' });
            var target = split[0] + '.' + split[1];
        
            return target;
        }
    }
}
