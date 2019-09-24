using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sza.SolutionExportPlugin.Shared.Enum
{
    public class Constants
    {
        public static string[] friendlyName = new[] { "Active", "Default", "Basic" };
        public static object[] exportSettings = new object[] {
            "Auto-numbering",
            "Calendar",
            "Customization",
            "Email tracking",
            "General",
            "Marketing",
            "Outlook Synchronization",
            "Relationship Roles",
            "ISV Config",
            "Sales"};
        public static object[] targets82 = new object[] {"8.2","8.1","8.0" };
        public static object[] targets81 = new object[] { "8.1", "8.0" };
        public static object[] targets80 = new object[] { "8.0" };
        public static object[] targets90 = new object[] { "9.*" };
    }
}
