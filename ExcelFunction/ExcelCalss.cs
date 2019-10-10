using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SHDocVw;

namespace ExcelFunction
{
    public class ExcelCalss
    {
        public static void WriteExcel1(string ExcelName, string Province, string WenLi, string Year, string MajorName, string Score, string School, string PiCi)
        {
            ShellWindows windows = new ShellWindows();
            Microsoft.Office.Interop.Excel.Workbook wb = null;

            Microsoft.Office.Interop.Excel.Application ExcelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            foreach (Microsoft.Office.Interop.Excel.Workbook item in ExcelApp.Workbooks)
            {
                if (item.Name.Contains(ExcelName))
                {
                    wb = item;
                    break;
                }
            }

            Microsoft.Office.Interop.Excel.Range Rng = wb.Sheets[1].Range("F100000").End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Offset[1, 0];

            Rng.Value = MajorName;
            Rng.Offset[0, -5].Value = Province;
            Rng.Offset[0, -4].Value = Year;
            Rng.Offset[0, -3].Value = School;
            Rng.Offset[0, -2].Value = WenLi;
            Rng.Offset[0, -1].Value = PiCi;
            Rng.Offset[0, 1].Value = Score;
        }
    }
}
