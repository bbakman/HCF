using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace HCF_Calculation
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path,int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        } 
        public string ReadCell(int i, int j)
        {
            
            i++;
            j++;
            
            if (ws.Cells[i,j].Value2 != null)
                return Convert.ToString(ws.Cells[i, j].Value2);
            
            else
                return "";
             
        }
        //public List<double> Readcolum()
        //{
        //    var range2 = ws.UsedRange
        //}
    }
}
