using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Database
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            string path = "C:/Users/sanya/Programs/Hahaton/BD.xlsx";
            DataBaseRequestsHandler excelfile = new DataBaseRequestsHandler(path);
            excelfile.ReadExcelFile();
            string[,] data = excelfile.GetInfoFromRequest("qqefv", 2);
            string[] row = new string[3];
            for(int i = 0; i < 3; i++)
            {
                row[i] = data[0, i];
            }
            excelfile.SetRow(row);
        }
    }
}
