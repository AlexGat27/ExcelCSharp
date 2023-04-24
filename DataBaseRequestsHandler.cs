using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace Database
{
    class DataBaseRequestsHandler : IDataBaseRequestsHandler
    {
        private string filename = string.Empty;
        private string[,] excelTable;
        private int totalRows = 0;
        private int totalCol = 0;
        ExcelPackage excelFile;
        ExcelWorksheet worksheet;
        public DataBaseRequestsHandler(string _path)
        {
            filename = _path;
            excelFile = new ExcelPackage(new FileInfo(filename));
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            worksheet = excelFile.Workbook.Worksheets[0];
            totalRows = worksheet.Dimension.End.Row;
            totalCol = worksheet.Dimension.End.Column;
        }
        public void ReadExcelFile()
        {
            excelTable = new string[totalRows, totalCol];
            for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
            {
                IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalCol].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                List<string> list = row.ToList<string>();

                for (int i = 0; i < list.Count; i++)
                {
                    excelTable[rowIndex - 1, i] = list[i];
                }
            }
        }
        public void DisplayTable()
        {
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalCol; j++)
                {
                    Console.Write(excelTable[i, j] + "  ");
                }
                Console.Write("\n");
            }
        }
        public string[,] GetInfoFromRequest(string request, int colonRequest)
        {
            int countTrueRows = 0;
            for (int i = 1; i < totalRows; i++)
            {
                if (excelTable[i, colonRequest] == request)
                    countTrueRows++;
            }
            string[,] dataRequest = new string[countTrueRows, totalCol];
            countTrueRows = 0;
            for (int i = 1; i < totalRows; i++)
            {
                if (excelTable[i, colonRequest] == request)
                {
                    for (int j = 0; j < totalCol; j++)
                    {
                        dataRequest[countTrueRows, j] = excelTable[i, j];
                    }
                    countTrueRows++;
                }
            }
            return dataRequest;
        }
        public void SetRow(string[] row)
        {
            worksheet.Cells[totalRows + 1, 1, totalRows + 1, totalCol].Value = row;
        }
    }
}
