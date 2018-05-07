using ExcelDataReader;
using System.Data;
using System.IO;
using System.Text;

namespace xls2csv
{
    class Program
    {
        static void Main(string[] args)
        {
            string xlsPath = args[0];
            string csvPath = args[1];
            var stream = File.Open(xlsPath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = xlsPath.EndsWith(".xlsx")
                ? ExcelReaderFactory.CreateOpenXmlReader(stream)
                : ExcelReaderFactory.CreateBinaryReader(stream);
            var result = excelReader.AsDataSet();
            excelReader.Close();
            converToCSV(csvPath, result, 0);
        }


        private static void converToCSV(string fileCSV, DataSet result, int ind = 0)
        {
            using (StreamWriter csv = new StreamWriter(fileCSV, false))
            {
                int row_no = 0;

                while (row_no < result.Tables[ind].Rows.Count)
                {
                    var row = new StringBuilder();
                    for (int i = 0; i < result.Tables[ind].Columns.Count; i++)
                    {
                        row.Append((i == 0 ? "" : ";") + result.Tables[ind].Rows[row_no][i].ToString());
                    }
                    row_no++;
                    row.Append("\r\n");
                    csv.Write(row);
                }
                csv.Close();
            }
        }
    }
}
