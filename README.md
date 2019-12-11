# ETL-MicrosoftExcel
1. Pre-action > Install ExcelDataReader package + ExcelDataReader.DataSet before you go.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using excel = Microsoft.Office.Interop.Excel;

namespace MicrosoftOfficeInteropEXCEL
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable table = new DataTable("Sales");
            table.Columns.Add("StoreID", typeof(string));
            //table.Columns.Add("Region",typeof(string));
            //table.Columns.Add("District", typeof(string));
            //table.Columns.Add("Territory",typeof(string));
            //table.Columns.Add("Cover", typeof(string));
            table.Columns.Add("Week",typeof(int));
            table.Columns.Add("Revenue", typeof(int));
            table.Columns.Add("Unit", typeof(int));

            // C:\Users\liam.lyu\Desktop\SalesTargetRaw.xlsx
            excel.Application app = new excel.Application();
            excel.Workbook workbook = app.Workbooks.Open(@"C:\Users\liam.lyu\Desktop\raw.xlsx");
            excel.Worksheet sheet = (excel.Worksheet)workbook.Sheets[2];
            excel.Range range = sheet.UsedRange;
            List<int> index = new List<int>();
            int outloop = range.Rows.Count;
            int inloop = range.Columns.Count;
            
            int unit = 0;
            int reve = 0;
            for(int i = 2; i < inloop -1; i ++)
            {
                // log the 2 is the valid read target data column index
                if(range.Cells[1,i].Value2 > range.Cells[1,i+1].Value2)
                {
                    index.Add(i);
                    break;
                }
            }
            //foreach (var i in index) Console.WriteLine(i);
            for (int i = 1; i <= outloop; i++)
            {
                for (int j = 1; j <= inloop; j++)
                {
                    // Log first column shifting
                    /*
                     * if (j == 1)
                    {
                        Console.Write("\r\n");
                    }
                     */
                     if(i == 1)
                    {
                        // log starting from Index = 2
                        if (j >= 2) index.Add(j);
                    }
                    Console.WriteLine("{0} = I , {1} = J ", i,j);
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                    {
                        //Console.Write(range.Cells[i, j].Value2.ToString() + "\t");
                        if (i >= 2 && j >= 2 && j <= index[0])
                        {
                            // log : Store id, week, revenue and unit
                            table.Rows.Add(range.Cells[i, 1].Value2.ToString(),Convert.ToInt32(range.Cells[1, j].Value) , Convert.ToInt64(range.Cells[i, j].Value2), -999);
                        }
                        if (i >= 2 && j >  index[0] && j<= (index[0]*2 -1))
                        {
                            table.Rows.Add(range.Cells[i, 1].Value2.ToString(), Convert.ToInt32(range.Cells[1, j].Value), -999,Convert.ToInt64(range.Cells[i, j].Value2));
                        }
                    }
                }
            }

            foreach (DataRow row in table.Rows)
            {
                Console.WriteLine("--- Row ---");
                foreach (var item in row.ItemArray)
                {
                    Console.Write("Item: "); // Print label.
                    Console.WriteLine(item);
                }
            }

            Console.ReadLine();
            
        }
    }
}
