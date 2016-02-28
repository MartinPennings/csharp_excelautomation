using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Interop;


namespace AutoExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            // Generate some testdata
            Dictionary<string,int[]> dict = new Dictionary<string,int[]>();
            for (int i = 1; i <=500; i+=1) {
                dict.Add("Number " + i.ToString(),new int[5] { 100+i, 200+i, 300+i, 400+i, 500+i });
            }

            // Create the Excel Sheet
            var excelApp = new Office.Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();

            // dynamic: types will not be checked at compile time
            dynamic workSheet = excelApp.ActiveSheet;
            
            try
            {
                // Generate the header
                workSheet.Cells[1, "A"] = "MY AUTOMATED EXCELSHEET";
                workSheet.Cells[2, "A"] = "Demonstrating Excel Automation";
                workSheet.Cells[4, "A"] = "FIELD NAME";
                workSheet.Cells[4, "B"] = "VALUE #1";
                workSheet.Cells[4, "C"] = "VALUE #2";
                workSheet.Cells[4, "D"] = "VALUE #3";
                workSheet.Cells[4, "E"] = "VALUE #4";
                workSheet.Cells[4, "F"] = "VALUE #5";


                // generate the value fields
                int row = 5;

                foreach (KeyValuePair<string, int[]> val in dict)
                {
                    workSheet.Cells[row, "A"] = val.Key;
                    workSheet.Cells[row, "B"] = val.Value[0];
                    workSheet.Cells[row, "C"] = val.Value[1];
                    workSheet.Cells[row, "D"] = val.Value[2];
                    workSheet.Cells[row, "E"] = val.Value[3];
                    workSheet.Cells[row, "F"] = val.Value[4];

                    row += 1;
                }
                
                //  AutoSize the columns
                for (int i = 1; i <= 5; i++)
                {
                    workSheet.Columns[i].AutoFit();
                }

                // Some layout:

                // Font type
                workSheet.Range(workSheet.Cells(1, 1), workSheet.Cells(row, 6)).Font.Name = "Arial";
                workSheet.Range(workSheet.Cells(1, 1), workSheet.Cells(1, 1)).Font.Size = 16;
                workSheet.Range(workSheet.Cells(2, 1), workSheet.Cells(row, 6)).Font.Size = 8;
                // Bold text:
                workSheet.Rows(4).Font.Bold = true;


                // Freeze the header
                workSheet.Application.ActiveWindow.SplitRow = 4;
                workSheet.Application.ActiveWindow.FreezePanes = true;

          
                // Select the first cell
                workSheet.Cells(1, 1).Select();
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel reported the following Error:");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Release the com object
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
            Console.WriteLine("Ready.");
            Console.Read();
        }
        
    }
}
