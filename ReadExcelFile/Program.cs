using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelFile
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"E:\Projects\ReadExcelFile\ReadExcelFile\TestFile.XLS"); // Path of Excel file
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            
            //Excel Row Count and Col Count
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            XmlWriter writer = XmlWriter.Create(@"Organizations.xml", settings);
            writer.WriteStartDocument();
            writer.WriteStartElement("Organizations");
            for (int i = 1; i <= rowCount; i++)
            {

                for (int j = 1; j <= 2; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        //var tt = xlRange.Cells[i, j].Value2.ToString(); //For Verification (Reference only)
                        if (i > 1)
                        {
                            writer.WriteStartElement("Organization");
                            writer.WriteElementString("OrganizationCode", xlRange.Cells[i, j].Value2.ToString());
                            writer.WriteElementString("FullName", xlRange.Cells[i, j + 1].Value2.ToString() ?? "");
                            writer.WriteElementString("Address1", xlRange.Cells[i, j + 2].Value2?.ToString());
                            writer.WriteElementString("Addres2", xlRange.Cells[i, j + 3].Value2?.ToString());
                            writer.WriteElementString("Country", xlRange.Cells[i, j + 4].Value2?.ToString());
                            writer.WriteElementString("City", xlRange.Cells[i, j + 5].Value2?.ToString());
                            writer.WriteElementString("PostCode", xlRange.Cells[i, j + 6].Value2?.ToString());
                            writer.WriteElementString("TVA", xlRange.Cells[i, j + 7].Value2?.ToString());
                            writer.WriteElementString("AccountGroup", xlRange.Cells[i, j + 8].Value2?.ToString());
                            writer.WriteElementString("ExternalCode", xlRange.Cells[i, j + 9].Value2?.ToString());
                            writer.WriteElementString("Receivable", xlRange.Cells[i, j + 10].Value2?.ToString());
                            writer.WriteElementString("Payable", xlRange.Cells[i, j + 11].Value2?.ToString());
                            writer.WriteEndElement();

                        }
                    }
                }
            }

            writer.WriteEndElement();
            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            
        }
    }
}