using System;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace SplitReport
{
    class Program
    {
        //https://www.c-sharpcorner.com/UploadFile/bd6c67/how-to-create-excel-file-using-C-Sharp/
        static void Main(string[] args)
        {
            var program = new Program();
            //@"C:\Users\bdarley\Documents\Med - Foundation October 2018 Reports Workbook.xlsx"
            //var masterFilePath = @"z:\master.xlsx"; //program.GetMasterFilePath();
            var destinationFolder = @"z:\trash";

            //var masterFilePath = program.GetMasterFilePath();
            //var destinationFolder = program.GetDestinationFolder();

            //var records = program.GetMasterTableRecords(masterFilePath);
            var contents = System.IO.File.ReadAllText(@"z:\raw-master-records.txt");
            var records = Newtonsoft.Json.JsonConvert.DeserializeObject<IEnumerable<MasterDataRecord>>(contents);
            //var serialized = Newtonsoft.Json.JsonConvert.SerializeObject(records);
            //System.IO.File.WriteAllText(@"z:\raw-master-records.txt", serialized);
            //return;
            foreach (var masterDataRecord in records)
            {
                masterDataRecord.Cash = Math.Round(masterDataRecord.Cash,0);
                masterDataRecord.Graystone = Math.Round(masterDataRecord.Graystone, 0);
                masterDataRecord.TotalAvailable = Math.Round(masterDataRecord.TotalAvailable, 0);
            }


            var groups = records.Where(c => !string.IsNullOrEmpty(c.DepartmentName) && (c.Cash != 0 && c.TotalAvailable != 0 && c.Graystone != 0)).GroupBy(c => c.DepartmentName);

            foreach (var group in groups.OrderBy(c => c.Key))
            {
                var fileName = System.IO.Path.Combine(destinationFolder, $"{group.Key}.xlsx");
                program.CreateSheet(fileName, group.Select(c => c).ToList());

            }

        }



        private void CreateSheet(string fileName, IReadOnlyCollection<MasterDataRecord> records)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel._Worksheet sheet = null;
            Excel.Range cellRange = null;
            try
            {
                Console.WriteLine(new string('-', 60));
                Console.WriteLine($"Creating File {fileName}, {records.Count()} records");
                if (System.IO.File.Exists(fileName))
                {
                    System.IO.File.Delete(fileName);
                }
                app = new Excel.Application { Visible = false, DisplayAlerts = false };
                workbook = app.Workbooks.Add(Type.Missing);

                sheet = (Excel.Worksheet)workbook.ActiveSheet;
                sheet.Name = "Fund Summary";


                sheet.PageSetup.PrintTitleRows = "$1:$9";
                sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLetter;
                sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                sheet.PageSetup.Zoom = false;
                sheet.PageSetup.FitToPagesTall = false;
                sheet.PageSetup.FitToPagesWide = 1;

                CreateReportHeader(sheet, records.First().DepartmentName);
                CreateDataHeader(sheet);
                var rowCount = CreateMainDataContent(records, sheet);

                CreateDataTotalFooter(records, sheet, rowCount);


                //cellRange = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, ExportToExcel().Columns.Count]];
                fileName = fileName.Replace("/", "-");
                fileName = fileName.Replace("&", "and");
                workbook.SaveAs(fileName); ;
                Console.WriteLine($"File {fileName} Created");


            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                if (sheet != null)
                {
                    Marshal.ReleaseComObject(sheet);
                }

                if (workbook != null)
                {
                    //close and release
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }

                //quit and release
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }


            }
        }

        private static void CreateDataTotalFooter(IReadOnlyCollection<MasterDataRecord> records, Excel._Worksheet sheet, int rowCount)
        {

            


            rowCount++;

            sheet.Cells[rowCount, 3] = records.Sum(c => c.Cash);
            sheet.Cells[rowCount, 4] = records.Sum(c => c.Graystone);
            sheet.Cells[rowCount, 5] = records.Sum(c => c.TotalAvailable);

            ((Excel.Range)sheet.Cells[rowCount, 3]).Formula = "=SUM(C8:C" + (rowCount - 1) + ")";
            ((Excel.Range)sheet.Cells[rowCount, 4]).Formula = "=SUM(D8:D" + (rowCount - 1) + ")";
            ((Excel.Range)sheet.Cells[rowCount, 5]).Formula = "=SUM(E8:E" + (rowCount - 1) + ")";


            var range = sheet.Cells.Range[sheet.Cells[rowCount, 3], sheet.Cells[rowCount, 5]];
            range.WrapText = true;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            
            range.EntireRow.Font.Bold = true;
            //range.Style = "Currency";
            var numberFormat = "$_(* #,##0_);$_(* (#,##0);_(* \" -\"??_);$_(@_)";
            range.Cells.NumberFormat = numberFormat;

            //var totalColumn = (Excel.Range) sheet.Cells[rowCount, 5];
            //.Font.Color = Color.Blue;
            var totalColumn = (Excel.Range) sheet.Cells[rowCount, 5];
            totalColumn.Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);


            var border = range.Borders;
            var bottomBorder = border[Excel.XlBordersIndex.xlEdgeBottom];
            bottomBorder.LineStyle = Excel.XlLineStyle.xlContinuous;
            bottomBorder.Weight = 2d;
        }

        private int CreateMainDataContent(IReadOnlyCollection<MasterDataRecord> records, Excel._Worksheet sheet)
        {
            Excel.Range cellRange;
            var rowCount = 8;

            foreach (var record in records.OrderBy(c=> c.PeopleSoftSource))
            {
                sheet.Cells[rowCount, 1] = record.PeopleSoftSource;
                sheet.Cells[rowCount, 2] = record.FundTitle;
                sheet.Cells[rowCount, 3] = (int)record.Cash;
                sheet.Cells[rowCount, 4] = (int)record.Graystone;
                sheet.Cells[rowCount, 5] = (int)record.TotalAvailable;

                var range = sheet.Cells.Range[sheet.Cells[rowCount, 3], sheet.Cells[rowCount, 5]];
                range.WrapText = true;
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                if (rowCount == 8)
                {

                    //range.Style = "Currency";
                    //var numberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* \" - \"??_);_(@_)";
                    //range.Cells.NumberFormat = numberFormat;
                    var numberFormat = "$_(* #,##0_);$_(* (#,##0);_(* \" -\"??_);$_(@_)";
                    range.Cells.NumberFormat = numberFormat;
                }
                else
                {
                    range.WrapText = true;
                    range.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                    //How to format cells
                    //https://support.office.com/en-us/article/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-US&rs=en-US&ad=US
                    //
                    //var numberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* \" -\"??_);_(@_)";
                    var numberFormat = "_(* #,##0_);_(* (#,##0);_(* \" -\"??_);_(@_)";
                    range.Cells.NumberFormat = numberFormat;
                }

                rowCount++;
            }
            //

            ((Excel.Range)sheet.Cells[1, 1]).ColumnWidth = 8;
            ((Excel.Range)sheet.Cells[1, 2]).EntireColumn.AutoFit();
            ((Excel.Range)sheet.Cells[1, 3]).ColumnWidth = 16;
            ((Excel.Range)sheet.Cells[1, 4]).ColumnWidth = 19.43;
            ((Excel.Range)sheet.Cells[1, 5]).ColumnWidth = 16;
            //Subtract one to get the border
            rowCount--;
            cellRange = sheet.Range[sheet.Cells[1, 3], sheet.Cells[rowCount, 5]];
            //cellRange.Columns.ColumnWidth = 20;


            //8
            //69.86
            //16
            //19.43
            //16

            var border = cellRange.Borders;
            var bottomBorder = border[Excel.XlBordersIndex.xlEdgeBottom];
            bottomBorder.LineStyle = Excel.XlLineStyle.xlContinuous;
            bottomBorder.Weight = 2d;
            //Re-position where everything should be
            rowCount++;
            return rowCount;
        }

        private void CreateDataHeader(Excel._Worksheet sheet)
        {
            sheet.Cells[6, 1] = "Source";
            sheet.Cells[6, 2] = "Title";
            sheet.Cells[6, 3] = "Cash";
            sheet.Cells[6, 4] = "Graystone Consulting Investments";
            sheet.Cells[6, 5] = "Total Available Balance";

            var range = sheet.Cells.Range[sheet.Cells[6, 1], sheet.Cells[6, 5]];
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            range.WrapText = true;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            
            range.Columns.RowHeight = 45;
            

            var border = range.Borders;
            var bottomBorder = border[Excel.XlBordersIndex.xlEdgeBottom];
            bottomBorder.LineStyle = Excel.XlLineStyle.xlContinuous;
            bottomBorder.Weight = 2d;

        }

        private void CreateReportHeader(Excel._Worksheet sheet, string departmentName)
        {
            void FormatCell(Excel.Range range)
            {
                range.Merge();
                range.Select();
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                range.WrapText = true;

                range.Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);

                range.Font.Name = "Calibri";
                range.Font.Bold = true;
                range.Font.Size = 11;
            }
            const string titleHeader = "The Medical Foundation of North Carolina, Inc.";

            var header1 = sheet.Cells.Range[sheet.Cells[1, 1], sheet.Cells[1, 5]];
            var header2 = sheet.Cells.Range[sheet.Cells[2, 1], sheet.Cells[2, 5]];
            var header3 = sheet.Cells.Range[sheet.Cells[3, 1], sheet.Cells[3, 5]];
            var header4 = sheet.Cells.Range[sheet.Cells[4, 1], sheet.Cells[4, 5]];

            header1.Cells[1, 1] = titleHeader;
            header1.Cells[2, 1] = departmentName;
            header1.Cells[3, 1] = "Fund Summary";
            header1.Cells[4, 1] = $"As of {GetReportDate():MMMM d, yyyy}";

            FormatCell(header1);
            FormatCell(header2);
            FormatCell(header3);
            FormatCell(header4);


        }

        private DateTime GetReportDate()
        {
            var firstOfMonth = DateTime.Parse($"{DateTime.Now.Month}/1/{DateTime.Now.Year}");
            //firstOfMonth = firstOfMonth.AddMonths(-1);
            var lastDayOfReportMonth = firstOfMonth.AddDays(-1);
            return lastDayOfReportMonth;

        }

        private IEnumerable<MasterDataRecord> GetMasterTableRecords(string filePath)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel._Worksheet sheet = null;

            try
            {
                app = new Excel.Application();
                workbook = app.Workbooks.Open(filePath);

                var sheets = workbook.Sheets.Cast<Excel._Worksheet>();

                sheet = sheets.FirstOrDefault(c => c.Name.ToUpper() == "MASTER");
                if (sheet == null)
                {
                    return Enumerable.Empty<MasterDataRecord>();
                }
                //var cashSheet = sheets.FirstOrDefault(c => c.Name.ToUpper() == "CASH");
                var rowCount = sheet.Rows.Count;
                var colCount = 20; //masterSheet.Columns.Count;
                                   //Department Name	PeopleSoft Source #	Fund Title	 Cash 	 Graystone 	 Total Available Expendable Bal 
                var skipCount = 0;
                Console.WriteLine("Processing Master Data Sheet...");
                var masterRecords = new List<MasterDataRecord>();
                for (var i = 2; i <= rowCount; i++)
                {
                    var masterRecord = new MasterDataRecord();
                    for (int j = 1; j < colCount; j++)
                    {
                        var range = sheet.UsedRange;

                        var cell = (Excel.Range)range.Cells[i, j];
                        if (cell != null && cell.Value2 != null)
                        {
                            if (j == 1)
                                masterRecord.DepartmentName = (string)cell.Value2;
                            if (j == 2)
                                masterRecord.PeopleSoftSource = (string)cell.Value2;
                            if (j == 3)
                                masterRecord.FundTitle = (string)cell.Value2;
                            if (j == 4)
                                masterRecord.Cash = (double)cell.Value2;
                            if (j == 5)
                                masterRecord.Graystone = (double)cell.Value2;
                            if (j == 6)
                                masterRecord.TotalAvailable = (double)cell.Value2;
                        }
                    }

                    if (!string.IsNullOrEmpty(masterRecord.PeopleSoftSource))
                    {
                        skipCount = 0;
                        masterRecords.Add(masterRecord);

                        Console.WriteLine($"Row {i}, Department {masterRecord.DepartmentName}");
                    }
                    else
                    {
                        Console.WriteLine($"Skipping Row {i}, no data to import");
                    }

                    skipCount++;
                    if (skipCount > 20)
                    {
                        break;
                    }
                }

                return masterRecords;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                if (sheet != null)
                {
                    Marshal.ReleaseComObject(sheet);
                }

                if (workbook != null)
                {
                    //close and release
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }

                //quit and release
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }


            }
        }
        private string GetDestinationFolder()
        {
            Console.WriteLine("Specify the destination path where the reports will be generated from the master file and press (enter)");
            var filePath = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                filePath = string.Empty;
            }
            var exists = System.IO.Directory.Exists(filePath);

            if (exists) return filePath;

            Console.WriteLine("Path does not exist, Would you like to create the path? (Y/N)");
            var character = Console.ReadKey();
            if (character.KeyChar.ToString().ToUpper() == "Y")
            {
                System.IO.Directory.CreateDirectory(filePath);
            }
            else
            {
                return GetDestinationFolder();
            }

            return filePath;

        }

        private string GetMasterFilePath()
        {
            Console.WriteLine("Specify the complete file path of the master excel file and press (enter)");
            var filePath = Console.ReadLine();
            var exists = System.IO.File.Exists(filePath);
            if (!exists)
            {
                Console.WriteLine("File could not be located, please try again.");
                return GetMasterFilePath();
            }

            return filePath;

        }
    }
}
