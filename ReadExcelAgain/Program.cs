using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelAgain
{
    enum DataType
    {
        ValueData,
        Menus,
        Space,
        Shift
    }
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Read Start!");
            List<VenueData> venueDatas = ReadVenue(DataType.ValueData);
            List<Menu> menus = ReadMenu(DataType.Menus);
            List<Shift> shifts = ReadShift(DataType.Shift);
            List<Space> spaces = ReadSpace(DataType.Space);

            Console.WriteLine("Read End!");

        }

        public static List<VenueData> ReadVenue(Enum DataType)
        {
            List<VenueData> venueDatas = new List<VenueData>(); 
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp != null)
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\test.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Sheets[Convert.ToInt16(DataType)+1];

                Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    var venueData = new VenueData();
                
                    if (i == 1) { continue;  }

                        for (int j = 1; j <= colCount; j++)
                        {
                            Microsoft.Office.Interop.Excel.Range range = (excelWorksheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range);

                            var RawcellValue = range.Value;
                            string cellValue = string.Empty;

                            if (RawcellValue != null)
                            {
                                cellValue = range.Value.ToString();
                            }

                            if (j == 1)
                            {
                                venueData.Sl = Convert.ToInt32(cellValue);
                            }
                            else if (j == 2)
                            {
                                venueData.Name = cellValue;
                            }
                            else if (j == 3)
                            {
                                venueData.Address1 = cellValue;

                            }
                            else if (j == 4)
                            {
                                venueData.Address2 = cellValue;

                            }
                            else if (j == 5)
                            {
                                venueData.Phone = cellValue;

                            }
                            else if (j == 6)
                            {
                                venueData.GEOLocation = cellValue;

                            }
                            else if (j == 7)
                            {
                                venueData.Location = cellValue;

                            }
                        }

                    venueDatas.Add(venueData);
                }

                excelWorkbook.Close();
                excelApp.Quit();
            }
            return venueDatas;
        }

        public static List<Menu> ReadMenu(Enum DataType)
        {
            List<Menu> menus = new List<Menu>();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp != null)
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\test.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Sheets[Convert.ToInt16(DataType) + 1];

                Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                   
                    var menu = new Menu();
                    if (i == 1) { continue; }

                    for (int j = 1; j <= colCount; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range range = (excelWorksheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range);

                        var RawcellValue = range.Value;
                        string cellValue = string.Empty;

                        if (RawcellValue != null)
                        {
                            cellValue = range.Value.ToString();
                        }

                        if (j == 1)
                        {
                            menu.VenueSl = Convert.ToInt32(cellValue);
                        }
                        else if (j == 2)
                        {
                            menu.Name = cellValue;
                        }
                        else if (j == 3)
                        {
                            menu.MenuDetails = cellValue;

                        }
                        else if (j == 4)
                        {
                            menu.Price =cellValue;

                        }
                        else if (j == 5)
                        {
                            menu.Vat = cellValue;

                        }
                        else if (j == 6)
                        {
                            menu.Sc = cellValue;

                        }
                       
                    }

                    menus.Add(menu);
                }

                excelWorkbook.Close();
                excelApp.Quit();
            }
            return menus;
        }

        public static List<Shift> ReadShift(Enum DataType)
        {
            List<Shift> shifts = new List<Shift>();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp != null)
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\test.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Sheets[Convert.ToInt16(DataType) + 1];

                Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    var shift = new Shift();

                    if (i == 1) { continue; }

                    for (int j = 1; j <= colCount; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range range = (excelWorksheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range);

                        var RawcellValue = range.Value;
                        string cellValue = string.Empty;

                        if (RawcellValue != null)
                        {
                            cellValue = range.Value.ToString();
                        }

                        if (j == 1)
                        {
                            shift.SapceSerial = Convert.ToInt32(cellValue);
                        }
                        else if (j == 2)
                        {
                            shift.Morning = cellValue;
                        }
                        else if (j == 3)
                        {
                            shift.Afternoon = cellValue;

                        }
                        else if (j == 4)
                        {
                            shift.Evening = cellValue;

                        }
                       
                    }

                    shifts.Add(shift);
                }

                excelWorkbook.Close();
                excelApp.Quit();
            }
            return shifts;
        }

        public static List<Space> ReadSpace(Enum DataType)
        {
            List<Space> spaces = new List<Space>();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp != null)
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\test.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Sheets[Convert.ToInt16(DataType) + 1];

                Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    var space = new Space();

                    if (i == 1) { continue; }

                    for (int j = 1; j <= colCount; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range range = (excelWorksheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range);

                        var RawcellValue = range.Value;
                        string cellValue = string.Empty;

                        if (RawcellValue != null)
                        {
                            cellValue = range.Value.ToString();
                        }

                        if (j == 1)
                        {
                            space.SpaceSerial = Convert.ToInt32(cellValue);
                        }
                        else if (j == 2)
                        {
                            space.VenueSerial = Convert.ToInt32(cellValue);
                        }
                        else if (j == 3)
                        {
                            space.FloorName = cellValue;

                        }
                        else if (j == 4)
                        {
                            space.Price = cellValue;

                        }
                        else if (j == 5)
                        {
                            space.Vat = cellValue;

                        }
                        else if (j == 6)
                        {
                            space.DiningCapacity = cellValue;

                        }
                        else if (j == 7)
                        {
                            space.WaitingCapacity = cellValue;

                        }
                        else if (j == 8)
                        {
                            space.TotalCapacity = Convert.ToInt32(cellValue);

                        }
                        else if (j == 9)
                        {
                            space.Amenity = cellValue;

                        }
                        else if (j == 10)
                        {
                            space.Parking = cellValue;

                        }
                    }

                    spaces.Add(space);
                }

                excelWorkbook.Close();
                excelApp.Quit();
            }
            return spaces;
        }
    }
}
