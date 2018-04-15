using Microsoft.Office.Interop.Excel;
using System;
using System.Configuration;

namespace MeetingRoom
{
    class ExcelHandler
    {
        #region
        /*
         * GetDataFromExcel():
         * Gets the list of available rooms from a excel file.
         */
        public static void GetDataFromExcel()
        {
            try
            {
                Application xlApp;
                Workbook xlWorkBook;
                Worksheet xlWorkSheet;
                Range range;

                string path = ConfigurationManager.AppSettings["ExcelPath"];
                xlApp = new Application();
                xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendErrorToText(ex);
            }
            finally
            {
                xlWorkBook.Close(true, null, null);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }
        #endregion

        #region
        /*
         * WriteDataToExcel():
         * Writes the booked room details to the excel file.
         */
        public static void WriteDataToExcel()
        {
            //TODO
        }
        #endregion
    }
}
