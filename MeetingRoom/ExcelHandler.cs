using Microsoft.Office.Interop.Excel;
using System;
using System.Configuration;
using System.Runtime.InteropServices;

namespace MeetingRoom
{
    class ExcelHandler
    {
        private static Application xlApp;
        private static Workbook xlWorkBook;
        private static Worksheet xlWorkSheet;
        private static Range range;
        private static string path;
        #region
        /*
         * GetDataFromExcel():
         * Gets the list of available rooms from a excel file.
         */
        public static void GetDataFromExcel()
        {
            try
            {
                path = ConfigurationManager.AppSettings["ExcelPath"];
                xlApp = new Application();
                xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
                string date = xlWorkSheet.Name;

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
