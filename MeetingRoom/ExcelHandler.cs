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
        public static void GetDataFromExcel(string date, int fromIndex, int toIndex)
        {
            try
            {
                path = ConfigurationManager.AppSettings["ExcelPath"];
                string sheetName = String.Empty;
                xlApp = new Application();
                xlWorkBook = xlApp.Workbooks.Open(path, ReadOnly: false);

                
                //Finding the correct worksheet
                foreach (Worksheet workSheet in xlWorkBook.Worksheets)
                {
                    sheetName = workSheet.Name;

                    if (sheetName.Equals(date))
                    {
                        xlWorkSheet = workSheet;
                        break;
                    }
                }

                if(xlWorkSheet == null)
                {
                    Worksheet xlWorkSheetSource = (Worksheet)xlWorkBook.Sheets["Template"];
                    xlWorkSheet = (Worksheet)xlWorkBook.Sheets[1];
                    xlWorkSheetSource.Copy(xlWorkSheet);
                    xlWorkSheet.Name = date;
                    xlWorkSheetSource = (Worksheet)xlWorkBook.Sheets[1];
                    xlWorkSheetSource.Name = "Template";
                }



            }
            catch (Exception ex)
            {
                ExceptionLogging.SendErrorToText(ex);
            }
            finally
            {
                xlWorkBook.Save();
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
