using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Runtime.InteropServices;

namespace MeetingRoom
{
    class ExcelHandler
    {
        private static Application xlApp;
        private static Workbook xlWorkBook;
        private static Worksheet xlWorkSheet;
        private static string path;

        #region GetDataFromExcel
        /*
         * GetDataFromExcel():
         * Gets the list of available rooms from a excel file.
         */
        public static List<string> GetDataFromExcel(DateTimeSlot dts)
        {
            int cellFromIndex = dts.From + 1;
            int cellToIndex = dts.To + 1;
            int rowIterator = 2;
            bool flag = false;
            List<string> meetingRoomList = new List<string>();
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

                    if (sheetName.Equals(dts.Date))
                    {
                        xlWorkSheet = workSheet;
                        break;
                    }
                }

                if (xlWorkSheet == null)
                {
                    Worksheet xlWorkSheetSource = (Worksheet)xlWorkBook.Sheets["Template"];
                    xlWorkSheet = (Worksheet)xlWorkBook.Sheets[1];
                    xlWorkSheetSource.Copy(xlWorkSheet);
                    xlWorkSheet.Name = dts.Date;
                    xlWorkSheetSource = (Worksheet)xlWorkBook.Sheets[1];
                    xlWorkSheetSource.Name = "Template";
                }

                //Read cells to display Meeting Rooms
                for (; rowIterator <= Convert.ToInt32(ConfigurationManager.AppSettings["MeetingRoomCount"]) + 1; rowIterator++)
                {
                    for (int i = cellFromIndex; i <= cellToIndex; i++)
                    {
                        if (xlWorkSheet.Cells[rowIterator, i].Value == 0) { flag = true; }
                        else { flag = false; break; }
                    }
                    if (flag) { meetingRoomList.Add(Convert.ToString(xlWorkSheet.Cells[rowIterator, 1].Value)); }
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

            return meetingRoomList;
        }
        #endregion

        #region WriteDataToExcel()
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
