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

        public static Application XlApp { get => xlApp; set => xlApp = value; }
        public static Workbook XlWorkBook { get => xlWorkBook; set => xlWorkBook = value; }
        public static Worksheet XlWorkSheet { get => xlWorkSheet; set => xlWorkSheet = value; }

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
                XlApp = new Application();
                XlWorkBook = XlApp.Workbooks.Open(path, ReadOnly: false);


                //Finding the correct worksheet
                foreach (Worksheet workSheet in XlWorkBook.Worksheets)
                {
                    sheetName = workSheet.Name;

                    if (sheetName.Equals(dts.Date))
                    {
                        XlWorkSheet = workSheet;
                        break;
                    }
                }

                if (XlWorkSheet == null)
                {
                    Worksheet xlWorkSheetSource = (Worksheet)XlWorkBook.Sheets["Template"];
                    XlWorkSheet = (Worksheet)XlWorkBook.Sheets[1];
                    xlWorkSheetSource.Copy(XlWorkSheet);
                    XlWorkSheet.Name = dts.Date;
                    xlWorkSheetSource = (Worksheet)XlWorkBook.Sheets[1];
                    xlWorkSheetSource.Name = "Template";
                }

                //Read cells to display Meeting Rooms
                for (; rowIterator <= Convert.ToInt32(ConfigurationManager.AppSettings["MeetingRoomCount"]) + 1; rowIterator++)
                {
                    for (int i = cellFromIndex; i <= cellToIndex; i++)
                    {
                        if (XlWorkSheet.Cells[rowIterator, i].Value == 0) { flag = true; }
                        else { flag = false; break; }
                    }
                    if (flag) { meetingRoomList.Add(Convert.ToString(XlWorkSheet.Cells[rowIterator, 1].Value)); }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendErrorToText(ex);
            }

            return meetingRoomList;
        }
        #endregion

        #region WriteDataToExcel()
        /*
         * WriteDataToExcel():
         * Writes the booked room details to the excel file.
         */
        public static bool WriteDataToExcel(DateTimeSlot dts)
        {
            bool status = true;
            try
            {
                int diff = dts.To - dts.From;
                int colIterator = dts.From + 1;
                int rowIndex = -1;

                for (int j = 1; j < 11; j++)
                {
                    string temp = XlWorkSheet.Cells[j + 1, 1].text;
                    if (XlWorkSheet.Cells[j + 1, 1].text.Contains(dts.MeetingRoomSelected)) { rowIndex = j + 1; break; }
                }

                for (int i = 0; i < diff; i++, colIterator++)
                {
                    XlWorkSheet.Cells[rowIndex, colIterator] = 1;
                }
            }
            catch (Exception ex)
            {
                status = false;
                ExceptionLogging.SendErrorToText(ex);
            }
            finally
            {
                XlWorkBook.Save();
                XlWorkBook.Close(true, null, null);
                XlApp.Quit();
                Marshal.ReleaseComObject(XlWorkSheet);
                Marshal.ReleaseComObject(XlWorkBook);
                Marshal.ReleaseComObject(XlApp);
            }

            return status;
        }
        #endregion
    }
}
