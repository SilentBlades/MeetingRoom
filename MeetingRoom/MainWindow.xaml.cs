using System;
using System.Diagnostics;
using System.Windows;

namespace MeetingRoom
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ResizeMode = ResizeMode.CanMinimize;
        }

        #region
        /*
         * submitButton_Click():
         * This method carries out data collection from GUI and sends it to the respective methods.
         */
        private void submitButton_Click(object sender, RoutedEventArgs e)
        {
            String dateInput = dateDatePicker.SelectedDate.ToString(); //Get data from DatePicker

            /*
             * Business logic to figure out time slot for the requested date
             */


            ExcelHandler.GetDataFromExcel();
            ExcelHandler.WriteDataToExcel();
        }
        #endregion

        #region
        /*
         * ResetButton_Click():
         * This method closes the current application and opens a new blank application.
         */
        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(Application.ResourceAssembly.Location);
            Application.Current.Shutdown();
        }
        #endregion
    }
}