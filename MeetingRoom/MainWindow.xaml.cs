﻿using System.Collections.Generic;
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
            SetFromTimerComboBox();
            SetToTimerComboBox();
            roomListComboBox.IsEnabled = false;
            ResizeMode = ResizeMode.CanMinimize;
        }

        #region submitButton_Click
        /*
         * submitButton_Click():
         * This method carries out data collection from GUI and sends it to WriteDataToExcel().
         */
        private void submitButton_Click(object sender, RoutedEventArgs e)
        {
            bool executionFlag = true;
            DateTimeSlot dts = new DateTimeSlot();
            //DatePicker Validations
            if (!string.IsNullOrEmpty(dateDatePicker.SelectedDate.ToString()))
            {
                string[] dateInput = dateDatePicker.SelectedDate.ToString().Split(' '); //Get data from DatePicker
                dts.Date = dateInput[0];
            }
            else
            {
                MessageBox.Show("Please select Date!",
                                "Error", MessageBoxButton.OK,
                                MessageBoxImage.Error);
                executionFlag = false;
            }

            //TimePicker Validations
            if (fromTimeComboBox.SelectedIndex == -1 || toTimeComboBox.SelectedIndex == -1)
            {
                MessageBox.Show("Please select Time!",
                                "Error", MessageBoxButton.OK,
                                MessageBoxImage.Error);
                executionFlag = false;
            }

            if(this.fromTimeComboBox.Items.IndexOf(fromTimeComboBox.SelectedItem) > this.fromTimeComboBox.Items.IndexOf(toTimeComboBox.SelectedItem))
            {
                MessageBox.Show("Start time cannot be after End time!",
                                "Error", MessageBoxButton.OK,
                                MessageBoxImage.Error);
                executionFlag = false;
            }

            if (executionFlag)
            {
                dts.From = this.fromTimeComboBox.Items.IndexOf(fromTimeComboBox.SelectedItem) + 1;
                dts.To = this.fromTimeComboBox.Items.IndexOf(toTimeComboBox.SelectedItem) + 1;

                dts.MeetingRoomList = ExcelHandler.GetDataFromExcel(dts);

                //binding list to dropdown
                if (dts.MeetingRoomList.Count > 0)
                {
                    foreach (var val in dts.MeetingRoomList)
                    {
                        roomListComboBox.Items.Add(val);
                    }
                    roomListComboBox.SelectedIndex = 0;
                    roomListComboBox.IsEnabled = true;
                }
                else
                {
                    MessageBox.Show("No Meeting Room is free for the selected dates and time!",
                                    "Critical Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }

            ExcelHandler.WriteDataToExcel();
        }
        #endregion

        #region ResetButton_Click
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

        #region SetFromTimerComboBox
        /*
         * Sets the from timer drop down
         */
        private void SetFromTimerComboBox()
        {
            this.fromTimeComboBox.Items.Add("0000");
            this.fromTimeComboBox.Items.Add("0030");
            this.fromTimeComboBox.Items.Add("0100");
            this.fromTimeComboBox.Items.Add("0130");
            this.fromTimeComboBox.Items.Add("0200");
            this.fromTimeComboBox.Items.Add("0230");
            this.fromTimeComboBox.Items.Add("0300");
            this.fromTimeComboBox.Items.Add("0330");
            this.fromTimeComboBox.Items.Add("0400");
            this.fromTimeComboBox.Items.Add("0430");
            this.fromTimeComboBox.Items.Add("0500");
            this.fromTimeComboBox.Items.Add("0530");
            this.fromTimeComboBox.Items.Add("0600");
            this.fromTimeComboBox.Items.Add("0630");
            this.fromTimeComboBox.Items.Add("0700");
            this.fromTimeComboBox.Items.Add("0730");
            this.fromTimeComboBox.Items.Add("0800");
            this.fromTimeComboBox.Items.Add("0830");
            this.fromTimeComboBox.Items.Add("0900");
            this.fromTimeComboBox.Items.Add("0930");
            this.fromTimeComboBox.Items.Add("1000");
            this.fromTimeComboBox.Items.Add("1030");
            this.fromTimeComboBox.Items.Add("1100");
            this.fromTimeComboBox.Items.Add("1130");
            this.fromTimeComboBox.Items.Add("1200");
            this.fromTimeComboBox.Items.Add("1230");
            this.fromTimeComboBox.Items.Add("1300");
            this.fromTimeComboBox.Items.Add("1330");
            this.fromTimeComboBox.Items.Add("1400");
            this.fromTimeComboBox.Items.Add("1430");
            this.fromTimeComboBox.Items.Add("1500");
            this.fromTimeComboBox.Items.Add("1530");
            this.fromTimeComboBox.Items.Add("1600");
            this.fromTimeComboBox.Items.Add("1630");
            this.fromTimeComboBox.Items.Add("1700");
            this.fromTimeComboBox.Items.Add("1730");
            this.fromTimeComboBox.Items.Add("1800");
            this.fromTimeComboBox.Items.Add("1830");
            this.fromTimeComboBox.Items.Add("1900");
            this.fromTimeComboBox.Items.Add("1930");
            this.fromTimeComboBox.Items.Add("2000");
            this.fromTimeComboBox.Items.Add("2030");
            this.fromTimeComboBox.Items.Add("2100");
            this.fromTimeComboBox.Items.Add("2130");
            this.fromTimeComboBox.Items.Add("2200");
            this.fromTimeComboBox.Items.Add("2230");
            this.fromTimeComboBox.Items.Add("2300");
            this.fromTimeComboBox.Items.Add("2330");
        }
        #endregion

        #region SetToTimerComboBox
        /*
         * Sets the from timer drop down
         */
        private void SetToTimerComboBox()
        {
            this.toTimeComboBox.Items.Add("0000");
            this.toTimeComboBox.Items.Add("0030");
            this.toTimeComboBox.Items.Add("0100");
            this.toTimeComboBox.Items.Add("0130");
            this.toTimeComboBox.Items.Add("0200");
            this.toTimeComboBox.Items.Add("0230");
            this.toTimeComboBox.Items.Add("0300");
            this.toTimeComboBox.Items.Add("0330");
            this.toTimeComboBox.Items.Add("0400");
            this.toTimeComboBox.Items.Add("0430");
            this.toTimeComboBox.Items.Add("0500");
            this.toTimeComboBox.Items.Add("0530");
            this.toTimeComboBox.Items.Add("0600");
            this.toTimeComboBox.Items.Add("0630");
            this.toTimeComboBox.Items.Add("0700");
            this.toTimeComboBox.Items.Add("0730");
            this.toTimeComboBox.Items.Add("0800");
            this.toTimeComboBox.Items.Add("0830");
            this.toTimeComboBox.Items.Add("0900");
            this.toTimeComboBox.Items.Add("0930");
            this.toTimeComboBox.Items.Add("1000");
            this.toTimeComboBox.Items.Add("1030");
            this.toTimeComboBox.Items.Add("1100");
            this.toTimeComboBox.Items.Add("1130");
            this.toTimeComboBox.Items.Add("1200");
            this.toTimeComboBox.Items.Add("1230");
            this.toTimeComboBox.Items.Add("1300");
            this.toTimeComboBox.Items.Add("1330");
            this.toTimeComboBox.Items.Add("1400");
            this.toTimeComboBox.Items.Add("1430");
            this.toTimeComboBox.Items.Add("1500");
            this.toTimeComboBox.Items.Add("1530");
            this.toTimeComboBox.Items.Add("1600");
            this.toTimeComboBox.Items.Add("1630");
            this.toTimeComboBox.Items.Add("1700");
            this.toTimeComboBox.Items.Add("1730");
            this.toTimeComboBox.Items.Add("1800");
            this.toTimeComboBox.Items.Add("1830");
            this.toTimeComboBox.Items.Add("1900");
            this.toTimeComboBox.Items.Add("1930");
            this.toTimeComboBox.Items.Add("2000");
            this.toTimeComboBox.Items.Add("2030");
            this.toTimeComboBox.Items.Add("2100");
            this.toTimeComboBox.Items.Add("2130");
            this.toTimeComboBox.Items.Add("2200");
            this.toTimeComboBox.Items.Add("2230");
            this.toTimeComboBox.Items.Add("2300");
            this.toTimeComboBox.Items.Add("2330");
        }
        #endregion
    }
}