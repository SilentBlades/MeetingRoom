﻿/*************************************************************************************

   Extended WPF Toolkit

   Copyright (C) 2007-2013 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features, controls, and fast professional support,
   pick up the Plus Edition at http://xceed.com/wpf_toolkit

   Stay informed: follow @datagrid on Twitter or Like http://facebook.com/datagrids

  ***********************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Data;
using System.Windows;

namespace Xceed.Wpf.AvalonDock.Converters
{
    [ValueConversion(typeof(bool), typeof(Visibility))] 
    public class InverseBoolToVisibilityConverter : IValueConverter 
    { 

        #region IValueConverter Members 
        /// <summary> 
        /// Converts a value. 
        /// </summary> 
        /// <param name="value">The value produced by the binding source.</param> 
        /// <param name="targetType">The type of the binding target property.</param> 
        /// <param name="parameter">The converter parameter to use.</param> 
        /// <param name="culture">The culture to use in the converter.</param> 
        /// <returns> 
        /// A converted value. If the method returns null, the valid null value is used. 
        /// </returns> 
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) 
        { 
                if (value is bool && targetType == typeof(Visibility)) 
                { 
                        bool val = !(bool)value; 
                        if (val) 
                                return Visibility.Visible; 
                        else 
                                if (parameter != null && parameter is Visibility ) 
                                        return parameter; 
                                else 
                                        return Visibility.Collapsed; 
                } 
                throw new ArgumentException("Invalid argument/return type. Expected argument: bool and return type: Visibility"); 
        } 

        /// <summary> 
        /// Converts a value. 
        /// </summary> 
        /// <param name="value">The value that is produced by the binding target.</param> 
        /// <param name="targetType">The type to convert to.</param> 
        /// <param name="parameter">The converter parameter to use.</param> 
        /// <param name="culture">The culture to use in the converter.</param> 
        /// <returns> 
        /// A converted value. If the method returns null, the valid null value is used. 
        /// </returns> 
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) 
        { 
                if (value is Visibility && targetType == typeof(bool)) 
                { 
                        Visibility val = (Visibility)value; 
                        if (val == Visibility.Visible) 
                                return false; 
                        else 
                                return true; 
                } 
                throw new ArgumentException("Invalid argument/return type. Expected argument: Visibility and return type: bool"); 
        } 
        #endregion 
    } 


}
