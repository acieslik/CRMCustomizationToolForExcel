using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Media.Imaging;
using System.Globalization;
using System.Windows.Data;
using System.Reflection;
using System.IO;
using System.Resources;
using System.Collections;

namespace DynamicsCRMCustomizationToolForExcel.AddIn.Components
{
    public class BoolToImageConverter : IValueConverter
    {

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new System.NotImplementedException();
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {

            if (value != null)
            {

                bool status = (bool)value;
                if (status == false)
                {
                    BitmapImage source = new BitmapImage();
                    source.BeginInit();
                    source.UriSource = new Uri(@"pack://application:,,,/DynamicsCRMCustomizationToolForExcel.AddIn;component/Resources/Error.gif");
                    source.EndInit();
                    return source;
                }
                else
                {
                    BitmapImage source = new BitmapImage();
                    source.BeginInit();
                    source.UriSource = new Uri(@"pack://application:,,,/DynamicsCRMCustomizationToolForExcel.AddIn;component/Resources/OK.png");
                    source.EndInit();
                    return source;
                }
            }

            return null;
        }


    }
}
