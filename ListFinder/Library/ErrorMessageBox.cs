using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ListFinder.Libraries
{
    public class ErrorMessageBox
    {
        ErrorMessageBox()
        {
        }

        public static MessageBoxResult Show(string messageBoxText)
        {
            return MessageBox.Show(messageBoxText, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
