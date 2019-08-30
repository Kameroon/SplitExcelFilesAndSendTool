using System;
using System.ComponentModel;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Windows;


namespace SplitExcelFiles
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        #region - PropertyChanged Notification -
        public event PropertyChangedEventHandler PropertyChanged;

        public void NotifyPropertyChanged(string property)
        {
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(property));
        }
        #endregion 

        // --  --
        public ViewModelBase() { }

        /// <summary>
        /// -- --
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public string RemoveSpecialCharacters(string str)
        {
            return Regex.Replace(str, "[^a-zA-Z0-9_]+", "_", RegexOptions.Compiled);
        }

        private static Regex digitsOnly = new Regex(@"[^\d]");
        public static string CleanValue(string val)
        {
            return digitsOnly.Replace(val, "");
        }

        #region -- Manege messageBox --
        /// <summary>
        /// - Display quetions messages -
        /// </summary>
        /// <param name="message"></param>
        public void DisplayInfoMessage(string message)
        {
            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                if (message != null || !string.IsNullOrWhiteSpace(message))
                    MessageBox.Show(message, " MVVM Application ", MessageBoxButton.OK, MessageBoxImage.Information);
            }));
        }

        /// <summary>
        /// - Display quetions messages -
        /// </summary>
        /// <param name="message"></param>
        public void DisplayQuestionMessage(string message)
        {
            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                if (message != null || !string.IsNullOrWhiteSpace(message))
                    MessageBox.Show(message, " MVVM Application ", MessageBoxButton.YesNo, MessageBoxImage.Question);
            }));
        }

        /// <summary>
        /// - Display error messages -
        /// </summary>
        /// <param name="message"></param>
        public void DisplayErrorMessage(string message)
        {
            if (message != null || !string.IsNullOrWhiteSpace(message))
            {
                Application.Current.Dispatcher.Invoke(new Action(() =>
                {
                    MessageBox.Show(Application.Current.MainWindow, message, " MVVM Application ", MessageBoxButton.OK, MessageBoxImage.Error);
                }));
            }
        }

        /// <summary>
        /// - Display cool messages -
        /// </summary>
        /// <param name="message"></param>
        public void DisplayMessage(string message)
        {
            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                if (message != null || !string.IsNullOrWhiteSpace(message))
                    MessageBox.Show(message, " MVVM Application ", MessageBoxButton.OK, MessageBoxImage.Information);
            }));
        }
        #endregion

        #region --  --
        // Function to test for Positive Integers
        public bool IsNaturalNumber(String strNumber)
        {
            Regex objNotNaturalPattern = new Regex("[^0-9]");
            Regex objNaturalPattern = new Regex("0*[1-9][0-9]*");
            return !objNotNaturalPattern.IsMatch(strNumber) &&
            objNaturalPattern.IsMatch(strNumber);
        }

        // Function to test for Positive Integers with zero inclusive
        public bool IsWholeNumber(String strNumber)
        {
            Regex objNotWholePattern = new Regex("[^0-9]");
            return !objNotWholePattern.IsMatch(strNumber);
        }

        // Function to Test for Integers both Positive & Negative
        public bool IsInteger(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9-]");
            Regex objIntPattern = new Regex("^-[0-9]+$|^[0-9]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }

        // Function to Test for Positive Number both Integer & Real
        public bool IsPositiveNumber(String strNumber)
        {
            Regex objNotPositivePattern = new Regex("[^0-9.]");
            Regex objPositivePattern = new Regex(
            "^[.][0-9]+$|[0-9]*[.]*[0-9]+$");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            return !objNotPositivePattern.IsMatch(strNumber) &&
            objPositivePattern.IsMatch(strNumber) &&
            !objTwoDotPattern.IsMatch(strNumber);
        }

        // Function to test whether the string is valid number or not
        public bool IsNumber(String strNumber)
        {
            Regex objNotNumberPattern = new Regex("[^0-9.-]");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            Regex objTwoMinusPattern = new Regex("[0-9]*[-][0-9]*[-][0-9]*");
            String strValidRealPattern = "^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$";
            String strValidIntegerPattern = "^([-]|[0-9])[0-9]*$";
            Regex objNumberPattern = new Regex("(" + strValidRealPattern + ")|(" + strValidIntegerPattern + ")");
            return !objNotNumberPattern.IsMatch(strNumber) &&
                   !objTwoDotPattern.IsMatch(strNumber) &&
                   !objTwoMinusPattern.IsMatch(strNumber) &&
                    objNumberPattern.IsMatch(strNumber);
        }

        // Function To test for Alphabets.
        public bool IsAlpha(String strToCheck)
        {
            Regex objAlphaPattern = new Regex("[^a-zA-Z]");
            return !objAlphaPattern.IsMatch(strToCheck);
        }

        // Function to Check for AlphaNumeric.
        public bool IsAlphaNumeric(String strToCheck)
        {
            Regex objAlphaNumericPattern = new Regex("[^a-zA-Z0-9]");
            return !objAlphaNumericPattern.IsMatch(strToCheck);
        }

        public void SplitNumberInWord(string str)
        {
            // -- remplace 'input' by 'str' --
            const string input = "There are 4 numbers in this string: 40, 30, and 10.";
            // Split on one or more non-digit characters.
            string[] numbers = System.Text.RegularExpressions.Regex.Split(input, @"\D+");
            foreach (string value in numbers)
            {
                if (!string.IsNullOrEmpty(value))
                {
                    int i = int.Parse(value);
                    Console.WriteLine("Number: {0}", i);
                }
            }
        }

        public bool IsNumeric(string value)
        {
            int parsedValue;
            if (!int.TryParse(value, out parsedValue))
            {
                return false;
            }
            else
                return true;
        }


        #region - Preparing mail mail -
        /// <summary>
        /// -  -
        /// </summary>
        /// <param name="emailaddress"></param>
        /// <returns></returns>
        private bool IsValid(string emailaddress)
        {
            try
            {
                MailAddress m = new MailAddress(emailaddress);

                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        /// <summary>
        /// - Vérifie l'adresse email -
        /// </summary>
        /// <param name="mailAd"></param>
        private void CheckEmailFormat(string mailAd)
        {
            string pattern = "^([0-9a-zA-Z]([-\\.\\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\\w]*[0-9a-zA-Z]\\.)+[a-zA-Z]{2,9})$";

            if (!Regex.IsMatch(mailAd, pattern))
                MessageBox.Show("L'adresse email : " + mailAd + " ne respecte pas le format requis ",
                    " XLSXDécoupeFiles ", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        #endregion
        #endregion


    }

    // ---  *********  ---
    public static class FocusExtension
    {
        public static readonly DependencyProperty IsFocusedProperty =
            DependencyProperty.RegisterAttached("IsFocused", typeof(bool?), typeof(FocusExtension), new FrameworkPropertyMetadata(IsFocusedChanged));

        public static bool? GetIsFocused(DependencyObject element)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }

            return (bool?)element.GetValue(IsFocusedProperty);
        }

        public static void SetIsFocused(DependencyObject element, bool? value)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }

            element.SetValue(IsFocusedProperty, value);
        }

        private static void IsFocusedChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var fe = (FrameworkElement)d;

            if (e.OldValue == null)
            {
                fe.GotFocus += FrameworkElement_GotFocus;
                fe.LostFocus += FrameworkElement_LostFocus;
            }

            if (!fe.IsVisible)
            {
                fe.IsVisibleChanged += new DependencyPropertyChangedEventHandler(fe_IsVisibleChanged);
            }

            if ((bool)e.NewValue)
            {
                fe.Focus();
            }
        }

        private static void fe_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var fe = (FrameworkElement)sender;
            if (fe.IsVisible && (bool)((FrameworkElement)sender).GetValue(IsFocusedProperty))
            {
                fe.IsVisibleChanged -= fe_IsVisibleChanged;
                fe.Focus();
            }
        }

        private static void FrameworkElement_GotFocus(object sender, RoutedEventArgs e)
        {
            ((FrameworkElement)sender).SetValue(IsFocusedProperty, true);
        }

        private static void FrameworkElement_LostFocus(object sender, RoutedEventArgs e)
        {
            ((FrameworkElement)sender).SetValue(IsFocusedProperty, false);
        }
    }
}
