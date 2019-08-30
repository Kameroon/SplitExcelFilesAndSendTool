 using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;


namespace SplitExcelFiles
{
    #region -- Gère la visibilité de boutom basé sur 2 ou plusieurs textbox --
    public class FullNameConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (values != null)
            {
                return values[0] + " " + values[1];
            }
            return " ";
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            string[] values = null;
            if (value != null)
                return values = value.ToString().Split(' ');
            return values;
        }
    }


    public class MultivalueConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool res = false;
            if (values.Count() >= 2)
            {
                if (string.IsNullOrEmpty(values[0].ToString()) ||
                    string.IsNullOrEmpty(values[1].ToString()) ||
                    string.IsNullOrEmpty(values[2].ToString()))
                    return false;
                else
                    return true;

                return res;
                List<string> fields = values.Select(i => i.ToString()).ToList();
            }
            else
                return false;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class MultivalueConverter2 : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool res = false;
            if (values.Count() >= 2)
            {
                if (string.IsNullOrEmpty(values[0].ToString()) ||
                    string.IsNullOrEmpty(values[1].ToString()) ||
                    string.IsNullOrEmpty(values[2].ToString()) ||
                    string.IsNullOrEmpty(values[3].ToString()))
                    return false;
                else
                    return true;

                return res;
                List<string> fields = values.Select(i => i.ToString()).ToList();
            }
            else
                return res;
            //return false;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    #endregion

    #region -- Gère la visibilité d'un combobox paraport au précédent  --
    public class VisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            //if (value != null && value.ToString().Equals("OTHERS"))
            if (value != null)
            {
                return Visibility.Visible;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    #endregion

    #region -- Gère l'affichage du soustype apres la saisi du soustypeID --
    public class SubTypeConverterVisibility : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {

            if (!string.IsNullOrWhiteSpace(value.ToString()) && Regex.IsMatch(value.ToString(), @"^\d+$") && value.ToString().Length >= 4)
            {
                return Visibility.Visible;
            }
            else
            {
                return Visibility.Collapsed;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    #endregion

    #region -- buton question  --
    public class BooleanToCollapsedVisibilityConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            //reverse conversion (false=>Visible, true=>collapsed) on any given parameter
            bool input = (null == parameter) ? (bool)value : !((bool)value);
            return (input) ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }

    //[ValueConversion(typeof(string), typeof(Visibility))]
    public class StringToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (string.IsNullOrEmpty((string)value))
            {
                return Visibility.Collapsed;
            }
            else
            {
                return Visibility.Visible;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    #endregion

    #region -- ComboBoxSelectedItemConverter --
    public class ComboBoxSelectedItemConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is ComboBoxItem)
            {
                return true;
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
    #endregion

    #region -- Combobox item count !!!!!!!!!!!!!!!!! --
    public class ComboBoxItemCountToEnabledConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && value.GetType() == typeof(Int32))
            {
                if ((int)value > 1)
                    return true;
            }

            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class ComboBoxItemCountToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && value.GetType() == typeof(Int32))
            {
                if ((int)value > 1)
                    return Visibility.Visible;
            }

            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    #endregion

    #region --  --
    public class MyConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            return ((bool)values[0] && (bool)values[1]) ? Visibility.Visible : Visibility.Collapsed;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            //return Binding.DoNothing;
            throw new NotImplementedException();
        }
    }

    public class MyBooleanToVisibilityConverter : IValueConverter
    {
        private BooleanToVisibilityConverter _converter = new BooleanToVisibilityConverter();
        private DependencyObject _dummy = new DependencyObject();

        private bool DesignMode
        {
            get
            {
                return DesignerProperties.GetIsInDesignMode(_dummy);
            }
        }

        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (DesignMode)
                return Visibility.Visible;
            else
                return _converter.Convert(value, targetType, parameter, culture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return _converter.ConvertBack(value, targetType, parameter, culture);
        }

        #endregion
    }

    public class BoolToOppositeBoolConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (targetType != typeof(bool))
                throw new InvalidOperationException("The target must be a boolean");

            return !(bool)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        #endregion
    }
    #endregion

    #region -- Set only numeric value --
    public class OnlyDigitsValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            var validationResult = new ValidationResult(true, null);
            if (value != null)
            {
                string val = value.ToString();
                if (!string.IsNullOrEmpty(val))
                {
                    var regex = new Regex("[^0-9.]-"); //regex that matches disallowed text
                    var parsingOk = !regex.IsMatch(val);
                    if (!parsingOk)
                    {
                        validationResult = new ValidationResult(false, "Ceci est une zone de chiffre ! ");
                    }
                }
            }
            return validationResult;
        }
    }

    public class ValueValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            if (!string.IsNullOrWhiteSpace(value.ToString()))
            {
                string result = Regex.Replace(value.ToString(), "[^0-9]", ""); // result = "9134445555"
                if (result.Length >= 4 && int.Parse(result) > 0)
                    return new ValidationResult(true, null);

                return new ValidationResult(true, null);
            }
            return new ValidationResult(false, "Juste des valeurs positives !");
        }
    }


    public class StringValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            var validationResult = new ValidationResult(true, null);
            if (value != null)
            {
                if (!string.IsNullOrEmpty(value.ToString()))
                {
                    value = Regex.Replace(value.ToString(), "[^a-zA-Z0-9_.]+", "");

                    //string s = Regex.Replace("s", "[^0-9A-Za-z]+", ",");
                    //string regExp = "[^0-9A-Za-z]+";
                    //value.ToString().Replace(regExp, "");
                    return new ValidationResult(true, null);
                }
            }
            return validationResult;
        }
    }

    public class AgeValidationRule00 : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            if (!string.IsNullOrWhiteSpace(value.ToString()) && value.ToString().Length >= 4 && IsNumeric(value.ToString()))
            {
                int wert = Convert.ToInt32(value);
                if (wert < 0)
                    return new ValidationResult(false, "Veuiller entrer un nombre valide !");
            }
            return new ValidationResult(true, null);
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
    }

    public class NumberValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            double result = 0.0;
            bool canConvert = double.TryParse(value as string, out result);
            return new ValidationResult(canConvert, "Not a valid double");
        }
    }

    public class DoubleToPersistantStringConverter : IValueConverter
    {
        private string lastConvertBackString;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is double)) return null;

            var stringValue = lastConvertBackString ?? value.ToString();
            lastConvertBackString = null;

            return stringValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is string)) return null;

            double result;
            if (double.TryParse((string)value, out result))
            {
                lastConvertBackString = (string)value;
                return result;
            }

            return null;
        }
    }

    #endregion

    #region --  --
    public class EmailValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            var validationResult = new ValidationResult(true, null);
            if (value != null)
            {
                if (!string.IsNullOrEmpty(value.ToString()))
                {
                    //var regex = new Regex("^[a-zA-Z0-9]{1,20}@[a-zA-Z0-9]{1,20}.[a-zA-Z]{2,3}$");
                    var regex = new Regex("^([0-9a-zA-Z]([-\\.\\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\\w]*[0-9a-zA-Z]\\.)+[a-zA-Z]{2,9})$");
                    var parsingOk = regex.IsMatch(value.ToString());
                    if (!parsingOk)
                    {
                        validationResult = new ValidationResult(false, "Vous devez saisir une adresse email valide !");
                    }
                }
            }
            return validationResult;
        }
    }

    public class MultiCheckedToEnabledConverter : IMultiValueConverter
    {
        #region Implementation of IMultiValueConverter

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values != null)
            {
                return values.OfType<bool>().Any(b => b);
            }
            return false;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            return new object[] { };
        }

        #endregion
    }
    #endregion


    public class ObjectToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value != null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

   
    // *!!!!!!!!!

    public class AddConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType,
               object parameter, System.Globalization.CultureInfo culture)
        {
            int result =
                Int32.Parse((string)values[0]) + Int32.Parse((string)values[1]);
            return result.ToString();
        }
        public object[] ConvertBack(object value, Type[] targetTypes,
               object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotSupportedException("Cannot convert back");
        }
    }


}