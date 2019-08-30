using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;


namespace SplitExcelFiles
{
    #region - Delegate command -
    public class DelegateCommand : ICommand
    {
        private readonly Predicate<object> _canExecute;
        private readonly Action<object> _execute;

        public event EventHandler CanExecuteChanged;

        public DelegateCommand(Action<object> execute) : this(execute, null) { }

        public DelegateCommand(Action<object> execute, Predicate<object> canExecute)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null ? true : _canExecute(parameter);
        }

        public void Execute(object parameter)
        {
            _execute(parameter);
        }

        public void RaiseCanExecuteChanged()
        {
            if (CanExecuteChanged != null)
            {
                CanExecuteChanged(this, EventArgs.Empty);
            }
        }
    }
    #endregion

    #region - MouseLeftButtonDown -
    public static class MouseLeftButtonDown
    {
        #region - Dependency properties -
        public static DependencyProperty CommandProperty = DependencyProperty.RegisterAttached("Command", typeof(ICommand),
            typeof(MouseLeftButtonDown), new UIPropertyMetadata(CommandChanged));

        public static DependencyProperty CommandParameterProperty = DependencyProperty.RegisterAttached("CommandParameter",
            typeof(object), typeof(MouseLeftButtonDown), new UIPropertyMetadata(null));
        #endregion

        public static object GetCommand(DependencyObject target)
        {
            return target.GetValue(CommandProperty);
        }

        public static void SetCommand(DependencyObject target, ICommand value)
        {
            target.SetValue(CommandProperty, value);
        }

        public static void SetCommandParameter(DependencyObject target, object value)
        {
            target.SetValue(CommandParameterProperty, value);
        }

        public static object GetCommandParameter(DependencyObject target)
        {
            return target.GetValue(CommandParameterProperty);
        }

        private static void CommandChanged(DependencyObject target, DependencyPropertyChangedEventArgs e)
        {
            FrameworkElement control = target as FrameworkElement;
            if (control != null)
            {
                if ((e.NewValue != null) && (e.OldValue == null))
                {
                    control.PreviewMouseLeftButtonDown += OnPreviewMouseLeftButtonDown;
                }
                else if ((e.NewValue == null) && (e.OldValue != null))
                {
                    control.PreviewMouseLeftButtonDown -= OnPreviewMouseLeftButtonDown;
                }
            }
        }

        private static void OnPreviewMouseLeftButtonDown(object sender, RoutedEventArgs e)
        {
            FrameworkElement control = sender as FrameworkElement;
            ICommand command = (ICommand)control.GetValue(CommandProperty);
            object commandParameter = control.GetValue(CommandParameterProperty);
            command.Execute(commandParameter);
        }
    }
    #endregion

    #region - MouseLeftButtonUp -
    public static class MouseLeftButtonUp
    {
        #region - Dependency properties -
        public static DependencyProperty CommandProperty = DependencyProperty.RegisterAttached("Command", typeof(ICommand),
            typeof(MouseLeftButtonUp), new UIPropertyMetadata(CommandChanged));

        public static DependencyProperty CommandParameterProperty = DependencyProperty.RegisterAttached("CommandParameter",
            typeof(object), typeof(MouseLeftButtonUp), new UIPropertyMetadata(null));
        #endregion

        public static void SetCommand(DependencyObject target, ICommand value)
        {
            target.SetValue(CommandProperty, value);
        }

        public static void SetCommandParameter(DependencyObject target, object value)
        {
            target.SetValue(CommandParameterProperty, value);
        }

        public static object GetCommandParameter(DependencyObject target)
        {
            return target.GetValue(CommandParameterProperty);
        }

        private static void CommandChanged(DependencyObject target, DependencyPropertyChangedEventArgs e)
        {
            FrameworkElement control = target as FrameworkElement;
            if (control != null)
            {
                if ((e.NewValue != null) && (e.OldValue == null))
                {
                    control.PreviewMouseLeftButtonUp += OnPreviewMouseLeftButtonUp;
                }
                else if ((e.NewValue == null) && (e.OldValue != null))
                {
                    control.PreviewMouseLeftButtonUp -= OnPreviewMouseLeftButtonUp;
                }
            }
        }

        private static void OnPreviewMouseLeftButtonUp(object sender, RoutedEventArgs e)
        {
            FrameworkElement control = sender as FrameworkElement;
            ICommand command = (ICommand)control.GetValue(CommandProperty);
            object commandParameter = control.GetValue(CommandParameterProperty);
            command.Execute(commandParameter);
        }
    }
    #endregion

    #region - MouseDoubleClick -
    public static class MouseDoubleClick
    {
        #region - Dependency properties -
        public static DependencyProperty CommandProperty = DependencyProperty.RegisterAttached("Command", typeof(ICommand),
            typeof(MouseDoubleClick), new UIPropertyMetadata(CommandChanged));

        public static DependencyProperty CommandParameterProperty = DependencyProperty.RegisterAttached("CommandParameter",
            typeof(object), typeof(MouseDoubleClick), new UIPropertyMetadata(null));
        #endregion

        public static void SetCommand(DependencyObject target, ICommand value)
        {
            target.SetValue(CommandProperty, value);
        }

        public static void SetCommandParameter(DependencyObject target, object value)
        {
            target.SetValue(CommandParameterProperty, value);
        }

        public static object GetCommandParameter(DependencyObject target)
        {
            return target.GetValue(CommandParameterProperty);
        }

        private static void CommandChanged(DependencyObject target, DependencyPropertyChangedEventArgs e)
        {
            Control control = target as Control;
            if (control != null)
            {
                if ((e.NewValue != null) && (e.OldValue == null))
                {
                    control.MouseDoubleClick += OnMouseDoubleClick;
                }
                else if ((e.NewValue == null) && (e.OldValue != null))
                {
                    control.MouseDoubleClick -= OnMouseDoubleClick;
                }
            }
        }

        private static void OnMouseDoubleClick(object sender, RoutedEventArgs e)
        {
            Control control = sender as Control;
            ICommand command = (ICommand)control.GetValue(CommandProperty);
            object commandParameter = control.GetValue(CommandParameterProperty);
            command.Execute(commandParameter);
        }
    }
    #endregion

    #region - Key Down -
    public static class PreviewKeyDown
    {
        #region - Dependency properties -
        public static DependencyProperty CommandProperty = DependencyProperty.RegisterAttached("Command", typeof(ICommand),
            typeof(PreviewKeyDown), new UIPropertyMetadata(CommandChanged));
        #endregion

        public static void SetCommand(DependencyObject target, ICommand value)
        {
            target.SetValue(CommandProperty, value);
        }

        private static void CommandChanged(DependencyObject target, DependencyPropertyChangedEventArgs e)
        {
            FrameworkElement control = target as FrameworkElement;
            if (control != null)
            {
                if ((e.NewValue != null) && (e.OldValue == null))
                {
                    control.PreviewKeyDown += OnPreviewKeyDown;
                }
                else if ((e.NewValue == null) && (e.OldValue != null))
                {
                    control.PreviewKeyDown -= OnPreviewKeyDown;
                }
            }
        }

        private static void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            FrameworkElement control = sender as FrameworkElement;
            ICommand command = (ICommand)control.GetValue(CommandProperty);
            command.Execute(e);
        }
    }
    #endregion

    #region -- Manege radio buttom --
    //public class MyRadioButton : RadioButton
    //{
    //    public object RadioValue
    //    {
    //        get { return (object)GetValue(RadioValueProperty); }
    //        set { SetValue(RadioValueProperty, value); }
    //    }

    //    // Using a DependencyProperty as the backing store for RadioValue.
    //    //This enables animation, styling, binding, etc...
    //    public static readonly DependencyProperty RadioValueProperty =
    //        DependencyProperty.Register(
    //            "RadioValue",
    //            typeof(object),
    //            typeof(MyRadioButton),
    //            new UIPropertyMetadata(null));

    //    public object RadioBinding
    //    {
    //        get { return (object)GetValue(RadioBindingProperty); }
    //        set { SetValue(RadioBindingProperty, value); }
    //    }

    //    // Using a DependencyProperty as the backing store for RadioBinding.
    //    //This enables animation, styling, binding, etc...
    //    public static readonly DependencyProperty RadioBindingProperty =
    //        DependencyProperty.Register(
    //            "RadioBinding",
    //            typeof(object),
    //            typeof(MyRadioButton),
    //            new FrameworkPropertyMetadata(
    //                null,
    //                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
    //                OnRadioBindingChanged));

    //    private static void OnRadioBindingChanged(
    //        DependencyObject d,
    //        DependencyPropertyChangedEventArgs e)
    //    {
    //        MyRadioButton rb = (MyRadioButton)d;
    //        if (rb.RadioValue.Equals(e.NewValue))
    //            rb.SetCurrentValue(RadioButton.IsCheckedProperty, true);
    //    }

    //    protected override void OnChecked(RoutedEventArgs e)
    //    {
    //        base.OnChecked(e);
    //        SetCurrentValue(RadioBindingProperty, RadioValue);
    //    }
    //}
    #endregion

    #region -- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! --
    /// <summary>
    /// -- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! --
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class DelegateCommand<T> : System.Windows.Input.ICommand
    {
        private readonly Predicate<T> _canExecute;
        private readonly Action<T> _execute;

        public DelegateCommand(Action<T> execute)
            : this(execute, null)
        {
        }

        public DelegateCommand(Action<T> execute, Predicate<T> canExecute)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            if (_canExecute == null)
                return true;

            return _canExecute((parameter == null) ? default(T) : (T)Convert.ChangeType(parameter, typeof(T)));
        }

        public void Execute(object parameter)
        {
            _execute((parameter == null) ? default(T) : (T)Convert.ChangeType(parameter, typeof(T)));
        }

        public event EventHandler CanExecuteChanged;
        public void RaiseCanExecuteChanged()
        {
            if (CanExecuteChanged != null)
                CanExecuteChanged(this, EventArgs.Empty);
        }
    }


    /// <summary>
    /// -- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! --
    /// </summary>
    public static class MouseCommandBehavior
    {
        #region -- TheCommandToRun --
        /// <summary>
        /// -- The comamnd which should be executed when the mouse is down --
        /// </summary>
        public static readonly DependencyProperty MouseDownCommandProperty =
            DependencyProperty.RegisterAttached("MouseDownCommand",
                typeof(ICommand),
                typeof(MouseCommandBehavior),
                new FrameworkPropertyMetadata(null, (obj, e) => OnMouseCommandChanged(obj, (ICommand)e.NewValue, "MouseDown")));

        /// <summary>
        /// -- Gets the MouseDownCommand property --
        /// </summary>
        public static ICommand GetMouseDownCommand(DependencyObject d)
        {
            return (ICommand)d.GetValue(MouseDownCommandProperty);
        }

        /// <summary>
        /// -- Sets the MouseDownCommand property --
        /// </summary>
        public static void SetMouseDownCommand(DependencyObject d, ICommand value)
        {
            d.SetValue(MouseDownCommandProperty, value);
        }

        /// <summary>
        /// -- The comamnd which should be executed when the mouse is up --
        /// </summary>
        public static readonly DependencyProperty MouseUpCommandProperty =
            DependencyProperty.RegisterAttached("MouseUpCommand",
                typeof(ICommand),
                typeof(MouseCommandBehavior),
                new FrameworkPropertyMetadata(null, new PropertyChangedCallback((obj, e) => OnMouseCommandChanged(obj, (ICommand)e.NewValue, "MouseUp"))));

        /// <summary>
        /// -- Gets the MouseUpCommand property --
        /// </summary>
        public static ICommand GetMouseUpCommand(DependencyObject d)
        {
            return (ICommand)d.GetValue(MouseUpCommandProperty);
        }

        /// <summary>
        /// -- Sets the MouseUpCommand property --
        /// </summary>
        public static void SetMouseUpCommand(DependencyObject d, ICommand value)
        {
            d.SetValue(MouseUpCommandProperty, value);
        }

        #endregion

        /// <summary>
        /// -- Registeres the event and calls the command when it gets fired --
        /// </summary>
        private static void OnMouseCommandChanged(DependencyObject d, ICommand command, string routedEventName)
        {
            if (String.IsNullOrEmpty(routedEventName) || command == null) return;

            var element = (FrameworkElement)d;
            switch (routedEventName)
            {
                case "MouseDown":
                    element.PreviewMouseDown += (obj, e) => command.Execute(null);
                    break;
                case "MouseUp":
                    element.PreviewMouseUp += (obj, e) => command.Execute(null);
                    break;
            }
        }
    }
    #endregion

}
