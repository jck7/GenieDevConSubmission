using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace ExcelGenie
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Remove this block as we'll set the icon in XAML
            // this.Resources["ApplicationIcon"] = new System.Windows.Media.Imaging.BitmapImage(new System.Uri("pack://application:,,,/ExcelGenie/Assets/ExcelGenie.ico"));
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            ShowExceptionDetails(e.Exception, "Dispatcher Unhandled Exception");
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                ShowExceptionDetails(ex, "Unhandled Exception");
            }
            else
            {
                ShowExceptionDetails(new Exception("An unknown error occurred."), "Unhandled Exception");
            }
        }

        private void ShowExceptionDetails(Exception ex, string title)
        {
            string errorMessage = $"An error occurred: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}";
            MessageBox.Show(errorMessage, title, MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
