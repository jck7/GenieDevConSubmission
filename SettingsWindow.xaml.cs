using System;
using System.Windows;
using System.Diagnostics;

namespace ExcelGenie
{
    public partial class SettingsWindow : Window
    {
        public string Email { get; private set; }

        public SettingsWindow(string currentEmail, string subscriptionStatus)
        {
            InitializeComponent();
            EmailTextBox.Text = currentEmail;
            SubscriptionStatusText.Text = subscriptionStatus;
            Email = currentEmail;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Email = EmailTextBox.Text;
            DialogResult = true;
            Close();
        }

        private void ContactSupport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo("http://genieforapps.com/support") { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to open the URL: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ManageSubscription_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo("https://www.genieforapps.com/manage") { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to open the URL: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}