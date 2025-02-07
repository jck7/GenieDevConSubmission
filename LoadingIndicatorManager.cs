using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;

namespace ExcelGenie
{
    public class LoadingIndicatorManager
    {
        private readonly MainWindow _mainWindow;
        private bool _isLoading = false;

        public LoadingIndicatorManager(MainWindow mainWindow)
        {
            _mainWindow = mainWindow;
        }

        public void ShowLoadingIndicator()
        {
            _mainWindow.Dispatcher.Invoke(() =>
            {
                // Reset all checks
                _mainWindow.Step1Check.Visibility = Visibility.Collapsed;
                _mainWindow.Step2Check.Visibility = Visibility.Collapsed;
                _mainWindow.Step3Check.Visibility = Visibility.Collapsed;
                _mainWindow.Step4Check.Visibility = Visibility.Collapsed;
                _mainWindow.Step5Check.Visibility = Visibility.Collapsed;
                _mainWindow.Step6Check.Visibility = Visibility.Collapsed;

                _mainWindow.CurrentOperationLabel.Text = "Starting process...";
                _mainWindow.PalantirProcessingOverlay.Visibility = Visibility.Visible;
            });
        }

        public void HideLoadingIndicator()
        {
            _mainWindow.Dispatcher.Invoke(() =>
            {
                _mainWindow.PalantirProcessingOverlay.Visibility = Visibility.Collapsed;
                _isLoading = false;
            });
        }

        public void UpdateLoadingStep(int step, string operation = "")
        {
            _mainWindow.Dispatcher.Invoke(() =>
            {
                switch (step)
                {
                    case 1:
                        _mainWindow.Step1Check.Visibility = Visibility.Visible;
                        _mainWindow.CurrentOperationLabel.Text = operation ?? "Processing Application Data...";
                        break;
                    case 2:
                        _mainWindow.Step1Check.Visibility = Visibility.Visible;
                        _mainWindow.Step2Check.Visibility = Visibility.Visible;
                        _mainWindow.CurrentOperationLabel.Text = operation ?? "Creating Workbook...";
                        break;
                    case 3:
                        _mainWindow.Step2Check.Visibility = Visibility.Visible;
                        _mainWindow.Step3Check.Visibility = Visibility.Visible;
                        _mainWindow.CurrentOperationLabel.Text = operation ?? "Processing Connections...";
                        break;
                    case 4:
                        _mainWindow.Step3Check.Visibility = Visibility.Visible;
                        _mainWindow.Step4Check.Visibility = Visibility.Visible;
                        _mainWindow.CurrentOperationLabel.Text = operation ?? "Processing Worksheets...";
                        break;
                    case 5:
                        _mainWindow.Step4Check.Visibility = Visibility.Visible;
                        _mainWindow.Step5Check.Visibility = Visibility.Visible;
                        _mainWindow.CurrentOperationLabel.Text = operation ?? "Processing Named Ranges...";
                        break;
                    case 6:
                        _mainWindow.Step5Check.Visibility = Visibility.Visible;
                        _mainWindow.Step6Check.Visibility = Visibility.Visible;
                        _mainWindow.CurrentOperationLabel.Text = operation ?? "Processing Ranges and Cells...";
                        break;
                }
            });
        }

        public async Task SimulateLoadingSteps()
        {
            if (_isLoading) return;
            
            _isLoading = true;
            ShowLoadingIndicator();

            try
            {
                // Step 1
                await Task.Delay(1000);
                UpdateLoadingStep(1);

                // Step 2
                await Task.Delay(1000);
                UpdateLoadingStep(2);

                // Step 3
                await Task.Delay(1000);
                UpdateLoadingStep(3);

                // Step 4
                await Task.Delay(1000);
                UpdateLoadingStep(4);

                // Step 5
                await Task.Delay(1000);
                UpdateLoadingStep(5);

                // Step 6
                await Task.Delay(1000);
                UpdateLoadingStep(6);

                // Hide after a short delay
                await Task.Delay(500);
                HideLoadingIndicator();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during loading simulation: {ex.Message}");
                HideLoadingIndicator();
            }
        }

        public void UpdateLoadingText(string text)
        {
            _mainWindow.Dispatcher.Invoke(() =>
            {
                _mainWindow.CurrentOperationLabel.Text = text;
            });
        }

        public void UpdateProgress(double progress)
        {
            _mainWindow.Dispatcher.Invoke(() =>
            {
                _mainWindow.LoadingProgressBar.Value = progress;
            });
        }
    }
} 