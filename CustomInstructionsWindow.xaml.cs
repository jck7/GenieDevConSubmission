using System.Windows;
using System.Windows.Input;

namespace ExcelGenie
{
    public partial class CustomInstructionsWindow : Window
    {
        public bool isDarkTheme { get; private set; }
        public string CustomInstructions { get; set; } = string.Empty;

        public CustomInstructionsWindow(bool isDarkTheme)
        {
            InitializeComponent();
            this.isDarkTheme = isDarkTheme;
            DataContext = this;

            // Load existing custom instructions if any
            LoadCustomInstructions();
        }

        private void LoadCustomInstructions()
        {
            string defaultInstructions = "Enter your custom instructions here. These instructions will be used to guide the AI in handling your Excel tasks.\n\n" +
                                       "Examples:\n" +
                                       "- Always use consistent formatting across similar data\n" +
                                       "- Prefer named ranges for better formula readability\n" +
                                       "- Add comments to complex formulas\n" +
                                       "- Use specific number formats for financial data";

            CustomInstructions = Properties.Settings.Default.CustomInstructions;
            InstructionsTextBox.Text = string.IsNullOrEmpty(CustomInstructions) ? defaultInstructions : CustomInstructions;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            CustomInstructions = InstructionsTextBox.Text;
            Properties.Settings.Default.CustomInstructions = CustomInstructions;
            Properties.Settings.Default.Save();
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
    }
} 