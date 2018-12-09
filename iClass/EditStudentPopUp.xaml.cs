using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;


namespace iClass
{
    /// <summary>
    /// Interaction logic for EditStudentPopUp.xaml
    /// </summary>
    public partial class EditStudentPopUp : Window
    {
        public EditStudentPopUp(Array data)
        {
            InitializeComponent();
            selectClassComboBox.Items.Clear();
            selectClassComboBox.ItemsSource = data;
        }

        private void ProceedButton_Click(object sender, RoutedEventArgs e)
        {
            proceedClassButton.IsEnabled = false;
            string className = selectClassComboBox.Text;
            if (string.IsNullOrWhiteSpace(selectClassComboBox.Text))
            {
                MessageBox.Show("Please select a class ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                proceedClassButton.IsEnabled = true;
            }
            else
            {
                EditClassWindow editClassWindow = new EditClassWindow(className);
                proceedClassButton.IsEnabled = true;
                editClassWindow.Show();
                this.Close();

            }
        }

        void Log(string data)
        {
            try
            {
                string path = @"C:\\Rocket\\iClass\\Logs\\" + DateTime.Now.ToString("dd-M-yyyy") + ".txt";

                System.IO.Directory.CreateDirectory(@"C:\Rocket\iClass\Logs\");

                if (File.Exists(path))
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine(data + "   :   " + DateTime.Now.ToString("dd-M-yyyy-hh-mm-ss"));
                        sw.Close();
                    }
                }
                else
                {
                    StreamWriter myFile = new StreamWriter(path);
                    myFile.WriteLine(data + "   :   " + DateTime.Now.ToString("dd-M-yyyy-hh-mm-ss"));
                    myFile.Close();



                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(Convert.ToString(exception), "Exception Occured during log");
            }
        }
    }
}
