using System;
using System.Collections.Generic;
using System.IO;
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

namespace iClass
{
    /// <summary>
    /// Interaction logic for DeleteClassPopUp.xaml
    /// </summary>
    
    public partial class DeleteClassPopUp : Window
    {
        Array className;
        public DeleteClassPopUp(Array data)
        {
            InitializeComponent();
            System.Media.SystemSounds.Exclamation.Play();
            className = data;


            selectClassComboBox.Items.Clear();
            selectClassComboBox.ItemsSource = className;

        }
        string classData = null;
        private void DeleteClass_ButtonClick(object sender, RoutedEventArgs e)
        {
            DeleteClassButton.IsEnabled = false;
            classData = selectClassComboBox.Text;
            if (string.IsNullOrWhiteSpace(classData))
            {
                System.Media.SystemSounds.Hand.Play();
                MessageBox.Show("Please select a class   ", "Error ", MessageBoxButton.OK, MessageBoxImage.Error);
                DeleteClassButton.IsEnabled = true;
                //this.Close();
            }
            else
            {
                classData = classData.Remove(classData.Length - 5);
                MessageBoxResult result = MessageBox.Show("Do you want to take backup of the class before deleting it.", "Information", MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                if (result == MessageBoxResult.Yes)
                {
                    System.IO.File.Move("C:\\Rocket\\iClass\\Class\\" + classData + ".xlsx", "C:\\Rocket\\iClass\\Backup\\" + classData + "@"+ DateTime.Now.ToString("dd.MM.yyyy_hh.mm.ss") + ".xlsx");
                    System.IO.File.Move("C:\\Rocket\\iClass\\Attendance\\" + classData + "_Class_Attendance.xlsx", "C:\\Rocket\\iClass\\Backup\\" + classData + "_Class_Attendance" + "@" + DateTime.Now.ToString("dd.MM.yyyy_hh.mm.ss") + ".xlsx");

                    File.Delete("C:\\Rocket\\iClass\\Class\\" + classData + ".xlsx");
                    File.Delete("C:\\Rocket\\iClass\\Attendance\\" + classData + "_Class_Attendance.xlsx");

                    MessageBox.Show("Class has been deleted and a copy has been saved in backup folder", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    this.Close();
                }
                else
                {
                    File.Delete("C:\\Rocket\\iClass\\Class\\" + classData + ".xlsx");
                    File.Delete("C:\\Rocket\\iClass\\Attendance\\" + classData + "_Class_Attendance.xlsx");

                    MessageBox.Show("Class has been deleted successfully", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    this.Close();
                }

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
