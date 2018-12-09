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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;


namespace iClass
{
    /// <summary>
    /// Interaction logic for CreateClassPopUp.xaml
    /// </summary>
    public partial class CreateClassPopUp : System.Windows.Window
    {
        string[] teacher = new string[100];
        string className, numberOfStudents, teacherName;
        public CreateClassPopUp(string[] data)
        {
            InitializeComponent();
            //Log("Create class pop up active now");
            teacher = data;
            selectTeacherComboBox.ItemsSource = null;
            selectTeacherComboBox.ItemsSource = teacher;
        }

        private void ProceedCreateClass_ButtonClick(object sender, RoutedEventArgs e)
        {
            Log("Proceed Create class button clicked");
            ProceedCreateClassButton.IsEnabled = false;
            className = classNameTextBox.Text;
            numberOfStudents = numberOfStudentsComboBox.Text;
            teacherName = selectTeacherComboBox.Text;
            
            if (string.IsNullOrWhiteSpace(classNameTextBox.Text) || (numberOfStudentsComboBox.SelectedItem == null) || (selectTeacherComboBox.SelectedItem == null))
            {
                Log("Error  :   Fill all the details");
                System.Media.SystemSounds.Hand.Play();
                MessageBox.Show("Please fill all the details   ", "Error ", MessageBoxButton.OK, MessageBoxImage.Error);
                ProceedCreateClassButton.IsEnabled = true;
                this.Close();
            }
            else
            {

                if (File.Exists("C:\\Rocket\\iClass\\Class\\" + className + ".xlsx") == true)
                {
                    System.Media.SystemSounds.Hand.Play();
                    MessageBoxResult res = MessageBox.Show("The class you are trying to create already exist             \n" +
                                                        "Please create another class                      ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    Log("Error  :   Class already exists");
                    ProceedCreateClassButton.IsEnabled = true;
                    if (res == MessageBoxResult.OK)
                    {
                        this.Close();

                    }
                }
                else
                {
                    Log("Create Class window called");
                    CreateClassWindow createClassWindow = new CreateClassWindow(className, numberOfStudents, teacherName);
                    createClassWindow.Show();
                    ProceedCreateClassButton.IsEnabled = true;
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
