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
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.IO;
using System.Diagnostics;
namespace iClass
{
    /// <summary>
    /// Interaction logic for UpdateClassPopUp.xaml
    /// </summary>
    public partial class UpdateClassPopUp : Window
    {
        public UpdateClassPopUp()
        {
            InitializeComponent();
        }

        private void AddStudent_ButtonClick(object sender, RoutedEventArgs e)
        {
            addStudentButton.IsEnabled = false;
            Array allClass;
            DirectoryInfo d = new DirectoryInfo(@"C:\\Rocket\\iClass\\Class");
            FileInfo[] Files = d.GetFiles("*.xlsx");
            allClass = Files;
            addStudentButton.IsEnabled = true;
            AddStudent addStudentWindow = new AddStudent(allClass);
            this.Close();
            addStudentWindow.Show();
        }

        private void EditStudent_ButtonClilck(object sender, RoutedEventArgs e)
        {
            editStudentButton.IsEnabled = false;
            Array allClass;
            DirectoryInfo d = new DirectoryInfo(@"C:\\Rocket\\iClass\\Class");
            FileInfo[] Files = d.GetFiles("*.xlsx");
            allClass = Files;
            editStudentButton.IsEnabled = false;
            EditStudentPopUp window = new EditStudentPopUp(allClass);
            this.Close();
            window.Show();
        }

        private void DeleteStudent_ButtonClick(object sender, RoutedEventArgs e)
        {

            Log("Delete Class button clicked     :   ");
            
            Array allClass;
            DirectoryInfo d = new DirectoryInfo(@"C:\\Rocket\\iClass\\Class");
            FileInfo[] Files = d.GetFiles("*.xlsx");
            allClass = Files;
            
            

            DeleteStudentWindow DeleteStudentWindow = new DeleteStudentWindow(allClass);
            this.Close();
            DeleteStudentWindow.Show();
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
