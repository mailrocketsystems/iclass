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
using System.ComponentModel;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;


namespace iClass
{
    /// <summary>
    /// Interaction logic for DeleteStudentWindow.xaml
    /// </summary>
    public partial class DeleteStudentWindow : System.Windows.Window
    {
        Array className;
        int parse; int row = 0;
        public DeleteStudentWindow(Array data)
        {
            InitializeComponent();

            className = data;
            selectClassComboBox.Items.Clear();
            selectClassComboBox.ItemsSource = className;
        }

        string classData = null;
        
        
        private void ProceedClass_ButtonClick(object sender, RoutedEventArgs e)
        {
            selectStudentLabel.IsEnabled = true;
            ProceedStudentButton.IsEnabled = true;
            selectStudentComboBox.IsEnabled = true;
            //selectStudentComboBox.Items.Clear();

            classData = selectClassComboBox.Text;
            if (string.IsNullOrWhiteSpace(classData))
            {
                Log("Error  :   Class not selected");
                System.Media.SystemSounds.Hand.Play();
                MessageBox.Show("Please select a class   ", "Error ", MessageBoxButton.OK, MessageBoxImage.Error);

            }
            else
            {
                FileInfo file = new FileInfo(@"C:\\Rocket\\iClass\\Class\\" + classData);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                    ExcelWorksheet worksheet = excelWorkBook.Worksheets.First();
                    
                    String studentNames = null;
                    int i =0;
                    for (parse = 2; parse <= 100; parse++)
                    {
                        if (worksheet.Cells[parse, 2].Value != null)
                        {
                            studentNames += Convert.ToString(worksheet.Cells[parse, 2].Value);
                            studentNames += ",";
                            
                        }
                        else
                        {
                           break;
                        }

                    }

                    if (studentNames != null)
                    {
                        string[] names = studentNames.Split(',');
                        selectStudentComboBox.ItemsSource = names;
                        excelPackage.Save();
                    }
                    else
                    {
                        MessageBox.Show("The class you have selected is empty", "Error", MessageBoxButton.OK, MessageBoxImage.Information);
                        this.Close();
                    }
                    
                    
                                        
                }
            }
        }

        private void ProceedStudent_ButtonClick(object sender, RoutedEventArgs e)
        {
            String student = selectStudentComboBox.Text;
            MessageBoxResult result = MessageBox.Show("Are you sure you want to delete student from the class", "Information", MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
            if (result == MessageBoxResult.Yes)
            {
                FileInfo file = new FileInfo(@"C:\\Rocket\\iClass\\Class\\" + classData);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                    ExcelWorksheet worksheet = excelWorkBook.Worksheets.First();
                    for (parse = 2; parse <= 100; parse++)
                    {
                        if (Convert.ToString(worksheet.Cells[parse, 2].Value) == student)
                        {
                            worksheet.DeleteRow(parse);
                            excelPackage.Save();
                            break;
                        }
                                                
                    }




                }

                classData = classData.Remove(classData.Length - 5);
                FileInfo fileName = new FileInfo(@"C:\\Rocket\\iClass\\Attendance\\" + classData + "_Class_Attendance.xlsx");
                using (ExcelPackage Package = new ExcelPackage(fileName))
                {
                    ExcelWorkbook excelWorkBook = Package.Workbook;
                    ExcelWorksheet worksheet = excelWorkBook.Worksheets.First();
                    for (parse = 2; parse <= 100; parse++)
                    {
                        if (Convert.ToString(worksheet.Cells[parse, 1].Value) == student)
                        {
                            worksheet.DeleteRow(parse);
                            Package.Save();
                            break;
                        }

                    }
                }
                MessageBox.Show("Student has been deleted successfully", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

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
