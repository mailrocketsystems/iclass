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
using System.Windows.Threading;
using System.Threading;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Diagnostics;

namespace iClass
{
    /// <summary>
    /// Interaction logic for AddStudent.xaml
    /// </summary>
    public partial class AddStudent : System.Windows.Window
    {
        CircularProgressBar progress = new CircularProgressBar();
        string className;
        int Sno = 0;
        int Rno = 0;
        int parse;
        public AddStudent(Array data)
        {
            InitializeComponent();
            selectClassComboBox.Items.Clear();
            selectClassComboBox.ItemsSource = data;
        }

        private void AddStudent_ButtonClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(studentNameTextBox.Text) || string.IsNullOrWhiteSpace(studentEmailTextBox.Text) || string.IsNullOrWhiteSpace(studentPhoneNumberTextBox.Text) || string.IsNullOrWhiteSpace(selectClassComboBox.Text))
            {
                MessageBox.Show("Please fill in all the details", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                addStudentButton.IsEnabled = false;
                BackgroundWorker worker = new BackgroundWorker();
                worker.WorkerReportsProgress = true;
                worker.DoWork += worker_DoWork;
                worker.ProgressChanged += worker_ProgressChanged;
                worker.RunWorkerCompleted += worker_RunWorkerCompleted;
                worker.RunWorkerAsync(10000);
                
            }
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Dispatcher.Invoke(() =>
            {
                //First we have to add student in Class file, and then we have to add the student in class attendance file.
                //First opening the class file and adding the student.
                className = selectClassComboBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                className = className.Remove(className.Length - 5); (sender as BackgroundWorker).ReportProgress(1);
                FileInfo file = new FileInfo(@"C:\\Rocket\\iClass\\Class\\" + className + ".xlsx");
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                    ExcelWorksheet xlWorksheet = excelWorkBook.Worksheets.First();

                    /* Checking the last empty row and column to save the new student details */
                    /* We have started parsing from 2 because 0 is not a column in excel and 1 column is taken by S.no,name,email..etc */
                    for (parse = 2; parse <= 100; parse++)
                    {
                        if (xlWorksheet.Cells[parse, 1].Value != null)
                        {
                            Sno++; (sender as BackgroundWorker).ReportProgress(1); //sno will give the last value of column
                        }
                        else
                        {
                            //MessageBox.Show(Convert.ToString(Sno));
                            break;
                        }

                    }

                    xlWorksheet.Cells[Sno + 2, 1].Value = Sno + 1; (sender as BackgroundWorker).ReportProgress(1);
                    xlWorksheet.Cells[Sno + 2, 2].Value = studentNameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    xlWorksheet.Cells[Sno + 2, 3].Value = (studentPhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    xlWorksheet.Cells[Sno + 2, 4].Value = studentEmailTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    excelPackage.Save();
                    //**************Adding student in class file....DONE*************************//
                }



                //Now adding the student in attendance file.
                FileInfo fileName = new FileInfo(@"C:\\Rocket\\iClass\\Attendance\\" + className + "_Class_Attendance.xlsx");
                using (ExcelPackage excelPackage = new ExcelPackage(fileName))
                {
                    ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                    ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();
                    /*We have started parsing from 3 because 1 and 2 are filled with Name and toatal attendance */
                    for (parse = 3; parse <= 100; parse++)
                    {
                        if (excelWorksheet.Cells[1, parse].Value != null)
                        {
                            Rno++; (sender as BackgroundWorker).ReportProgress(1); //sno will give the last value of column
                        }
                        else
                        {
                            //MessageBox.Show(Convert.ToString(Rno));
                            break;
                        }

                    }

                    excelWorksheet.Cells[Sno + 2, 1].Value = studentNameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);

                    excelWorksheet.Cells[Sno + 2, 2].Value = Convert.ToInt16("0"); (sender as BackgroundWorker).ReportProgress(1);

                    for (int i = 3; i <= Rno + 2; i++)
                    {
                        excelWorksheet.Cells[Sno + 2, i].Value = "NA"; (sender as BackgroundWorker).ReportProgress(1);
                    }

                    excelPackage.Save();
                }
            });
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                progress.Show();
            }

        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            addStudentButton.IsEnabled = true;
            this.Close();
            progress.Close();

            System.Media.SystemSounds.Exclamation.Play();
            MessageBox.Show("Student added successfully", "Save Success ", MessageBoxButton.OK, MessageBoxImage.Information);
            
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
