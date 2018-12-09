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
using System.Diagnostics;


namespace iClass
{
    /// <summary>
    /// Interaction logic for AddTeacherPopUp.xaml
    /// </summary>
    public partial class AddTeacherPopUp : System.Windows.Window
    {
        CircularProgressBar progress = new CircularProgressBar();
        
        public AddTeacherPopUp()
        {
            InitializeComponent();
            
        }

        string name, emailID, phoneNumber;
        int Sno = 0;
        int parse;
        private void SaveTeacher_ButtonClick(object sender, RoutedEventArgs e)
        {
            SaveTeacherDetailsButton.IsEnabled = false;
             name = teacherNameTextBox.Text;
             emailID = teacherEmailIdTextBox.Text;
             phoneNumber = teacherPhoneNumberTextBox.Text;

             if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(emailID) || string.IsNullOrWhiteSpace(phoneNumber))
             {
                 System.Media.SystemSounds.Hand.Play();
                 MessageBox.Show("Please fill all the details   ", "Error ", MessageBoxButton.OK, MessageBoxImage.Error);
                 this.Close();
             }
             else
             {
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
                    Mouse.OverrideCursor = Cursors.Wait;
                    if (File.Exists(@"C:\\Rocket\\iClass\\Teacher\\TeacherDetails.xlsx"))
                    {

                        (sender as BackgroundWorker).ReportProgress(1);
                        openAndSaveData();
                        (sender as BackgroundWorker).ReportProgress(0);
                    }
                    if (!(File.Exists(@"C:\\Rocket\\iClass\\Teacher\\TeacherDetails.xlsx")))
                    {

                        (sender as BackgroundWorker).ReportProgress(1);
                        createAndSaveData();
                        (sender as BackgroundWorker).ReportProgress(0);
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
            Mouse.OverrideCursor = null;
            SaveTeacherDetailsButton.IsEnabled = true;
            this.Close();
            progress.Close();
            
            System.Media.SystemSounds.Exclamation.Play();
            MessageBox.Show("Teacher added successfully", "Save Success ", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void openAndSaveData()
        {
            //Opening the TeacherDetails.xlsx file
            FileInfo file = new FileInfo(@"C:\\Rocket\\iClass\\Teacher\\TeacherDetails.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet worksheet = excelWorkBook.Worksheets.First();


                var rowCnt = worksheet.Dimension.End.Row;
                
                worksheet.Cells[rowCnt + 1, 1].Value = rowCnt;
                worksheet.Cells[rowCnt + 1, 2].Value = name;
                worksheet.Cells[rowCnt + 1, 3].Value = emailID;
                worksheet.Cells[rowCnt + 1, 4].Value = phoneNumber;

                worksheet.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Column(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                worksheet.Column(1).AutoFit();
                worksheet.Column(2).AutoFit();
                worksheet.Column(3).AutoFit();
                worksheet.Column(4).AutoFit();

                excelPackage.Save();
                Log("Teacher added : " + name);
            } 
            
        }

        private void createAndSaveData()
        {
            var fileName = "TeacherDetails.xlsx";
            var outputDir = @"C:\\Rocket\\iClass\\Teacher\\";

            // Create the file using the FileInfo object
            var file = new FileInfo(outputDir + fileName);

            using (var package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Teacher Details");

                // --------- Data and styling goes here -------------- //

                worksheet.Cells[1, 1].Value = "S.No";
                worksheet.Cells[1, 2].Value = " Name ";
                worksheet.Cells[1, 3].Value = " Email Address ";
                worksheet.Cells[1, 4].Value = " Phone Number ";
                worksheet.Cells[1, 1].Style.Font.Bold = true;
                worksheet.Cells[1, 2].Style.Font.Bold = true;
                worksheet.Cells[1, 3].Style.Font.Bold = true;
                worksheet.Cells[1, 4].Style.Font.Bold = true;
                worksheet.Column(1).AutoFit();
                worksheet.Column(2).AutoFit();
                worksheet.Column(3).AutoFit();
                worksheet.Column(4).AutoFit();

                worksheet.Cells[2, 1].Value = Convert.ToInt32( "1");
                worksheet.Cells[2, 2].Value = name;
                worksheet.Cells[2, 3].Value = emailID;
                worksheet.Cells[2, 4].Value = phoneNumber;

                package.Save();
                Log("Teacher added : " + name);

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



 