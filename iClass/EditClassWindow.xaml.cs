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
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Diagnostics;

namespace iClass
{
    /// <summary>
    /// Interaction logic for EditClassWindow.xaml
    /// </summary>
    public partial class EditClassWindow : System.Windows.Window
    {
        CircularProgressBar progress = new CircularProgressBar();
        string className; int parse; int Sno = 0;
        public EditClassWindow(string data)
        {
            InitializeComponent();
            className = data;
            className = className.Remove(className.Length - 5);
            enableTextBox();
        }

        private void saveStudentDetailsButton_Click(object sender, RoutedEventArgs e)
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync(10000);
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Dispatcher.Invoke(() =>
            {
                /*Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open("C:\\Rocket\\iClass\\Class\\" + className + ".xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet excelWorksheet = (Excel._Worksheet)excelWorkbook.Sheets[1];
                Excel.Range xlRange = excelWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;*/

                FileInfo file = new FileInfo(@"C:\\Rocket\\iClass\\Class\\" + className + ".xlsx");
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                    ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();

                    
                    
                    if (excelWorksheet.Cells[2, 2].Value != null)
                    {
                        excelWorksheet.Cells[2, 2].Value = s1NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[2, 4].Value = s1EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[2, 3].Value != null)
                    {
                        excelWorksheet.Cells[2, 3].Value = Convert.ToInt32(s1PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[3, 2].Value != null)
                    {
                        excelWorksheet.Cells[3, 2].Value = s2NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[3, 4].Value = s2EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[3, 3].Value != null)
                    {
                        excelWorksheet.Cells[3, 3].Value = Convert.ToInt32(s2PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[4, 2].Value != null)
                    {
                        excelWorksheet.Cells[4, 2].Value = s3NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[4, 4].Value = s3EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[4, 3].Value != null)
                    {
                        excelWorksheet.Cells[4, 3].Value = Convert.ToInt32(s3PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[5, 2].Value != null)
                    {
                        excelWorksheet.Cells[5, 2].Value = s4NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[5, 4].Value = s4EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[5, 3].Value != null)
                    {
                        excelWorksheet.Cells[5, 3].Value = Convert.ToInt32(s4PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[6, 2].Value != null)
                    {
                        excelWorksheet.Cells[6, 2].Value = s5NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[6, 4].Value = s5EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[6, 3].Value != null)
                    {
                        excelWorksheet.Cells[6, 3].Value = Convert.ToInt32(s5PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[7, 2].Value != null)
                    {
                        excelWorksheet.Cells[7, 2].Value = s6NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[7, 4].Value = s6EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[7, 3].Value != null)
                    {
                        excelWorksheet.Cells[7, 3].Value = Convert.ToInt32(s6PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[8, 2].Value != null)
                    {
                        excelWorksheet.Cells[8, 2].Value = s7NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[8, 4].Value = s7EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[8, 3].Value != null)
                    {
                        excelWorksheet.Cells[8, 3].Value = Convert.ToInt32(s7PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[9, 2].Value != null)
                    {
                        excelWorksheet.Cells[9, 2].Value = s8NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[9, 4].Value = s8EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[9, 3].Value != null)
                    {
                        excelWorksheet.Cells[9, 3].Value = Convert.ToInt32(s8PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[10, 2].Value != null)
                    {
                        excelWorksheet.Cells[10, 2].Value = s9NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[10, 4].Value = s9EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[10, 3].Value != null)
                    {
                        excelWorksheet.Cells[10, 3].Value = Convert.ToInt32(s9PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[11, 2].Value != null)
                    {
                        excelWorksheet.Cells[11, 2].Value = s10NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[11, 4].Value = s10EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[11, 3].Value != null)
                    {
                        excelWorksheet.Cells[11, 3].Value = Convert.ToInt32(s10PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[12, 2].Value != null)
                    {
                        excelWorksheet.Cells[12, 2].Value = s11NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[12, 4].Value = s11EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[12, 3].Value != null)
                    {
                        excelWorksheet.Cells[12, 3].Value = Convert.ToInt32(s11PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[13, 2].Value != null)
                    {
                        excelWorksheet.Cells[13, 2].Value = s12NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[13, 4].Value = s12EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[13, 3].Value != null)
                    {
                        excelWorksheet.Cells[13, 3].Value = Convert.ToInt32(s12PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[14, 2].Value != null)
                    {
                        excelWorksheet.Cells[14, 2].Value = s13NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[14, 4].Value = s13EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[14, 3].Value != null)
                    {
                        excelWorksheet.Cells[14, 3].Value = Convert.ToInt32(s13PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[15, 2].Value != null)
                    {
                        excelWorksheet.Cells[15, 2].Value = s14NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[15, 4].Value = s14EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[15, 3].Value != null)
                    {
                        excelWorksheet.Cells[15, 3].Value = Convert.ToInt32(s14PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[16, 2].Value != null)
                    {
                        excelWorksheet.Cells[16, 2].Value = s15NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[16, 4].Value = s15EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[16, 3].Value != null)
                    {
                        excelWorksheet.Cells[16, 3].Value = Convert.ToInt32(s15PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[17, 2].Value != null)
                    {
                        excelWorksheet.Cells[17, 2].Value = s16NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[17, 4].Value = s16EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[17, 3].Value != null)
                    {
                        excelWorksheet.Cells[17, 3].Value = Convert.ToInt32(s16PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[18, 2].Value != null)
                    {
                        excelWorksheet.Cells[18, 2].Value = s17NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[18, 4].Value = s17EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[18, 3].Value != null)
                    {
                        excelWorksheet.Cells[18, 3].Value = Convert.ToInt32(s17PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[19, 2].Value != null)
                    {
                        excelWorksheet.Cells[19, 2].Value = s18NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[19, 4].Value = s18EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[19, 3].Value != null)
                    {
                        excelWorksheet.Cells[19, 3].Value = Convert.ToInt32(s18PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[20, 2].Value != null)
                    {
                        excelWorksheet.Cells[20, 2].Value = s19NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[20, 4].Value = s19EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[20, 3].Value != null)
                    {
                        excelWorksheet.Cells[20, 3].Value = Convert.ToInt32(s19PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[21, 2].Value != null)
                    {
                        excelWorksheet.Cells[21, 2].Value = s20NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[21, 4].Value = s20EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[21, 3].Value != null)
                    {
                        excelWorksheet.Cells[21, 3].Value = Convert.ToInt32(s20PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[22, 2].Value != null)
                    {
                        excelWorksheet.Cells[22, 2].Value = s21NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[22, 4].Value = s21EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[22, 3].Value != null)
                    {
                        excelWorksheet.Cells[22, 3].Value = Convert.ToInt32(s21PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[23, 2].Value != null)
                    {
                        excelWorksheet.Cells[23, 2].Value = s22NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[23, 4].Value = s22EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[23, 3].Value != null)
                    {
                        excelWorksheet.Cells[23, 3].Value = Convert.ToInt32(s22PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[24, 2].Value != null)
                    {
                        excelWorksheet.Cells[24, 2].Value = s23NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[24, 4].Value = s23EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[24, 3].Value != null)
                    {
                        excelWorksheet.Cells[24, 3].Value = Convert.ToInt32(s23PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[25, 2].Value != null)
                    {
                        excelWorksheet.Cells[25, 2].Value = s24NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[25, 4].Value = s24EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[25, 3].Value != null)
                    {
                        excelWorksheet.Cells[25, 3].Value = Convert.ToInt32(s24PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[26, 2].Value != null)
                    {
                        excelWorksheet.Cells[26, 2].Value = s25NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[26, 4].Value = s25EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[26, 3].Value != null)
                    {
                        excelWorksheet.Cells[26, 3].Value = Convert.ToInt32(s25PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    if (excelWorksheet.Cells[27, 2].Value != null)
                    {
                        excelWorksheet.Cells[27, 2].Value = s26NameTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    excelWorksheet.Cells[27, 4].Value = s26EmailIdTextBox.Text; (sender as BackgroundWorker).ReportProgress(1);
                    if (excelWorksheet.Cells[27, 3].Value != null)
                    {
                        excelWorksheet.Cells[27, 3].Value = Convert.ToInt32(s26PhoneNumberTextBox.Text); (sender as BackgroundWorker).ReportProgress(1);
                    }

                    excelWorksheet.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    excelWorksheet.Column(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    excelWorksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    excelWorksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    excelWorksheet.Column(1).AutoFit();
                    excelWorksheet.Column(2).AutoFit();
                    excelWorksheet.Column(3).AutoFit();
                    excelWorksheet.Column(4).AutoFit();


                    excelPackage.Save();
                    
                }

                /*Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Rocket\\iClass\\Class\\" + className + "_Class_Attendance" + ".xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                Excel.Range excelRange = xlWorksheet.UsedRange;

                int rCount = excelRange.Rows.Count;
                int cCount = excelRange.Columns.Count;*/

                FileInfo fileName = new FileInfo(@"C:\\Rocket\\iClass\\Attendance\\" + className + "_Class_Attendance" + ".xlsx");
                using (ExcelPackage excelPackage = new ExcelPackage(fileName))
                {
                    ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                    ExcelWorksheet xlWorksheet = excelWorkBook.Worksheets.First();

                    if (xlWorksheet.Cells[2, 1].Value != null)
                    {
                        xlWorksheet.Cells[2, 1].Value = s1NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[3, 1].Value != null)
                    {
                        xlWorksheet.Cells[3, 1].Value = s2NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[4, 1].Value != null)
                    {
                        xlWorksheet.Cells[4, 1].Value = s3NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[5, 1].Value != null)
                    {
                        xlWorksheet.Cells[5, 1].Value = s4NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[6, 1].Value != null)
                    {
                        xlWorksheet.Cells[6, 1].Value = s5NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[7, 1].Value != null)
                    {
                        xlWorksheet.Cells[7, 1].Value = s6NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[8, 1].Value != null)
                    {
                        xlWorksheet.Cells[8, 1].Value = s7NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[9, 1].Value != null)
                    {
                        xlWorksheet.Cells[9, 1].Value = s8NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[10, 1].Value != null)
                    {
                        xlWorksheet.Cells[10, 1].Value = s9NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[11, 1].Value != null)
                    {
                        xlWorksheet.Cells[11, 1].Value = s10NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[12, 1].Value != null)
                    {
                        xlWorksheet.Cells[12, 1].Value = s11NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[13, 1].Value != null)
                    {
                        xlWorksheet.Cells[13, 1].Value = s12NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[14, 1].Value != null)
                    {
                        xlWorksheet.Cells[14, 1].Value = s13NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[15, 1].Value != null)
                    {
                        xlWorksheet.Cells[15, 1].Value = s14NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[16, 1].Value != null)
                    {
                        xlWorksheet.Cells[16, 1].Value = s15NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[17, 1].Value != null)
                    {
                        xlWorksheet.Cells[17, 1].Value = s16NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[18, 1].Value != null)
                    {
                        xlWorksheet.Cells[18, 1].Value = s17NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[19, 1].Value != null)
                    {
                        xlWorksheet.Cells[19, 1].Value = s18NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[20, 1].Value != null)
                    {
                        xlWorksheet.Cells[20, 1].Value = s19NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[21, 1].Value != null)
                    {
                        xlWorksheet.Cells[21, 1].Value = s20NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[22, 1].Value != null)
                    {
                        xlWorksheet.Cells[22, 1].Value = s21NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[23, 1].Value != null)
                    {
                        xlWorksheet.Cells[23, 1].Value = s22NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[24, 1].Value != null)
                    {
                        xlWorksheet.Cells[24, 1].Value = s23NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[25, 1].Value != null)
                    {
                        xlWorksheet.Cells[25, 1].Value = s24NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[26, 1].Value != null)
                    {
                        xlWorksheet.Cells[26, 1].Value = s25NameTextBox.Text;
                    }

                    if (xlWorksheet.Cells[27, 1].Value != null)
                    {
                        xlWorksheet.Cells[27, 1].Value = s26NameTextBox.Text;
                    }

                    xlWorksheet.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    xlWorksheet.Column(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    xlWorksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    xlWorksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    xlWorksheet.Column(1).AutoFit();
                    xlWorksheet.Column(2).AutoFit();
                    xlWorksheet.Column(3).AutoFit();
                    xlWorksheet.Column(4).AutoFit();

                    excelPackage.Save();

                   
                }
                this.Close();
                MessageBox.Show("Details have been successfully updated", "Information", MessageBoxButton.OK, MessageBoxImage.Information);


            });
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //CircularProgressBar progress = new CircularProgressBar();
            if (e.ProgressPercentage == 1)
            {

                progress.Show();
            }
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progress.Close();
        }

        private void clearAllDetailsButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void backButton_Click(object sender, RoutedEventArgs e)
        {

        }


        void enableTextBox()
        {

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open("C:\\Rocket\\iClass\\Class\\" + className + ".xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel._Worksheet excelWorksheet = (Excel._Worksheet)excelWorkbook.Sheets[1];
            Excel.Range xlRange = excelWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            /* We have started parsing from 2 because 0 is not a column in excel and 1 column is taken by S.no,name,email..etc */
            for (parse = 2; parse <= 100; parse++)
            {
                if (xlRange.Cells[parse, 1].Value2 != null)
                {
                    Sno++;
                }
                else
                {
                    //MessageBox.Show(Convert.ToString(Sno));
                    break;
                }

            }


            switch (Sno)
            {
                case 1:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 2:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 3:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 4:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 5:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 6:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 7:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 8:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 9:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 10:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 11:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 12:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 13:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 14:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 15:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 16:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 17:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 18:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    s18NameTextBox.IsEnabled = true; s18EmailIdTextBox.IsEnabled = true; s18PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 19:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    s18NameTextBox.IsEnabled = true; s18EmailIdTextBox.IsEnabled = true; s18PhoneNumberTextBox.IsEnabled = true;
                    s19NameTextBox.IsEnabled = true; s19EmailIdTextBox.IsEnabled = true; s19PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 20:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    s18NameTextBox.IsEnabled = true; s18EmailIdTextBox.IsEnabled = true; s18PhoneNumberTextBox.IsEnabled = true;
                    s19NameTextBox.IsEnabled = true; s19EmailIdTextBox.IsEnabled = true; s19PhoneNumberTextBox.IsEnabled = true;
                    s20NameTextBox.IsEnabled = true; s20EmailIdTextBox.IsEnabled = true; s20PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 21:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    s18NameTextBox.IsEnabled = true; s18EmailIdTextBox.IsEnabled = true; s18PhoneNumberTextBox.IsEnabled = true;
                    s19NameTextBox.IsEnabled = true; s19EmailIdTextBox.IsEnabled = true; s19PhoneNumberTextBox.IsEnabled = true;
                    s20NameTextBox.IsEnabled = true; s20EmailIdTextBox.IsEnabled = true; s20PhoneNumberTextBox.IsEnabled = true;
                    s21NameTextBox.IsEnabled = true; s21EmailIdTextBox.IsEnabled = true; s21PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 22:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    s18NameTextBox.IsEnabled = true; s18EmailIdTextBox.IsEnabled = true; s18PhoneNumberTextBox.IsEnabled = true;
                    s19NameTextBox.IsEnabled = true; s19EmailIdTextBox.IsEnabled = true; s19PhoneNumberTextBox.IsEnabled = true;
                    s20NameTextBox.IsEnabled = true; s20EmailIdTextBox.IsEnabled = true; s20PhoneNumberTextBox.IsEnabled = true;
                    s21NameTextBox.IsEnabled = true; s21EmailIdTextBox.IsEnabled = true; s21PhoneNumberTextBox.IsEnabled = true;
                    s22NameTextBox.IsEnabled = true; s22EmailIdTextBox.IsEnabled = true; s22PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 23:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    s18NameTextBox.IsEnabled = true; s18EmailIdTextBox.IsEnabled = true; s18PhoneNumberTextBox.IsEnabled = true;
                    s19NameTextBox.IsEnabled = true; s19EmailIdTextBox.IsEnabled = true; s19PhoneNumberTextBox.IsEnabled = true;
                    s20NameTextBox.IsEnabled = true; s20EmailIdTextBox.IsEnabled = true; s20PhoneNumberTextBox.IsEnabled = true;
                    s21NameTextBox.IsEnabled = true; s21EmailIdTextBox.IsEnabled = true; s21PhoneNumberTextBox.IsEnabled = true;
                    s22NameTextBox.IsEnabled = true; s22EmailIdTextBox.IsEnabled = true; s22PhoneNumberTextBox.IsEnabled = true;
                    s23NameTextBox.IsEnabled = true; s23EmailIdTextBox.IsEnabled = true; s23PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 24:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    s18NameTextBox.IsEnabled = true; s18EmailIdTextBox.IsEnabled = true; s18PhoneNumberTextBox.IsEnabled = true;
                    s19NameTextBox.IsEnabled = true; s19EmailIdTextBox.IsEnabled = true; s19PhoneNumberTextBox.IsEnabled = true;
                    s20NameTextBox.IsEnabled = true; s20EmailIdTextBox.IsEnabled = true; s20PhoneNumberTextBox.IsEnabled = true;
                    s21NameTextBox.IsEnabled = true; s21EmailIdTextBox.IsEnabled = true; s21PhoneNumberTextBox.IsEnabled = true;
                    s22NameTextBox.IsEnabled = true; s22EmailIdTextBox.IsEnabled = true; s22PhoneNumberTextBox.IsEnabled = true;
                    s23NameTextBox.IsEnabled = true; s23EmailIdTextBox.IsEnabled = true; s23PhoneNumberTextBox.IsEnabled = true;
                    s24NameTextBox.IsEnabled = true; s24EmailIdTextBox.IsEnabled = true; s24PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 25:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    s18NameTextBox.IsEnabled = true; s18EmailIdTextBox.IsEnabled = true; s18PhoneNumberTextBox.IsEnabled = true;
                    s19NameTextBox.IsEnabled = true; s19EmailIdTextBox.IsEnabled = true; s19PhoneNumberTextBox.IsEnabled = true;
                    s20NameTextBox.IsEnabled = true; s20EmailIdTextBox.IsEnabled = true; s20PhoneNumberTextBox.IsEnabled = true;
                    s21NameTextBox.IsEnabled = true; s21EmailIdTextBox.IsEnabled = true; s21PhoneNumberTextBox.IsEnabled = true;
                    s22NameTextBox.IsEnabled = true; s22EmailIdTextBox.IsEnabled = true; s22PhoneNumberTextBox.IsEnabled = true;
                    s23NameTextBox.IsEnabled = true; s23EmailIdTextBox.IsEnabled = true; s23PhoneNumberTextBox.IsEnabled = true;
                    s24NameTextBox.IsEnabled = true; s24EmailIdTextBox.IsEnabled = true; s24PhoneNumberTextBox.IsEnabled = true;
                    s25NameTextBox.IsEnabled = true; s25EmailIdTextBox.IsEnabled = true; s25PhoneNumberTextBox.IsEnabled = true;
                    break;
                case 26:
                    s1NameTextBox.IsEnabled = true; s1EmailIdTextBox.IsEnabled = true; s1PhoneNumberTextBox.IsEnabled = true;
                    s2NameTextBox.IsEnabled = true; s2EmailIdTextBox.IsEnabled = true; s2PhoneNumberTextBox.IsEnabled = true;
                    s3NameTextBox.IsEnabled = true; s3EmailIdTextBox.IsEnabled = true; s3PhoneNumberTextBox.IsEnabled = true;
                    s4NameTextBox.IsEnabled = true; s4EmailIdTextBox.IsEnabled = true; s4PhoneNumberTextBox.IsEnabled = true;
                    s5NameTextBox.IsEnabled = true; s5EmailIdTextBox.IsEnabled = true; s5PhoneNumberTextBox.IsEnabled = true;
                    s6NameTextBox.IsEnabled = true; s6EmailIdTextBox.IsEnabled = true; s6PhoneNumberTextBox.IsEnabled = true;
                    s7NameTextBox.IsEnabled = true; s7EmailIdTextBox.IsEnabled = true; s7PhoneNumberTextBox.IsEnabled = true;
                    s8NameTextBox.IsEnabled = true; s8EmailIdTextBox.IsEnabled = true; s8PhoneNumberTextBox.IsEnabled = true;
                    s9NameTextBox.IsEnabled = true; s9EmailIdTextBox.IsEnabled = true; s9PhoneNumberTextBox.IsEnabled = true;
                    s10NameTextBox.IsEnabled = true; s10EmailIdTextBox.IsEnabled = true; s10PhoneNumberTextBox.IsEnabled = true;
                    s11NameTextBox.IsEnabled = true; s11EmailIdTextBox.IsEnabled = true; s11PhoneNumberTextBox.IsEnabled = true;
                    s12NameTextBox.IsEnabled = true; s12EmailIdTextBox.IsEnabled = true; s12PhoneNumberTextBox.IsEnabled = true;
                    s13NameTextBox.IsEnabled = true; s13EmailIdTextBox.IsEnabled = true; s13PhoneNumberTextBox.IsEnabled = true;
                    s14NameTextBox.IsEnabled = true; s14EmailIdTextBox.IsEnabled = true; s14PhoneNumberTextBox.IsEnabled = true;
                    s15NameTextBox.IsEnabled = true; s15EmailIdTextBox.IsEnabled = true; s15PhoneNumberTextBox.IsEnabled = true;
                    s16NameTextBox.IsEnabled = true; s16EmailIdTextBox.IsEnabled = true; s16PhoneNumberTextBox.IsEnabled = true;
                    s17NameTextBox.IsEnabled = true; s17EmailIdTextBox.IsEnabled = true; s17PhoneNumberTextBox.IsEnabled = true;
                    s18NameTextBox.IsEnabled = true; s18EmailIdTextBox.IsEnabled = true; s18PhoneNumberTextBox.IsEnabled = true;
                    s19NameTextBox.IsEnabled = true; s19EmailIdTextBox.IsEnabled = true; s19PhoneNumberTextBox.IsEnabled = true;
                    s20NameTextBox.IsEnabled = true; s20EmailIdTextBox.IsEnabled = true; s20PhoneNumberTextBox.IsEnabled = true;
                    s21NameTextBox.IsEnabled = true; s21EmailIdTextBox.IsEnabled = true; s21PhoneNumberTextBox.IsEnabled = true;
                    s22NameTextBox.IsEnabled = true; s22EmailIdTextBox.IsEnabled = true; s22PhoneNumberTextBox.IsEnabled = true;
                    s23NameTextBox.IsEnabled = true; s23EmailIdTextBox.IsEnabled = true; s23PhoneNumberTextBox.IsEnabled = true;
                    s24NameTextBox.IsEnabled = true; s24EmailIdTextBox.IsEnabled = true; s24PhoneNumberTextBox.IsEnabled = true;
                    s25NameTextBox.IsEnabled = true; s25EmailIdTextBox.IsEnabled = true; s25PhoneNumberTextBox.IsEnabled = true;
                    s26NameTextBox.IsEnabled = true; s26EmailIdTextBox.IsEnabled = true; s26PhoneNumberTextBox.IsEnabled = true;
                    break;
                default:

                    MessageBox.Show("Number of students is more than the maximum number of students\n\n", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                    break;


            }

            s1NameTextBox.Text = (xlRange.Cells[2, 2] as Excel.Range).Value2;
            s1PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[2, 3] as Excel.Range).Value2);
            s1EmailIdTextBox.Text = (xlRange.Cells[2, 4] as Excel.Range).Value2;

            s2NameTextBox.Text = (xlRange.Cells[3, 2] as Excel.Range).Value2;
            s2PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[3, 3] as Excel.Range).Value2);
            s2EmailIdTextBox.Text = (xlRange.Cells[3, 4] as Excel.Range).Value2;

            s3NameTextBox.Text = (xlRange.Cells[4, 2] as Excel.Range).Value2;
            s3PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[4, 3] as Excel.Range).Value2);
            s3EmailIdTextBox.Text = (xlRange.Cells[4, 4] as Excel.Range).Value2;

            s4NameTextBox.Text = (xlRange.Cells[5, 2] as Excel.Range).Value2;
            s4PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[5, 3] as Excel.Range).Value2);
            s4EmailIdTextBox.Text = (xlRange.Cells[5, 4] as Excel.Range).Value2;

            s5NameTextBox.Text = (xlRange.Cells[6, 2] as Excel.Range).Value2;
            s5PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[6, 3] as Excel.Range).Value2);
            s5EmailIdTextBox.Text = (xlRange.Cells[6, 4] as Excel.Range).Value2;

            s6NameTextBox.Text = (xlRange.Cells[7, 2] as Excel.Range).Value2;
            s6PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[7, 3] as Excel.Range).Value2);
            s6EmailIdTextBox.Text = (xlRange.Cells[7, 4] as Excel.Range).Value2;

            s7NameTextBox.Text = (xlRange.Cells[8, 2] as Excel.Range).Value2;
            s7PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[8, 3] as Excel.Range).Value2);
            s7EmailIdTextBox.Text = (xlRange.Cells[8, 4] as Excel.Range).Value2;

            s8NameTextBox.Text = (xlRange.Cells[9, 2] as Excel.Range).Value2;
            s8PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[9, 3] as Excel.Range).Value2);
            s8EmailIdTextBox.Text = (xlRange.Cells[9, 4] as Excel.Range).Value2;

            s9NameTextBox.Text = (xlRange.Cells[10, 2] as Excel.Range).Value2;
            s9PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[10, 3] as Excel.Range).Value2);
            s9EmailIdTextBox.Text = (xlRange.Cells[10, 4] as Excel.Range).Value2;

            s10NameTextBox.Text = (xlRange.Cells[11, 2] as Excel.Range).Value2;
            s10PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[11, 3] as Excel.Range).Value2);
            s10EmailIdTextBox.Text = (xlRange.Cells[11, 4] as Excel.Range).Value2;

            s11NameTextBox.Text = (xlRange.Cells[12, 2] as Excel.Range).Value2;
            s11PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[12, 3] as Excel.Range).Value2);
            s11EmailIdTextBox.Text = (xlRange.Cells[12, 4] as Excel.Range).Value2;

            s12NameTextBox.Text = (xlRange.Cells[13, 2] as Excel.Range).Value2;
            s12PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[13, 3] as Excel.Range).Value2);
            s12EmailIdTextBox.Text = (xlRange.Cells[13, 4] as Excel.Range).Value2;

            s13NameTextBox.Text = (xlRange.Cells[14, 2] as Excel.Range).Value2;
            s13PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[14, 3] as Excel.Range).Value2);
            s13EmailIdTextBox.Text = (xlRange.Cells[14, 4] as Excel.Range).Value2;

            s14NameTextBox.Text = (xlRange.Cells[15, 2] as Excel.Range).Value2;
            s14PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[15, 3] as Excel.Range).Value2);
            s14EmailIdTextBox.Text = (xlRange.Cells[15, 4] as Excel.Range).Value2;

            s15NameTextBox.Text = (xlRange.Cells[16, 2] as Excel.Range).Value2;
            s15PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[16, 3] as Excel.Range).Value2);
            s15EmailIdTextBox.Text = (xlRange.Cells[16, 4] as Excel.Range).Value2;

            s16NameTextBox.Text = (xlRange.Cells[17, 2] as Excel.Range).Value2;
            s16PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[17, 3] as Excel.Range).Value2);
            s16EmailIdTextBox.Text = (xlRange.Cells[17, 4] as Excel.Range).Value2;

            s17NameTextBox.Text = (xlRange.Cells[18, 2] as Excel.Range).Value2;
            s17PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[18, 3] as Excel.Range).Value2);
            s17EmailIdTextBox.Text = (xlRange.Cells[18, 4] as Excel.Range).Value2;

            s18NameTextBox.Text = (xlRange.Cells[19, 2] as Excel.Range).Value2;
            s18PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[19, 3] as Excel.Range).Value2);
            s18EmailIdTextBox.Text = (xlRange.Cells[19, 4] as Excel.Range).Value2;

            s19NameTextBox.Text = (xlRange.Cells[20, 2] as Excel.Range).Value2;
            s19PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[20, 3] as Excel.Range).Value2);
            s19EmailIdTextBox.Text = (xlRange.Cells[20, 4] as Excel.Range).Value2;

            s20NameTextBox.Text = (xlRange.Cells[21, 2] as Excel.Range).Value2;
            s20PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[21, 3] as Excel.Range).Value2);
            s20EmailIdTextBox.Text = (xlRange.Cells[21, 4] as Excel.Range).Value2;

            s21NameTextBox.Text = (xlRange.Cells[22, 2] as Excel.Range).Value2;
            s21PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[22, 3] as Excel.Range).Value2);
            s21EmailIdTextBox.Text = (xlRange.Cells[22, 4] as Excel.Range).Value2;

            s22NameTextBox.Text = (xlRange.Cells[23, 2] as Excel.Range).Value2;
            s22PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[23, 3] as Excel.Range).Value2);
            s22EmailIdTextBox.Text = (xlRange.Cells[23, 4] as Excel.Range).Value2;

            s23NameTextBox.Text = (xlRange.Cells[24, 2] as Excel.Range).Value2;
            s23PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[24, 3] as Excel.Range).Value2);
            s23EmailIdTextBox.Text = (xlRange.Cells[24, 4] as Excel.Range).Value2;

            s24NameTextBox.Text = (xlRange.Cells[25, 2] as Excel.Range).Value2;
            s24PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[25, 3] as Excel.Range).Value2);
            s24EmailIdTextBox.Text = (xlRange.Cells[25, 4] as Excel.Range).Value2;

            s25NameTextBox.Text = (xlRange.Cells[26, 2] as Excel.Range).Value2;
            s25PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[26, 3] as Excel.Range).Value2);
            s25EmailIdTextBox.Text = (xlRange.Cells[26, 4] as Excel.Range).Value2;

            s26NameTextBox.Text = (xlRange.Cells[27, 2] as Excel.Range).Value2;
            s26PhoneNumberTextBox.Text = Convert.ToString((xlRange.Cells[27, 3] as Excel.Range).Value2);
            s26EmailIdTextBox.Text = (xlRange.Cells[27, 4] as Excel.Range).Value2;





            excelWorkbook.Close();
            excelApp.Quit();



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
