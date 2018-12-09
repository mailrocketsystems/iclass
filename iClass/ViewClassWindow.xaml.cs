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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Windows.Threading;
using System.Threading;
using System.Net;
using System.Net.Mail;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Diagnostics;


namespace iClass
{
    /// <summary>
    /// Interaction logic for ViewClassWindow.xaml
    /// </summary>
    public partial class ViewClassWindow : System.Windows.Window
    {
        string className;
        public ViewClassWindow(string classData)
        {
            InitializeComponent();
            className = classData;
            function();           
            
        }

        public void function()
        {
            /*Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Rocket\\iClass\\Class\\" + className + ".xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;*/
            FileInfo file = new FileInfo(@"C:\\Rocket\\iClass\\Class\\" + className + ".xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet worksheet = excelWorkBook.Worksheets.First();

                s1NameTextBox.Text = Convert.ToString(worksheet.Cells[2, 2].Value);
                s1PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[2, 3].Value);
                s1EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[2, 4].Value);

                s2NameTextBox.Text = Convert.ToString(worksheet.Cells[3, 2].Value);
                s2PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[3, 3].Value);
                s2EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[3, 4].Value);

                s3NameTextBox.Text = Convert.ToString(worksheet.Cells[4, 2].Value);
                s3PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[4, 3].Value);
                s3EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[4, 4].Value);

                s4NameTextBox.Text = Convert.ToString(worksheet.Cells[5, 2].Value);
                s4PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[5, 3].Value);
                s4EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[5, 4].Value);

                s5NameTextBox.Text = Convert.ToString(worksheet.Cells[6, 2].Value);
                s5PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[6, 3].Value);
                s5EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[6, 4].Value);

                s6NameTextBox.Text = Convert.ToString(worksheet.Cells[7, 2].Value);
                s6PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[7, 3].Value);
                s6EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[7, 4].Value);

                s7NameTextBox.Text = Convert.ToString(worksheet.Cells[8, 2].Value);
                s7PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[8, 3].Value);
                s7EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[8, 4].Value);

                s8NameTextBox.Text = Convert.ToString(worksheet.Cells[9, 2].Value);
                s8PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[9, 3].Value);
                s8EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[9, 4].Value);

                s9NameTextBox.Text = Convert.ToString(worksheet.Cells[10, 2].Value);
                s9PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[10, 3].Value);
                s9EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[10, 4].Value);

                s10NameTextBox.Text = Convert.ToString(worksheet.Cells[11, 2].Value);
                s10PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[11, 3].Value);
                s10EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[11, 4].Value);

                s11NameTextBox.Text = Convert.ToString(worksheet.Cells[12, 2].Value);
                s11PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[12, 3].Value);
                s11EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[12, 4].Value);

                s12NameTextBox.Text = Convert.ToString(worksheet.Cells[13, 2].Value);
                s12PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[13, 3].Value);
                s12EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[13, 4].Value);

                s13NameTextBox.Text = Convert.ToString(worksheet.Cells[14, 2].Value);
                s13PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[14, 3].Value);
                s13EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[14, 4].Value);

                s14NameTextBox.Text = Convert.ToString(worksheet.Cells[15, 2].Value);
                s14PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[15, 3].Value);
                s14EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[15, 4].Value);

                s15NameTextBox.Text = Convert.ToString(worksheet.Cells[16, 2].Value);
                s15PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[16, 3].Value);
                s15EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[16, 4].Value);

                s16NameTextBox.Text = Convert.ToString(worksheet.Cells[17, 2].Value);
                s16PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[17, 3].Value);
                s16EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[17, 4].Value);

                s17NameTextBox.Text = Convert.ToString(worksheet.Cells[18, 2].Value);
                s17PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[18, 3].Value);
                s17EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[18, 4].Value);

                s18NameTextBox.Text = Convert.ToString(worksheet.Cells[19, 2].Value);
                s18PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[19, 3].Value);
                s18EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[19, 4].Value);

                s19NameTextBox.Text = Convert.ToString(worksheet.Cells[20, 2].Value);
                s19PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[20, 3].Value);
                s19EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[20, 4].Value);

                s20NameTextBox.Text = Convert.ToString(worksheet.Cells[21, 2].Value);
                s20PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[21, 3].Value);
                s20EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[21, 4].Value);

                s21NameTextBox.Text = Convert.ToString(worksheet.Cells[22, 2].Value);
                s21PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[22, 3].Value);
                s21EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[22, 4].Value);

                s22NameTextBox.Text = Convert.ToString(worksheet.Cells[23, 2].Value);
                s22PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[23, 3].Value);
                s22EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[23, 4].Value);

                s23NameTextBox.Text = Convert.ToString(worksheet.Cells[24, 2].Value);
                s23PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[24, 3].Value);
                s23EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[24, 4].Value);

                s24NameTextBox.Text = Convert.ToString(worksheet.Cells[25, 2].Value);
                s24PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[25, 3].Value);
                s24EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[25, 4].Value);

                s25NameTextBox.Text = Convert.ToString(worksheet.Cells[26, 2].Value);
                s25PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[26, 3].Value);
                s25EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[26, 4].Value);

                s26NameTextBox.Text = Convert.ToString(worksheet.Cells[27, 2].Value);
                s26PhoneNumberTextBox.Text = Convert.ToString(worksheet.Cells[27, 3].Value);
                s26EmailIdTextBox.Text = Convert.ToString(worksheet.Cells[27, 4].Value);

            }

            
            FileInfo fileName = new FileInfo(@"C:\\Rocket\\iClass\\Attendance\\" + className + "_Class_Attendance.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(fileName))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet worksheet = excelWorkBook.Worksheets.First();


                if (worksheet.Cells[2, 2].Value != null)
                {
                    int s1a = Convert.ToInt16(worksheet.Cells[2, 2].Value);
                    s1Attendance.Text = Convert.ToString(Convert.ToString(s1a) + "%");
                }

                if (worksheet.Cells[3, 2].Value != null)
                {
                    int s2a = Convert.ToInt16(worksheet.Cells[3, 2].Value);
                    s2Attendance.Text = Convert.ToString(Convert.ToString(s2a) + "%");
                }

                if (worksheet.Cells[4, 2].Value != null)
                {
                    int s3a = Convert.ToInt16(worksheet.Cells[4, 2].Value);
                    s3Attendance.Text = Convert.ToString(Convert.ToString(s3a) + "%");
                }

                if (worksheet.Cells[5, 2].Value != null)
                {
                    int s4a = Convert.ToInt16(worksheet.Cells[5, 2].Value);
                    s4Attendance.Text = Convert.ToString(Convert.ToString(s4a) + "%");
                }

                if (worksheet.Cells[6, 2].Value != null)
                {
                    int s5a = Convert.ToInt16(worksheet.Cells[6, 2].Value);
                    s5Attendance.Text = Convert.ToString(Convert.ToString(s5a) + "%");
                }

                if (worksheet.Cells[7, 2].Value != null)
                {
                    int s6a = Convert.ToInt16(worksheet.Cells[7, 2].Value);
                    s6Attendance.Text = Convert.ToString(Convert.ToString(s6a) + "%");
                }

                if (worksheet.Cells[8, 2].Value != null)
                {
                    int s7a = Convert.ToInt16(worksheet.Cells[8, 2].Value);
                    s7Attendance.Text = Convert.ToString(Convert.ToString(s7a) + "%");
                }

                if (worksheet.Cells[9, 2].Value != null)
                {
                    int s8a = Convert.ToInt16(worksheet.Cells[9, 2].Value);
                    s8Attendance.Text = Convert.ToString(Convert.ToString(s8a) + "%");
                }

                if (worksheet.Cells[10, 2].Value != null)
                {
                    int s9a = Convert.ToInt16(worksheet.Cells[10, 2].Value);
                    s9Attendance.Text = Convert.ToString(Convert.ToString(s9a) + "%");
                }

                if (worksheet.Cells[11, 2].Value != null)
                {
                    int s10a = Convert.ToInt16(worksheet.Cells[11, 2].Value);
                    s10Attendance.Text = Convert.ToString(Convert.ToString(s10a) + "%");
                }

                if (worksheet.Cells[12, 2].Value != null)
                {
                    int s11a = Convert.ToInt16(worksheet.Cells[12, 2].Value);
                    s11Attendance.Text = Convert.ToString(Convert.ToString(s11a) + "%");
                }

                if (worksheet.Cells[13, 2].Value != null)
                {
                    int s12a = Convert.ToInt16(worksheet.Cells[13, 2].Value);
                    s12Attendance.Text = Convert.ToString(Convert.ToString(s12a) + "%");
                }

                if (worksheet.Cells[14, 2].Value != null)
                {
                    int s13a = Convert.ToInt16(worksheet.Cells[14, 2].Value);
                    s13Attendance.Text = Convert.ToString(Convert.ToString(s13a) + "%");
                }

                if (worksheet.Cells[15, 2].Value != null)
                {
                    int s14a = Convert.ToInt16(worksheet.Cells[15, 2].Value);
                    s14Attendance.Text = Convert.ToString(Convert.ToString(s14a) + "%");
                }

                if (worksheet.Cells[16, 2].Value != null)
                {
                    int s15a = Convert.ToInt16(worksheet.Cells[16, 2].Value);
                    s15Attendance.Text = Convert.ToString(Convert.ToString(s15a) + "%");
                }

                if (worksheet.Cells[17, 2].Value != null)
                {
                    int s16a = Convert.ToInt16(worksheet.Cells[17, 2].Value);
                    s16Attendance.Text = Convert.ToString(Convert.ToString(s16a) + "%");
                }

                if (worksheet.Cells[18, 2].Value != null)
                {
                    int s17a = Convert.ToInt16(worksheet.Cells[18, 2].Value);
                    s17Attendance.Text = Convert.ToString(Convert.ToString(s17a) + "%");
                }

                if (worksheet.Cells[19, 2].Value != null)
                {
                    int s18a = Convert.ToInt16(worksheet.Cells[19, 2].Value);
                    s18Attendance.Text = Convert.ToString(Convert.ToString(s18a) + "%");
                }

                if (worksheet.Cells[20, 2].Value != null)
                {
                    int s19a = Convert.ToInt16(worksheet.Cells[20, 2].Value);
                    s19Attendance.Text = Convert.ToString(Convert.ToString(s19a) + "%");
                }

                if (worksheet.Cells[21, 2].Value != null)
                {
                    int s20a = Convert.ToInt16(worksheet.Cells[21, 2].Value);
                    s20Attendance.Text = Convert.ToString(Convert.ToString(s20a) + "%");
                }

                if (worksheet.Cells[22, 2].Value != null)
                {
                    int s21a = Convert.ToInt16(worksheet.Cells[22, 2].Value);
                    s21Attendance.Text = Convert.ToString(Convert.ToString(s21a) + "%");
                }

                if (worksheet.Cells[23, 2].Value != null)
                {
                    int s22a = Convert.ToInt16(worksheet.Cells[23, 2].Value);
                    s22Attendance.Text = Convert.ToString(Convert.ToString(s22a) + "%");
                }

                if (worksheet.Cells[24, 2].Value != null)
                {
                    int s23a = Convert.ToInt16(worksheet.Cells[24, 2].Value);
                    s23Attendance.Text = Convert.ToString(Convert.ToString(s23a) + "%");
                }

                if (worksheet.Cells[25, 2].Value != null)
                {
                    int s24a = Convert.ToInt16(worksheet.Cells[25, 2].Value);
                    s24Attendance.Text = Convert.ToString(Convert.ToString(s24a) + "%");
                }

                if (worksheet.Cells[26, 2].Value != null)
                {
                    int s25a = Convert.ToInt16(worksheet.Cells[26, 2].Value);
                    s25Attendance.Text = Convert.ToString(Convert.ToString(s25a) + "%");
                }

                if (worksheet.Cells[27, 2].Value != null)
                {
                    int s26a = Convert.ToInt16(worksheet.Cells[27, 2].Value);
                    s26Attendance.Text = Convert.ToString(Convert.ToString(s26a) + "%");
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
