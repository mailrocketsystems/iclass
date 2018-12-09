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
using System.Runtime.InteropServices;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;


namespace iClass
{
    /// <summary>
    /// Interaction logic for StartClassWindow.xaml
    /// </summary>
    public partial class StartClassWindow : System.Windows.Window
    {
        System.Net.Mail.Attachment attachment;
        MailMessage mail = new MailMessage();
        string className;
        //bool runOnce = false;
        bool attendaceStatus = false;
        bool attachStatus = false;
        int parse;
        int Sno = 0;
        int enable = 0;
        CircularProgressBar progress = new CircularProgressBar();
        string s1email, s2email, s3email, s4email, s5email, s6email, s7email, s8email, s9email, s10email, s11email, s12email, s13email = null;
        string s14email, s15email, s16email, s17email, s18email, s19email, s20email, s21email, s22email, s23email, s24email, s25email, s26email = null;
        public StartClassWindow(string classData)
        {
            InitializeComponent();
            Log("Start class window Active Now");
            className = classData;
            startTimeTextBox.Text = DateTime.Now.ToLongTimeString();
            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += timer_Tick;
            timer.Start();
            Thread networkThread = new Thread(Check_Network_Status);
            networkThread.IsBackground = true;
            networkThread.Start();
            classNameTextBox.Text = className;
            function();

        }
        public StartClassWindow()
        {
        }
        
        /************ EMAIL SEND ****************/
        private void EmailSend_ClickButton(object sender, RoutedEventArgs e)
        {
            BackgroundWorker work = new BackgroundWorker();
            work.WorkerReportsProgress = true;
            work.DoWork += work_DoWork;
            work.ProgressChanged += work_ProgressChanged;
            work.RunWorkerCompleted += work_RunWorkerCompleted;
            work.RunWorkerAsync();
                        
        }

        void work_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Dispatcher.Invoke(() =>
                {
                    Log("Email Send button clicked");
                    Mouse.OverrideCursor = Cursors.Wait;
                    (sender as BackgroundWorker).ReportProgress(1);
                    SmtpClient SmtpServer = new SmtpClient("smtp.zoho.com");

                    mail.From = new MailAddress("rocketeducation@zoho.com");
                    if (enable == 1) { mail.To.Add(s1email); }
                    if (enable == 2) { mail.To.Add(s1email); mail.To.Add(s2email); }
                    if (enable == 3) { mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); }
                    if (enable == 4) { mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); }
                    if (enable == 5) { mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); }
                    if (enable == 6) { mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); }
                    if (enable == 7) { mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email); }
                    if (enable == 8)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email);
                        mail.To.Add(s7email); mail.To.Add(s8email);
                    }
                    if (enable == 9)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email);
                    }
                    if (enable == 10)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email);
                    }
                    if (enable == 11)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email);
                    }
                    if (enable == 12)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email);
                    }
                    if (enable == 13)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email);
                    }
                    if (enable == 14)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                    }
                    if (enable == 15)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email);
                    }
                    if (enable == 16)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email);
                    }
                    if (enable == 17)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email);
                    }
                    if (enable == 18)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email); mail.To.Add(s18email);
                    }
                    if (enable == 19)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email); mail.To.Add(s18email); mail.To.Add(s19email);
                    }
                    if (enable == 20)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email); mail.To.Add(s18email); mail.To.Add(s19email); mail.To.Add(s20email);
                    }
                    if (enable == 21)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email); mail.To.Add(s18email); mail.To.Add(s19email); mail.To.Add(s20email);
                        mail.To.Add(s21email);
                    }
                    if (enable == 22)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email); mail.To.Add(s18email); mail.To.Add(s19email); mail.To.Add(s20email);
                        mail.To.Add(s21email); mail.To.Add(s22email);
                    }
                    if (enable == 23)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email); mail.To.Add(s18email); mail.To.Add(s19email); mail.To.Add(s20email);
                        mail.To.Add(s21email); mail.To.Add(s22email); mail.To.Add(s23email);
                    }

                    if (enable == 24)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email); mail.To.Add(s18email); mail.To.Add(s19email); mail.To.Add(s20email);
                        mail.To.Add(s21email); mail.To.Add(s22email); mail.To.Add(s23email); mail.To.Add(s24email);
                    }
                    if (enable == 25)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email); mail.To.Add(s18email); mail.To.Add(s19email); mail.To.Add(s20email);
                        mail.To.Add(s21email); mail.To.Add(s22email); mail.To.Add(s23email); mail.To.Add(s24email); mail.To.Add(s25email);
                    }
                    if (enable == 26)
                    {
                        mail.To.Add(s1email); mail.To.Add(s2email); mail.To.Add(s3email); mail.To.Add(s4email); mail.To.Add(s5email); mail.To.Add(s6email); mail.To.Add(s7email);
                        mail.To.Add(s8email); mail.To.Add(s9email); mail.To.Add(s10email); mail.To.Add(s11email); mail.To.Add(s12email); mail.To.Add(s13email); mail.To.Add(s14email);
                        mail.To.Add(s15email); mail.To.Add(s16email); mail.To.Add(s17email); mail.To.Add(s18email); mail.To.Add(s19email); mail.To.Add(s20email);
                        mail.To.Add(s21email); mail.To.Add(s22email); mail.To.Add(s23email); mail.To.Add(s24email); mail.To.Add(s25email); mail.To.Add(s26email);
                    }

                    //File.Copy(@"C:\\Rocket\\iClass\\Attendance\\" + className + "_Class_Attendance.xlsx", @"C:\\Rocket\\iClass\\Attendance\\" + className + "_Attendance.xlsx");

                    attachment = new System.Net.Mail.Attachment(@"C:\\Rocket\\iClass\\Attendance\\" + className + "_Class_Attendance.xlsx");
                    mail.Attachments.Add(attachment);
                    //lbFiles.Items.Add(@"C:\\Rocket\\iClass\\Attendance\\" + className + "_Class_Attendance.xlsx");
                    lbFiles.Items.Add(className + "_Attendance.xlsx");
                    mail.Subject = emailSubjectTextBox.Text;
                    mail.Body = emailBodyTextBox.Text;

                    //System.Net.Mail.Attachment attachment;
                    //attachment = new System.Net.Mail.Attachment(Convert.ToString(lbFiles.Items));

                    if ((string.IsNullOrWhiteSpace(emailSubjectTextBox.Text)) && (string.IsNullOrWhiteSpace(emailBodyTextBox.Text)))
                    {
                        Log("Error  :   Subject or Body can not be left empty");
                        MessageBox.Show("Subject or Body can not be left empty", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        if (attendaceStatus == false || attachStatus == false)
                        {
                            Log("Error  :   You have either not saved today's attendance or have not attachements");
                            MessageBoxResult result = MessageBox.Show("You have either not saved today's attendance or have not attachements\n" + "Are you sure you want to continue", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Error);
                            if (result == MessageBoxResult.Yes)
                            {
                                SmtpServer.Port = 587;
                                SmtpServer.Credentials = new System.Net.NetworkCredential("rocketeducation", "Truthisgod@27");
                                SmtpServer.EnableSsl = true;
                                SmtpServer.Send(mail);
                                Log("Email Sent Successfully");
                                sendCircle.Fill = Brushes.Green;
                                sendCircle.Stroke = Brushes.Green;
                                sendPath.Fill = Brushes.Green;
                                sendPath.Stroke = Brushes.Green;
                                completeCircle.Fill = Brushes.Green;
                                completeCircle.Stroke = Brushes.Green;
                                emailLabel.Content = "Send Success";
                                completeLabel.Content = "Class Completed";
                                MessageBox.Show("Class completed successfully ", "Congratulations", MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                            else
                            {
                            }
                        }
                        else
                        {
                            SmtpServer.Port = 587;
                            SmtpServer.Credentials = new System.Net.NetworkCredential("rocketeducation", "Truthisgod@27");
                            SmtpServer.EnableSsl = true;
                            SmtpServer.Send(mail);
                            Log("Email Sent Successfully");
                            sendCircle.Fill = Brushes.Green;
                            sendCircle.Stroke = Brushes.Green;
                            sendPath.Fill = Brushes.Green;
                            sendPath.Stroke = Brushes.Green;
                            completeCircle.Fill = Brushes.Green;
                            completeCircle.Stroke = Brushes.Green;
                            emailLabel.Content = "Send Success";
                            completeLabel.Content = "Class Completed";
                            MessageBox.Show("Class completed successfully ", "Congratulations", MessageBoxButton.OK, MessageBoxImage.Information);
                        }


                    }
                    mail.Dispose();
                    //File.Delete(@"C:\\Rocket\\iClass\\Attendance\\" + className + "_Attendance.xlsx");
                    (sender as BackgroundWorker).ReportProgress(0);
                });
        }
        void work_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                //progress.Show();
            }
        }

        void work_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Mouse.OverrideCursor = null;
            //progress.Close();

        }

        
        /************ Attach Files ****************/
        private void AttachFiles_ClickButton(object sender, RoutedEventArgs e)
        {
            
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                string filename = dlg.FileName;
               
                attachment = new System.Net.Mail.Attachment(filename);
                mail.Attachments.Add(attachment);
                lbFiles.Items.Add(filename);
                attachStatus = true;
            }
            attachFileCircle.Fill = Brushes.Green;
            attachFilePath.Fill = Brushes.Green;
            attachLabel.Content = "Attachement Done";
            System.Media.SystemSounds.Exclamation.Play();
            attachFileCircle.Stroke = Brushes.Green;
            attachFilePath.Stroke = Brushes.Green;
        }

        /*********** BACK BUTTON **********************/
        private void Back_ClickButton(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


        //**************************** SAVE ATTENDANCE *******************************//
        private void SaveAttendance_ButtonClick(object sender, RoutedEventArgs e)
        {
            SaveAttendanceButton.IsEnabled = false;
            BackgroundWorker worker = new BackgroundWorker(); 
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync();
            
                        
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Dispatcher.Invoke(() =>
                {
                    Mouse.OverrideCursor = Cursors.Wait;
                    int row = 0;
                    int col = 0;
                    int newCol = 0;
                    int newRow = 0;
                    try
                    {
                    FileInfo file = new FileInfo(@"C:\\Rocket\\iClass\\Attendance\\" + className + "_Class_Attendance.xlsx");   //Open the attendace file
                    using (ExcelPackage excelPackage = new ExcelPackage(file))
                    {
                        ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                        ExcelWorksheet xlWorksheet = excelWorkBook.Worksheets.First();


                        for (parse = 3; parse <= 100; parse++)
                        {
                            if (xlWorksheet.Cells[1, parse].Value != null)
                            {
                                col++; (sender as BackgroundWorker).ReportProgress(1);
                            }
                            else
                            {
                                // MessageBox.Show(Convert.ToString(Sno));
                                break;
                            }

                        }

                        for (parse = 2; parse <= 100; parse++)
                        {
                            if (xlWorksheet.Cells[parse, 1].Value != null)
                            {
                                row++; (sender as BackgroundWorker).ReportProgress(1);
                            }
                            else
                            {
                                // MessageBox.Show(Convert.ToString(Sno));
                                break;
                            }

                        }

                       
                            xlWorksheet.Cells[1, col + 3].Value = DateTime.Now.ToShortDateString(); (sender as BackgroundWorker).ReportProgress(1);
                            xlWorksheet.Column(col + 3).AutoFit(); (sender as BackgroundWorker).ReportProgress(1);
                            xlWorksheet.Cells[1, col + 3].Style.Font.Bold = true;
                            xlWorksheet.Column(col + 3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            bool isChecked = true;
                            if (enable == 1)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 2)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 3)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 4)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 5)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 6)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 7)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 8)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 9)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 10)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 11)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 12)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 13)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 14)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 15)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 16)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }

                            else if (enable == 17)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }

                            else if (enable == 18)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 18 *****/
                                if (s18P.IsChecked == false && s18A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 18 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true && s18A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 18", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s18A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }

                            else if (enable == 19)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 18 *****/
                                if (s18P.IsChecked == false && s18A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 18 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true && s18A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 18", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s18A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 19 *****/
                                if (s19P.IsChecked == false && s19A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 19 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true && s19A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 19", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s19A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 20)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 18 *****/
                                if (s18P.IsChecked == false && s18A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 18 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true && s18A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 18", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s18A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 19 *****/
                                if (s19P.IsChecked == false && s19A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 19 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true && s19A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 19", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s19A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 20 *****/
                                if (s20P.IsChecked == false && s20A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 20 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true && s20A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 20", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s20A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }

                            else if (enable == 21)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 18 *****/
                                if (s18P.IsChecked == false && s18A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 18 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true && s18A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 18", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s18A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 19 *****/
                                if (s19P.IsChecked == false && s19A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 19 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true && s19A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 19", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s19A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 20 *****/
                                if (s20P.IsChecked == false && s20A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 20 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true && s20A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 20", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s20A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 21 *****/
                                if (s21P.IsChecked == false && s21A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 21 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true && s21A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 21", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s21A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }

                            else if (enable == 22)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 18 *****/
                                if (s18P.IsChecked == false && s18A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 18 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true && s18A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 18", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s18A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 19 *****/
                                if (s19P.IsChecked == false && s19A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 19 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true && s19A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 19", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s19A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 20 *****/
                                if (s20P.IsChecked == false && s20A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 20 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true && s20A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 20", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s20A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 21 *****/
                                if (s21P.IsChecked == false && s21A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 21 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true && s21A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 21", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s21A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 22 *****/
                                if (s22P.IsChecked == false && s22A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 22 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true && s22A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 22", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s22A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }
                            else if (enable == 23)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 18 *****/
                                if (s18P.IsChecked == false && s18A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 18 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true && s18A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 18", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s18A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 19 *****/
                                if (s19P.IsChecked == false && s19A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 19 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true && s19A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 19", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s19A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 20 *****/
                                if (s20P.IsChecked == false && s20A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 20 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true && s20A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 20", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s20A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 21 *****/
                                if (s21P.IsChecked == false && s21A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 21 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true && s21A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 21", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s21A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 22 *****/
                                if (s22P.IsChecked == false && s22A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 22 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true && s22A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 22", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s22A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 23 *****/
                                if (s23P.IsChecked == false && s23A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 23 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s23P.IsChecked == true && s23A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 23", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s23P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[24, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s23A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[24, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }

                            else if (enable == 24)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 18 *****/
                                if (s18P.IsChecked == false && s18A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 18 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true && s18A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 18", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s18A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 19 *****/
                                if (s19P.IsChecked == false && s19A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 19 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true && s19A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 19", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s19A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 20 *****/
                                if (s20P.IsChecked == false && s20A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 20 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true && s20A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 20", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s20A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 21 *****/
                                if (s21P.IsChecked == false && s21A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 21 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true && s21A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 21", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s21A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 22 *****/
                                if (s22P.IsChecked == false && s22A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 22 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true && s22A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 22", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s22A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 23 *****/
                                if (s23P.IsChecked == false && s23A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 23 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s23P.IsChecked == true && s23A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 23", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s23P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[24, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s23A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[24, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 24 *****/
                                if (s24P.IsChecked == false && s24A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 24 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s24P.IsChecked == true && s24A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 24", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s24P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[25, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s24A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[25, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }

                            else if (enable == 25)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 18 *****/
                                if (s18P.IsChecked == false && s18A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 18 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true && s18A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 18", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s18A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 19 *****/
                                if (s19P.IsChecked == false && s19A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 19 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true && s19A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 19", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s19A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 20 *****/
                                if (s20P.IsChecked == false && s20A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 20 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true && s20A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 20", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s20A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 21 *****/
                                if (s21P.IsChecked == false && s21A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 21 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true && s21A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 21", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s21A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 22 *****/
                                if (s22P.IsChecked == false && s22A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 22 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true && s22A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 22", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s22A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 23 *****/
                                if (s23P.IsChecked == false && s23A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 23 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s23P.IsChecked == true && s23A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 23", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s23P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[24, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s23A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[24, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 24 *****/
                                if (s24P.IsChecked == false && s24A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 24 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s24P.IsChecked == true && s24A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 24", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s24P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[25, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s24A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[25, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 25 *****/
                                if (s25P.IsChecked == false && s25A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 25 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s25P.IsChecked == true && s25A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 25", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s25P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[26, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s25A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[26, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }

                            else if (enable == 26)
                            {
                                /****** STUDENT1*****/
                                if (s1P.IsChecked == false && s1A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 1 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s1P.IsChecked == true && s1A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 1", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                                }
                                else if (s1P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s1A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[2, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT2*****/
                                s2P.IsEnabled = true; s2A.IsEnabled = true;
                                if (s2P.IsChecked == false && s2A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 2 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s2P.IsChecked == true && s2A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 2", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s2P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s2A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[3, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT3 *****/
                                s3P.IsEnabled = true; s3A.IsEnabled = true;
                                if (s3P.IsChecked == false && s3A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 3 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s3P.IsChecked == true && s3A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 3", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s3P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s3A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[4, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT4 *****/
                                if (s4P.IsChecked == false && s4A.IsChecked == false)
                                {
                                    MessageBox.Show("Please select student 4 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    isChecked = false;
                                }
                                else if (s4P.IsChecked == true && s4A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 4", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s4P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s4A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[5, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT5 *****/
                                if (s5P.IsChecked == false && s5A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 5 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true && s5A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 5", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s5P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s5A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[6, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT6 *****/
                                if (s6P.IsChecked == false && s6A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 6 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true && s6A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 6", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s6P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s6A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[7, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT7 *****/
                                if (s7P.IsChecked == false && s7A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 7 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true && s7A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 7", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s7P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s7A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[8, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT8 *****/
                                if (s8P.IsChecked == false && s8A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 8 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true && s8A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 8", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s8P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s8A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[9, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 9 *****/
                                if (s9P.IsChecked == false && s9A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 9 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true && s9A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 9", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s9P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s9A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[10, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 10 *****/
                                if (s10P.IsChecked == false && s10A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 10 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true && s10A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 10", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s10P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s10A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[11, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 11 *****/
                                if (s11P.IsChecked == false && s11A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 11 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true && s11A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 11", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s11P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s11A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[12, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 12 *****/
                                if (s12P.IsChecked == false && s12A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 12 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true && s12A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 12", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s12P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s12A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[13, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 13 *****/
                                if (s13P.IsChecked == false && s13A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 13 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true && s13A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 13", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s13P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s13A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[14, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 14 *****/
                                if (s14P.IsChecked == false && s14A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 14 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true && s14A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 14", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s14P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s14A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[15, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 15 *****/
                                if (s15P.IsChecked == false && s15A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 15 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true && s15A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 15", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s15P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s15A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[16, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 16 *****/
                                if (s16P.IsChecked == false && s16A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 16 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true && s16A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 16", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s16P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s16A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[17, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 17 *****/
                                if (s17P.IsChecked == false && s17A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 17 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true && s17A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 17", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s17P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s17A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[18, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 18 *****/
                                if (s18P.IsChecked == false && s18A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 18 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true && s18A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 18", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s18P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s18A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[19, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 19 *****/
                                if (s19P.IsChecked == false && s19A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 19 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true && s19A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 19", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s19P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s19A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[20, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 20 *****/
                                if (s20P.IsChecked == false && s20A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 20 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true && s20A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 20", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s20P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s20A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[21, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 21 *****/
                                if (s21P.IsChecked == false && s21A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 21 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true && s21A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 21", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s21P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s21A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[22, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 22 *****/
                                if (s22P.IsChecked == false && s22A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 22 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true && s22A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 22", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s22P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s22A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[23, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 23 *****/
                                if (s23P.IsChecked == false && s23A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 23 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s23P.IsChecked == true && s23A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 23", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s23P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[24, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s23A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[24, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 24 *****/
                                if (s24P.IsChecked == false && s24A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 24 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s24P.IsChecked == true && s24A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 24", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s24P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[25, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s24A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[25, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 25 *****/
                                if (s25P.IsChecked == false && s25A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 25 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s25P.IsChecked == true && s25A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 25", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s25P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[26, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s25A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[26, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                /****** STUDENT 26 *****/
                                if (s26P.IsChecked == false && s26A.IsChecked == false)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Please select student 26 attendance ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s26P.IsChecked == true && s26A.IsChecked == true)
                                {
                                    isChecked = false;
                                    MessageBox.Show("Student cannot be present and absent at the same time \n" + "Please correct your mistake for student 26", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                else if (s26P.IsChecked == true)
                                {
                                    xlWorksheet.Cells[27, col + 3].Value = Convert.ToInt32(" 1 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                                else if (s26A.IsChecked == true)
                                {
                                    xlWorksheet.Cells[27, col + 3].Value = Convert.ToInt32(" 0 "); (sender as BackgroundWorker).ReportProgress(1);
                                }
                            }

                            /**********CHECKBOX AREA END ***************************/

                            if (isChecked == false)
                            {
                                MessageBox.Show("Attendance can not be saved\n " + "Please fill all the attendance details", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                SaveAttendanceButton.IsEnabled = true;
                                isChecked = true;
                            }
                            else
                            {
                                excelPackage.Save();

                            }

                            
                        }

                        FileInfo fileName = new FileInfo(@"C:\\Rocket\\iClass\\Attendance\\" + className + "_Class_Attendance.xlsx");
                        using (ExcelPackage Package = new ExcelPackage(fileName))
                        {
                            ExcelWorkbook WorkBook = Package.Workbook;
                            ExcelWorksheet worksheet = WorkBook.Worksheets.First();

                            double data = 0;
                            double count = 0;
                            int avg = 0;
                            int r = 2;  //used for row
                            for (int i = 2; i <= row + 1; i++)  //Column    
                            {

                                for (int j = 3; j <= col + 3; j++)  //Row
                                {
                                    if (worksheet.Cells[i, j].Value == Convert.ToString("NA"))
                                    {
                                        
                                    }
                                    else
                                    {
                                        data += Convert.ToDouble(worksheet.Cells[i, j].Value); (sender as BackgroundWorker).ReportProgress(1);
                                        count++; (sender as BackgroundWorker).ReportProgress(1);
                                    }

                                }
                                avg = Convert.ToInt32((data / count) * 100); (sender as BackgroundWorker).ReportProgress(1);
                                worksheet.Cells[r, 2].Value = null;
                                worksheet.Cells[r, 2].Value = avg; (sender as BackgroundWorker).ReportProgress(1);
                                r = r + 1; //we have saved first student attendance so increment row to save next student.
                                data = 0;    //clear data contents
                                avg = 0;
                                count = 0;
                            }
                            //worksheet.Cells[10, 10].Value = "Hello";

                            Package.Save();
                        }
                        Log("Attendace saved successfully");
                            MessageBox.Show("Attendance saved successfully     ", "Congratulations", MessageBoxButton.OK, MessageBoxImage.Information);
                            attendanceLabel.Content = " Attendace Saved";
                            attendanceCircle.Fill = Brushes.Green;
                            attendancePath.Fill = Brushes.Green;
                            attendancePath.Stroke = Brushes.Green;
                            attendanceCircle.Stroke = Brushes.Green;
                    }
                    catch (Exception ex)
                    {
                         MessageBox.Show("Exception occured during attendance save.\n " + Convert.ToString(ex), "Exception");
                    }


                });
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                //progress.Show();
            }
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Mouse.OverrideCursor = null;
            //progress.Close();
            
        }

        bool connected = true;
        private void Check_Network_Status(object obj)
        {
            while (true)
            {
                
                connected = System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();
                if (connected == true)
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        networkCircle.Fill = Brushes.Green;
                        networkCircle.Stroke = Brushes.Green;
                        networkPath.Fill = Brushes.Green;
                        networkPath.Stroke = Brushes.Green;
                        networkLabel.Content = "Network Stable";
                        emailSendButton.IsEnabled = true;

                    });

                }
                if (connected == false)
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        networkCircle.Fill = Brushes.Red;
                        networkCircle.Stroke = Brushes.Red;
                        networkPath.Fill = Brushes.Red;
                        networkPath.Stroke = Brushes.Red;
                        networkLabel.Content = "No Network ";
                        emailSendButton.IsEnabled = false;
                    });
                }
                //Thread.Sleep(1000);
            }
        }




        void timer_Tick(object sender, EventArgs e)
        {
            currentTimeTextBox.Text = DateTime.Now.ToLongTimeString();
            currentDateTextBox.Text = DateTime.Now.ToShortDateString();
        }

        void function()
        {
            /*
             * This function reads the class file and fetch all the students name.
             * Show all the student name in text box in start class window
             * Also enable only those text box which has data in them. and disables rest of the text box.
             * It also extracts email ids.
             */
            FileInfo file = new FileInfo(@"C:\\Rocket\\iClass\\Class\\" + className + ".xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet worksheet = excelWorkBook.Worksheets.First();

                teacherNameLabel.Content = worksheet.Cells[2, 5].Value;

                
                
                s1NameTextBox.Text = Convert.ToString(worksheet.Cells[2, 2].Value);
                s1email = Convert.ToString(worksheet.Cells[2, 4].Value);

                s2NameTextBox.Text = Convert.ToString(worksheet.Cells[3, 2].Value);
                s2email = Convert.ToString(worksheet.Cells[3, 4].Value);

                s3NameTextBox.Text = Convert.ToString(worksheet.Cells[4, 2].Value);
                s3email = Convert.ToString(worksheet.Cells[4, 4].Value);

                s4NameTextBox.Text = Convert.ToString(worksheet.Cells[5, 2].Value);
                s4email = Convert.ToString(worksheet.Cells[5, 4].Value);

                s5NameTextBox.Text = Convert.ToString(worksheet.Cells[6, 2].Value);
                s5email = Convert.ToString(worksheet.Cells[6, 4].Value);

                s6NameTextBox.Text = Convert.ToString(worksheet.Cells[7, 2].Value);
                s6email = Convert.ToString(worksheet.Cells[7, 4].Value);

                s7NameTextBox.Text = Convert.ToString(worksheet.Cells[8, 2].Value);
                s7email = Convert.ToString(worksheet.Cells[8, 4].Value);

                s8NameTextBox.Text = Convert.ToString(worksheet.Cells[9, 2].Value);
                s8email = Convert.ToString(worksheet.Cells[9, 4].Value);

                s9NameTextBox.Text = Convert.ToString(worksheet.Cells[10, 2].Value);
                s9email = Convert.ToString(worksheet.Cells[10, 4].Value);

                s10NameTextBox.Text = Convert.ToString(worksheet.Cells[11, 2].Value);
                s10email = Convert.ToString(worksheet.Cells[11, 4].Value);

                s11NameTextBox.Text = Convert.ToString(worksheet.Cells[12, 2].Value);
                s11email = Convert.ToString(worksheet.Cells[12, 4].Value);

                s12NameTextBox.Text = Convert.ToString(worksheet.Cells[13, 2].Value);
                s12email = Convert.ToString(worksheet.Cells[13, 4].Value);

                s13NameTextBox.Text = Convert.ToString(worksheet.Cells[14, 2].Value);
                s13email = Convert.ToString(worksheet.Cells[14, 4].Value);

                s14NameTextBox.Text = Convert.ToString(worksheet.Cells[15, 2].Value);
                s14email = Convert.ToString(worksheet.Cells[15, 4].Value);

                s15NameTextBox.Text = Convert.ToString(worksheet.Cells[16, 2].Value);
                s15email = Convert.ToString(worksheet.Cells[16, 4].Value);

                s16NameTextBox.Text = Convert.ToString(worksheet.Cells[17, 2].Value);
                s16email = Convert.ToString(worksheet.Cells[17, 4].Value);

                s17NameTextBox.Text = Convert.ToString(worksheet.Cells[18, 2].Value);
                s17email = Convert.ToString(worksheet.Cells[18, 4].Value);

                s18NameTextBox.Text = Convert.ToString(worksheet.Cells[19, 2].Value);
                s18email = Convert.ToString(worksheet.Cells[19, 4].Value);

                s19NameTextBox.Text = Convert.ToString(worksheet.Cells[20, 2].Value);
                s19email = Convert.ToString(worksheet.Cells[20, 4].Value);

                s20NameTextBox.Text = Convert.ToString(worksheet.Cells[21, 2].Value);
                s20email = Convert.ToString(worksheet.Cells[21, 4].Value);

                s21NameTextBox.Text = Convert.ToString(worksheet.Cells[22, 2].Value);
                s21email = Convert.ToString(worksheet.Cells[22, 4].Value);

                s22NameTextBox.Text = Convert.ToString(worksheet.Cells[23, 2].Value);
                s22email = Convert.ToString(worksheet.Cells[23, 4].Value);

                s23NameTextBox.Text = Convert.ToString(worksheet.Cells[24, 2].Value);
                s23email = Convert.ToString(worksheet.Cells[24, 4].Value);

                s24NameTextBox.Text = Convert.ToString(worksheet.Cells[25, 2].Value);
                s24email = Convert.ToString(worksheet.Cells[25, 4].Value);

                s25NameTextBox.Text = Convert.ToString(worksheet.Cells[26, 2].Value);
                s25email = Convert.ToString(worksheet.Cells[26, 4].Value);

                s26NameTextBox.Text = Convert.ToString(worksheet.Cells[27, 2].Value);
                s26email = Convert.ToString(worksheet.Cells[27, 4].Value);

                int rowCnt = 0;
                for (parse = 2; parse <= 100; parse++)
                {
                    if (worksheet.Cells[parse, 1].Value != null)
                    {
                        rowCnt++; 
                    }
                    else
                    {
                        
                        break; 
                    }

                }



                if (rowCnt == 1)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    enable = 1;
                }
                else if (rowCnt == 2)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    enable = 2;
                }
                else if (rowCnt == 3)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    enable = 3;
                }
                else if (rowCnt == 4)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    enable = 4;
                }
                else if (rowCnt == 5)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    enable = 5;
                }
                else if (rowCnt == 6)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    enable = 6;
                }
                else if (rowCnt == 7)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    enable = 7;
                }
                else if (rowCnt == 8)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    enable = 8;
                }
                else if (rowCnt == 9)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    enable = 9;
                }
                else if (rowCnt == 10)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    enable = 10;
                }
                else if (rowCnt == 11)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    enable = 11;
                }
                else if (rowCnt == 12)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    enable = 12;
                }
                else if (rowCnt == 13)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    enable = 13;
                }
                else if (rowCnt == 14)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    enable = 14;
                }
                else if (rowCnt == 15)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    enable = 15;
                }
                else if (rowCnt == 16)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    enable = 16;
                }
                else if (rowCnt == 17)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    enable = 17;
                }
                else if (rowCnt == 18)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    s18P.IsEnabled = true; s18A.IsEnabled = true;
                    enable = 18;
                }
                else if (rowCnt == 19)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    s18P.IsEnabled = true; s18A.IsEnabled = true;
                    s19P.IsEnabled = true; s19A.IsEnabled = true;
                    enable = 19;
                }
                else if (rowCnt == 20)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    s18P.IsEnabled = true; s18A.IsEnabled = true;
                    s19P.IsEnabled = true; s19A.IsEnabled = true;
                    s20P.IsEnabled = true; s20A.IsEnabled = true;
                    enable = 20;
                }
                else if (rowCnt == 21)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    s18P.IsEnabled = true; s18A.IsEnabled = true;
                    s19P.IsEnabled = true; s19A.IsEnabled = true;
                    s20P.IsEnabled = true; s20A.IsEnabled = true;
                    s21P.IsEnabled = true; s21A.IsEnabled = true;
                    enable = 21;
                }
                else if (rowCnt == 22)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    s18P.IsEnabled = true; s18A.IsEnabled = true;
                    s19P.IsEnabled = true; s19A.IsEnabled = true;
                    s20P.IsEnabled = true; s20A.IsEnabled = true;
                    s21P.IsEnabled = true; s21A.IsEnabled = true;
                    s22P.IsEnabled = true; s22A.IsEnabled = true;
                    enable = 22;
                }
                else if (rowCnt == 23)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    s18P.IsEnabled = true; s18A.IsEnabled = true;
                    s19P.IsEnabled = true; s19A.IsEnabled = true;
                    s20P.IsEnabled = true; s20A.IsEnabled = true;
                    s21P.IsEnabled = true; s21A.IsEnabled = true;
                    s22P.IsEnabled = true; s22A.IsEnabled = true;
                    s23P.IsEnabled = true; s23A.IsEnabled = true;
                    enable = 23;
                }
                else if (rowCnt == 24)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    s18P.IsEnabled = true; s18A.IsEnabled = true;
                    s19P.IsEnabled = true; s19A.IsEnabled = true;
                    s20P.IsEnabled = true; s20A.IsEnabled = true;
                    s21P.IsEnabled = true; s21A.IsEnabled = true;
                    s22P.IsEnabled = true; s22A.IsEnabled = true;
                    s23P.IsEnabled = true; s23A.IsEnabled = true;
                    s24P.IsEnabled = true; s24A.IsEnabled = true;
                    enable = 24;
                }
                else if (rowCnt == 25)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    s18P.IsEnabled = true; s18A.IsEnabled = true;
                    s19P.IsEnabled = true; s19A.IsEnabled = true;
                    s20P.IsEnabled = true; s20A.IsEnabled = true;
                    s21P.IsEnabled = true; s21A.IsEnabled = true;
                    s22P.IsEnabled = true; s22A.IsEnabled = true;
                    s23P.IsEnabled = true; s23A.IsEnabled = true;
                    s24P.IsEnabled = true; s24A.IsEnabled = true;
                    s25P.IsEnabled = true; s25A.IsEnabled = true;
                    enable = 25;
                }
                else if (rowCnt == 26)
                {
                    s1P.IsEnabled = true; s1A.IsEnabled = true;
                    s2P.IsEnabled = true; s2A.IsEnabled = true;
                    s3P.IsEnabled = true; s3A.IsEnabled = true;
                    s4P.IsEnabled = true; s4A.IsEnabled = true;
                    s5P.IsEnabled = true; s5A.IsEnabled = true;
                    s6P.IsEnabled = true; s6A.IsEnabled = true;
                    s7P.IsEnabled = true; s7A.IsEnabled = true;
                    s8P.IsEnabled = true; s8A.IsEnabled = true;
                    s9P.IsEnabled = true; s9A.IsEnabled = true;
                    s10P.IsEnabled = true; s10A.IsEnabled = true;
                    s11P.IsEnabled = true; s11A.IsEnabled = true;
                    s12P.IsEnabled = true; s12A.IsEnabled = true;
                    s13P.IsEnabled = true; s13A.IsEnabled = true;
                    s14P.IsEnabled = true; s14A.IsEnabled = true;
                    s15P.IsEnabled = true; s15A.IsEnabled = true;
                    s16P.IsEnabled = true; s16A.IsEnabled = true;
                    s17P.IsEnabled = true; s17A.IsEnabled = true;
                    s18P.IsEnabled = true; s18A.IsEnabled = true;
                    s19P.IsEnabled = true; s19A.IsEnabled = true;
                    s20P.IsEnabled = true; s20A.IsEnabled = true;
                    s21P.IsEnabled = true; s21A.IsEnabled = true;
                    s22P.IsEnabled = true; s22A.IsEnabled = true;
                    s23P.IsEnabled = true; s23A.IsEnabled = true;
                    s24P.IsEnabled = true; s24A.IsEnabled = true;
                    s25P.IsEnabled = true; s25A.IsEnabled = true;
                    s26P.IsEnabled = true; s26A.IsEnabled = true;
                    enable = 26;
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
