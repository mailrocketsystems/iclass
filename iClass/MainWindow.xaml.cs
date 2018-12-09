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
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace iClass
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        CircularProgressBar progress = new CircularProgressBar();
        //System.Windows.Application currApp = new System.Windows.Application();
        public MainWindow()
        {
            InitializeComponent();
            //Log("Dashboard  :   Active");
            
        }

        /**********************************************/
        /******* CREATE CLASS BUTTON ******************/
        /*********************************************/
        string[] teacherName = new string[100];
        private void CreateClass_buttonClick(object sender, RoutedEventArgs e)
        {
            CreateClassButton.IsEnabled = false;
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync(10000);
            
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            Log("Create Class Button Clicked");
            int count = 0; (sender as BackgroundWorker).ReportProgress(1);
            FileInfo file = new FileInfo(@"C:\\Rocket\\iClass\\Teacher\\TeacherDetails.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet worksheet = excelWorkBook.Worksheets.First();
                
                /* Saving all the updated details of teacher in teacherName */
                for (int parse = 2; parse <= 100; parse++)
                {

                    if (worksheet.Cells[parse, 1].Value != null)
                    {
                        teacherName[count] = Convert.ToString(worksheet.Cells[parse, 2].Value); (sender as BackgroundWorker).ReportProgress(1);
                        count = count + 1; (sender as BackgroundWorker).ReportProgress(1);
                    }
                    else
                    {
                        break; 
                    }

                }
            }
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
            CreateClassButton.IsEnabled = true;
            CreateClassPopUp createClassPopUp = new CreateClassPopUp(teacherName);
            createClassPopUp.Show();
            
            progress.Visibility = Visibility.Hidden;        //we cannot close the progress window because it will be used again so visiblity is hidden. 
            
        }
        /**********************************************/
        /******* CREATE CLASS BUTTON ******************/
        /*********************************************/


        /* Shut Down the application when user clicks red cross button*/
        void DataWindow_Closing(object sender, CancelEventArgs e)
        {
            Log("********************* Application Shut down on *********************");
            System.Windows.Application curApp = System.Windows.Application.Current;
            curApp.Shutdown();
           
            
        }


        
        private void AddTeacher_ButtonClick(object sender, RoutedEventArgs e)
        {
            AddTeacherPopUp addTeacherPopUp = new AddTeacherPopUp();
            addTeacherPopUp.Show();
        }


        //**************START CLASS*********************//
        private void StartClass_ButtonClick(object sender, RoutedEventArgs e)
        {
            Log("Start Class button clicked     :   ");
            StartClassButton.IsEnabled = false;
            Array allClass;
            DirectoryInfo d = new DirectoryInfo(@"C:\\Rocket\\iClass\\Class");
            FileInfo[] Files = d.GetFiles("*.xlsx");
            allClass = Files;
            StartClassButton.IsEnabled = true;
            StartClassPopUp startClassPopUp = new StartClassPopUp(allClass);
            Log("Start Class pop up called");
            startClassPopUp.Show();
        }

        /************* VIEW CLASS *************************/
        private void ViewClass_buttonClick(object sender, RoutedEventArgs e)
        {
            
            Array allClass;
            //string allClass = string.Empty;
            DirectoryInfo d = new DirectoryInfo(@"C:\\Rocket\\iClass\\Class");
            FileInfo[] Files = d.GetFiles("*.xlsx");
            allClass = Files;
            
            ViewClassPopup viewClassPopUp = new ViewClassPopup(allClass);
            viewClassPopUp.Show();
        }

        /************* Delete CLASS *************************/
        private void DeleteClass_buttonClick(object sender, RoutedEventArgs e)
        {
            DeleteClassButton.IsEnabled = false;
            Array allClass;
            //string allClass = string.Empty;
            DirectoryInfo d = new DirectoryInfo(@"C:\\Rocket\\iClass\\Class");
            FileInfo[] Files = d.GetFiles("*.xlsx");
            allClass = Files;
            DeleteClassButton.IsEnabled = true;
            DeleteClassPopUp deleteClass = new DeleteClassPopUp(allClass);
            deleteClass.Show();
        }

        private void UpdateClass_ButtonClick(object sender, RoutedEventArgs e)
        {
            UpdateClassPopUp updateWindow = new UpdateClassPopUp();
            updateWindow.Show();
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
