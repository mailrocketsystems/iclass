using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    /// Interaction logic for StartClassPopUp.xaml
    /// </summary>
    public partial class StartClassPopUp : Window
    {
        Array className;
        CircularProgressBar progress = new CircularProgressBar();
        public StartClassPopUp(Array data)
        {
            InitializeComponent();
            //Log("Start Class pop up Active Now");
            System.Media.SystemSounds.Exclamation.Play();
            className = data;
            selectClassComboBox.Items.Clear();
            selectClassComboBox.ItemsSource = className;
            
        }
        string classData = null;
        private void StartClass_ButtonClick(object sender, RoutedEventArgs e)
        {
            StartClassButton.IsEnabled = false;
            classData = selectClassComboBox.Text;
            if (string.IsNullOrWhiteSpace(classData))
            {
                Log("Error  :   Class not selected");
                System.Media.SystemSounds.Hand.Play();
                MessageBox.Show("Please select a class   ", "Error ", MessageBoxButton.OK, MessageBoxImage.Error);
                StartClassButton.IsEnabled = true;
                //this.Close();
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
                    Log("Start Class button clicked");
                    progress.Show();
                    Mouse.OverrideCursor = Cursors.Wait;
                    classData = classData.Remove(classData.Length - 5); (sender as BackgroundWorker).ReportProgress(1);
                    Log("Start class window called");
                    StartClassWindow startClass = new StartClassWindow(classData); (sender as BackgroundWorker).ReportProgress(1);
                    startClass.Show(); (sender as BackgroundWorker).ReportProgress(0);
                    this.Close();
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
            StartClassButton.IsEnabled = true;
            progress.Close();
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
