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
using System.Threading;
using Microsoft.Win32;
using System.Net.Mail;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using WPFCustomMessageBox;


namespace iClass
{
    /// <summary>
    /// Interaction logic for Login.xaml
    /// </summary>
    public partial class Login : Window
    {
        MailMessage mail = new MailMessage();
        string id;
        CircularProgressBar progress = new CircularProgressBar();
        
        public Login()
        {
            InitializeComponent();
            
            
            //Thread statusThread = new Thread(Check_registration_status);
            //statusThread.IsBackground = true;
            //statusThread.Start();
            
            createFolders();
            checkID();
            statusShow();
            Log("********************* Application Started on *********************");
                        
        }

        private void createFolders()
        {
            if (!Directory.Exists(@"C:\\Rocket\\"))
            {
                Directory.CreateDirectory(@"C:\\Rocket\\");
                Directory.CreateDirectory(@"C:\\Rocket\\iClass\\");
                Directory.CreateDirectory(@"C:\\Rocket\\iClass\\Attendance\\");
                Directory.CreateDirectory(@"C:\\Rocket\\iClass\\Teacher\\");
                Directory.CreateDirectory(@"C:\\Rocket\\iClass\\Class\\");
                Directory.CreateDirectory(@"C:\\Rocket\\iClass\\Logs\\");
                Directory.CreateDirectory(@"C:\\Rocket\\iClass\\Backup\\");
                Directory.CreateDirectory(@"C:\\Rocket\\iClass\\Files\\");
                
            }
        }

        void statusShow()
        {
            RegistryKey sReg = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Node32");
            if (sReg != null)
            {
                string value = Convert.ToString(sReg.GetValue("KEY"));
                if (value == "1")
                {
                    userIdTextBox.IsEnabled = true;
                    passwordTextBox.IsEnabled = true;
                    

                }
                else if (value == "0")
                {
                    userIdTextBox.IsEnabled = false;
                    passwordTextBox.IsEnabled = false;
                    
                }
                else if (value == "2")
                {
                    userIdTextBox.IsEnabled = true;
                    passwordTextBox.IsEnabled = true;
                    
                }
            }
            else
            {
                userIdTextBox.IsEnabled = false;
                passwordTextBox.IsEnabled = false;
                
            }
        }

       /* void Check_registration_status()
        {
            while (true)
            {
                this.Dispatcher.Invoke(() =>
                {
                    RegistryKey sReg = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Node32");
                    if (sReg != null)
                    {
                        string value = Convert.ToString(sReg.GetValue("KEY"));
                        if (value == "1")
                        {
                            userIdTextBox.IsEnabled = true;
                            passwordTextBox.IsEnabled = true;
                            
                        }
                        else if (value == "0")
                        {
                            userIdTextBox.IsEnabled = false;
                            passwordTextBox.IsEnabled = false;
                            
                        }
                    }
                    else
                    {
                        userIdTextBox.IsEnabled = false;
                        passwordTextBox.IsEnabled = false;
                        
                    }
                });
            }
        }*/


        private void SignIn_buttonClick(object sender, RoutedEventArgs e)
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
                    
                    Mouse.OverrideCursor = Cursors.Wait;
                    (sender as BackgroundWorker).ReportProgress(1);
                    bool isConnected = CheckForInternetConnection();
                    if (isConnected == true)
                    {
                        
                        RegistryKey sReg = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Node32");
                        if (sReg != null)
                        {
                            
                            string value = Convert.ToString(sReg.GetValue("KEY"));
                            DateTime end = Convert.ToDateTime(sReg.GetValue("END"));
                            if (value == "1")
                            {
                                Log("Registration Plan  :   30 DAYS TRIAL PERIOD");
                                DateTime today = GetNetworkTime();
                                Log("NT Time    :   " + today);
                                TimeSpan diff = end - today;
                                if(end < today)
                                {
                                    RegistryKey cReg = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Node32");
                                    cReg.SetValue("KEY", "0");
                                    userIdTextBox.IsEnabled = false;
                                    passwordTextBox.IsEnabled = false;
                                    signInButton.IsEnabled = false;
                                    Log("Trial Period Status    :   DEACTIVATED");
                                    MessageBox.Show("You have used your trial period. Please purchase license", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                                    
                                }
                                else
                                {
                                    if (userIdTextBox.Text == "admin" && passwordTextBox.Password == "admin")
                                    {
                                        MessageBox.Show("You are using trial version which will expire on " + end, "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                                        Log("Log in success on " + DateTime.Now.ToString("dd-M-yyyy-hh-mm-ss"));
                                        Log("Dashboard Called");
                                        MainWindow window = new MainWindow();
                                        window.Show();
                                        this.Close();

                                    }
                                    else
                                    {
                                        Log("Wrong credentials");
                                        MessageBox.Show("Please enter correct username and password", "Wrong Credentials", MessageBoxButton.OK, MessageBoxImage.Error);
                                    }
                                }
                            }
                            else if (value == "2")
                            {
                                Log("Registration Plan  :   1 Year");
                                DateTime today = GetNetworkTime();
                                Log("NT Time    :   " + today);
                                //TimeSpan diff = end - today;
                                if (end < today)
                                {
                                    RegistryKey cReg = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Node32");
                                    cReg.SetValue("KEY", "0");
                                    userIdTextBox.IsEnabled = false;
                                    passwordTextBox.IsEnabled = false;
                                    signInButton.IsEnabled = false;
                                    Log("1 Year License Period Status    :   DEACTIVATED");
                                    MessageBox.Show("You have used your License period. Please purchase license", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                                }
                                else
                                {
                                    string path = @"C:\\Rocket\\iClass\\Files\\license.dll";
                                    if (File.Exists(path) == true)
                                    {
                                        try
                                        {
                                            var DLL = Assembly.LoadFile(path);
                                            var class1Type = DLL.GetType("license.Class1");
                                            dynamic license = Activator.CreateInstance(class1Type);

                                            if (license.CheckProductID(id) == true)
                                            {
                                                if (license.CheckUserPassword(userIdTextBox.Text, passwordTextBox.Password) == true)
                                                {
                                                    MessageBox.Show("1 Year license version activated till " + end, "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                                                    Log("Log in success on " + DateTime.Now.ToString("dd-M-yyyy-hh-mm-ss"));
                                                    Log("Dashboard Called");
                                                    MainWindow window = new MainWindow();
                                                    window.Show();
                                                    this.Close();
                                                }
                                                else
                                                {
                                                    Log("Wrong credentials");
                                                    MessageBox.Show("Please enter correct username and password", "Wrong Credentials", MessageBoxButton.OK, MessageBoxImage.Error);
                                                }
                                            }
                                            else
                                            {
                                                Log("Wrong license file detected");
                                                MessageBox.Show("Wrong license file detected", "Licensing Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                            }
                                        }
                                        catch (Exception exception)
                                        {
                                            MessageBox.Show(Convert.ToString(exception));
                                        }
                                    }
                                    else
                                    {
                                        Log("License File not found");
                                        MessageBox.Show("License File not found", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    }
                                }
                            }
                            else if (value == "0")
                            {
                                Log("Trial Period Status    :   DEACTIVATED");
                                MessageBox.Show("Your license or trial period has expired. Please purchase license to continue using this software", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            }
                        }
                        else
                        {
                            //If registry not found, that means user has not registered the software
                            Log("Registration Status    :   NOT REGISTERED");
                            MessageBox.Show("Please register to continue using this software", "Information", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        }
                    }
                    else
                    {
                        Log("Internet Connection    :   NOT ACTIVE");
                        MessageBox.Show("No active internet connection found. You need stable connection to login", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
            progress.Close();
        }

        private void ProductInfo_ButtonClick(object sender, RoutedEventArgs e)
        {
            RegistrationWindow register = new RegistrationWindow();
            register.Show();
        }

        void checkID()
        {
            ManagementObject dsk = new ManagementObject(@"win32_logicaldisk.deviceid=""c:""");
            dsk.Get();
            id = dsk["VolumeSerialNumber"].ToString();
            Log("Product ID : " + id);

        }

        public static DateTime GetNetworkTime()
        {
            //default Windows time server
            const string ntpServer = "time.windows.com";

            // NTP message size - 16 bytes of the digest (RFC 2030)
            var ntpData = new byte[48];

            //Setting the Leap Indicator, Version Number and Mode values
            ntpData[0] = 0x1B; //LI = 0 (no warning), VN = 3 (IPv4 only), Mode = 3 (Client Mode)

            var addresses = Dns.GetHostEntry(ntpServer).AddressList;

            //The UDP port number assigned to NTP is 123
            var ipEndPoint = new IPEndPoint(addresses[0], 123);
            //NTP uses UDP
            var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);

            socket.Connect(ipEndPoint);

            //Stops code hang if NTP is blocked
            socket.ReceiveTimeout = 3000;

            socket.Send(ntpData);
            socket.Receive(ntpData);
            socket.Close();

            //Offset to get to the "Transmit Timestamp" field (time at which the reply 
            //departed the server for the client, in 64-bit timestamp format."
            const byte serverReplyTime = 40;

            //Get the seconds part
            ulong intPart = BitConverter.ToUInt32(ntpData, serverReplyTime);

            //Get the seconds fraction
            ulong fractPart = BitConverter.ToUInt32(ntpData, serverReplyTime + 4);

            //Convert From big-endian to little-endian
            intPart = SwapEndianness(intPart);
            fractPart = SwapEndianness(fractPart);

            var milliseconds = (intPart * 1000) + ((fractPart * 1000) / 0x100000000L);

            //**UTC** time
            var networkDateTime = (new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc)).AddMilliseconds((long)milliseconds);

            return Convert.ToDateTime(networkDateTime.ToShortDateString());
        }


        static uint SwapEndianness(ulong x)
        {
            return (uint)(((x & 0x000000ff) << 24) +
                           ((x & 0x0000ff00) << 8) +
                           ((x & 0x00ff0000) >> 8) +
                           ((x & 0xff000000) >> 24));
        }

        public static bool CheckForInternetConnection()
        {
            //return System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();
            try
            {
                using (var client = new WebClient())
                {
                    using (var stream = client.OpenRead("http://www.google.com"))
                    {
                      
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        void windowClose(object sender, CancelEventArgs e)
        {
            bool wasCodeClosed = new StackTrace().GetFrames().FirstOrDefault(x => x.GetMethod() == typeof(Window).GetMethod("Close")) != null;
            if (wasCodeClosed)
            {
                // Closed with this.Close()
            }
            else
            {
                System.Windows.Application curApp = System.Windows.Application.Current;
                curApp.Shutdown();
            }

            //base.OnClosing(e);
   
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
