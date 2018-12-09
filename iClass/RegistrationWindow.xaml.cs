using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Mail;
using System.Net.Sockets;
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


namespace iClass
{
    /// <summary>
    /// Interaction logic for RegistrationWindow.xaml
    /// </summary>
    public partial class RegistrationWindow : Window
    {
        string id;
        
        public RegistrationWindow()
        {
            InitializeComponent();
            checkID();
            checkStatus();
        }

        private void Trial_ButtonClick(object sender, RoutedEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Wait;
            bool isConnected = CheckForInternetConnection();
            if (isConnected == true)
            {
                //First check if registry is already created or not.
                //If registry is already created(in case of trial version), show blacklisted error message 
                RegistryKey oReg = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Node32");
                if (oReg != null)
                {
                    MessageBox.Show("You have used your trial period. Please purchase license to continue using this software", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {

                    if (string.IsNullOrWhiteSpace(fullNameTextBox.Text) || string.IsNullOrWhiteSpace(emailIdTextBox.Text) || string.IsNullOrWhiteSpace(phoneNumberTextBox.Text))
                    {
                        MessageBox.Show("Please fill in all details to register the software", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        trialButton.IsEnabled = true;
                    }
                    else
                    {

                        DateTime endDate = GetNetworkTime();
                        endDate = endDate.AddDays(30);
                        RegistryKey cReg = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Node32");
                        cReg.SetValue("KEY", "1");
                        cReg.SetValue("END", endDate);

                        SmtpClient SmtpServer = new SmtpClient("smtp.zoho.com");
                         MailMessage mail = new MailMessage();
                         mail.From = new MailAddress("rocketeducation@zoho.com");
                         mail.To.Add("abhinavrawat92@gmail.com");
                         mail.Subject = "New User Registration Information";
                         mail.Body = "New user has requested for the trial version of the iClass\n\n" + "Following are the details of user: \n\n" + "Name: " + fullNameTextBox.Text + "\n" + "Email ID: " + emailIdTextBox.Text + "\n" + "Phone Number: " + phoneNumberTextBox.Text + "\n" + "Product Key: " + id + "\n" + "Last Date: " + endDate;
                         SmtpServer.Port = 587;
                         SmtpServer.Credentials = new System.Net.NetworkCredential("rocketeducation", "Truthisgod@27");
                         SmtpServer.EnableSsl = true;
                         SmtpServer.Send(mail);

                        productActivationStatusTextBox.Text = "Trial Version. Activated till " + endDate;
                        MessageBox.Show("Product activated for 30 days. Expiry date: " + endDate + "\n" + "Default User ID & Password for trail period is admin\n" + "Please restart the application to activate it", "Congratulations", MessageBoxButton.OK, MessageBoxImage.Information);
                        trialButton.IsEnabled = false;

                        System.Windows.Application curApp1 = System.Windows.Application.Current;
                        curApp1.Shutdown();
                    }

                }
            }
            else
            {
                MessageBox.Show("No active internet connection found. You need stable internet connection for registration", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            Mouse.OverrideCursor = null;
        }

        private void License_ButtonClick(object sender, RoutedEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Wait;
            if (string.IsNullOrWhiteSpace(fullNameTextBox.Text) || string.IsNullOrWhiteSpace(emailIdTextBox.Text) || string.IsNullOrWhiteSpace(phoneNumberTextBox.Text))
            {
                MessageBox.Show("Please fill in all details to register the software", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                
            }
            else
            {
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                Nullable<bool> result = dlg.ShowDialog();
                if (result == true)
                {

                    string filename = dlg.FileName;

                    System.IO.Directory.CreateDirectory(@"C:\Rocket\iClass\Files\");

                    if (File.Exists(@"C:\Rocket\iClass\Files\license.dll") == true)
                    {
                        System.IO.File.Delete(@"C:\Rocket\iClass\Files\license.dll");
                    }

                   System.IO.File.Copy(filename, @"C:\Rocket\iClass\Files\" + System.IO.Path.GetFileName(filename));
                    
                }

                bool isConnected = CheckForInternetConnection();
                if (isConnected == true)
                {
                    DateTime endDate = GetNetworkTime();
                    endDate = endDate.AddDays(365);
                    RegistryKey cReg = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Node32");
                    cReg.SetValue("KEY", "2");
                    cReg.SetValue("END", endDate);

                    SmtpClient SmtpServer = new SmtpClient("smtp.zoho.com");
                    MailMessage mail = new MailMessage();
                    mail.From = new MailAddress("rocketeducation@zoho.com");
                    mail.To.Add("abhinavrawat92@gmail.com");
                    mail.Subject = "User License Information";
                    mail.Body = "User has successfully installed the license file\n\n" + "Following are the details of user: \n\n" + "Name: " + fullNameTextBox.Text + "\n" + "Email ID: " + emailIdTextBox.Text + "\n" + "Phone Number: " + phoneNumberTextBox.Text + "\n" + "Product Key: " + id + "\n" + "Last Date: " + endDate;
                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("rocketeducation", "Truthisgod@27");
                    SmtpServer.EnableSsl = true;
                    SmtpServer.Send(mail);

                    productActivationStatusTextBox.Text = "Product Activated till " + endDate;
                    MessageBox.Show("Product activated for 1 year. Expiry date: " + endDate + "\n" + "Please restart the application to activate it", "Congratulations", MessageBoxButton.OK, MessageBoxImage.Information);
                    licenseButton.IsEnabled = false;

                    System.Windows.Application curApp1 = System.Windows.Application.Current;
                    curApp1.Shutdown();
                    Mouse.OverrideCursor = null;


                }
            }
            Mouse.OverrideCursor = null;

        }

        void checkID()
        {
            ManagementObject dsk = new ManagementObject(@"win32_logicaldisk.deviceid=""c:""");
            dsk.Get();
             id = dsk["VolumeSerialNumber"].ToString();

            productNameTextBox.Text = "iClass, Class Management Application";
            productKeyTextBox.Text = id;
        }

        void checkStatus()
        {
            RegistryKey sReg = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Node32");
            if (sReg != null)
            {
                string value = Convert.ToString(sReg.GetValue("KEY"));
                DateTime end = Convert.ToDateTime(sReg.GetValue("END"));
                
                if (value == "1")
                {
                    productActivationStatusTextBox.Text = "Trial Version. Activated till " + end;
                    trialButton.IsEnabled = false;

                }
                else if (value == "0")
                {
                    productActivationStatusTextBox.Text = "Deactivated. Please purchase license " ;
                    trialButton.IsEnabled = false;
                    licenseButton.IsEnabled = true;
                }
                else if (value == "2")
                {
                    productActivationStatusTextBox.Text = "Product Activated till  " + end;
                    trialButton.IsEnabled = false;
                    licenseButton.IsEnabled = false;
                }
            }
            else 
            {
                productActivationStatusTextBox.Text = "Not Registered ";
            }
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
           // return System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();
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
