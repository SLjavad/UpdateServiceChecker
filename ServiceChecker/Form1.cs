using IWshRuntimeLibrary;
using ServiceChecker.Properties;
using System;
using System.Drawing;
using System.IO;
using System.Management;
using System.Reflection;
using System.ServiceProcess;
using System.Windows.Forms;

namespace ServiceChecker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        ServiceController service = new ServiceController("Windows Update");
        string path = "Win32_Service.Name='" + "wuauserv" + "'";
        ManagementPath managementPath;
        ManagementObject managementObject;
        Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
        Assembly curAssembly = Assembly.GetExecutingAssembly();

        // Custom Functions
        private void UpdateText(string MainText , string startMode , Color colorMain , Color startModeColor)
        {
            label1.ForeColor = colorMain;
            label2.ForeColor = startModeColor;
            label1.Text = MainText;
            label2.Text = "StartMode is : " + startMode;
        }
        private void CreateShortcut(string shortcutName, string shortcutPath, string targetFileLocation)
        {
            string shortcutLocation = System.IO.Path.Combine(shortcutPath, shortcutName + ".lnk");
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);

            shortcut.Description = "Windows Update Checker Shortcut"; // The description of the shortcut
            shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
            shortcut.Save();                                          // Save the shortcut
        }

        // System Functions
        private void button_StartModeClick(object sender , EventArgs e)
        {
            if (managementObject["StartMode"].ToString() == "Auto" || managementObject["StartMode"].ToString() == "Manual")
            {
                object[] parameters = new object[1] { "Disabled" };
                managementObject.InvokeMethod("ChangeStartMode", parameters);
                managementObject = new ManagementObject(managementPath);
                UpdateText("Good , Windows Update has been Stopped", managementObject["StartMode"].ToString(), Color.Green, Color.Green);
                button1.Enabled = false;
                button1.Text = "Disabled";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped);
                if (managementObject["StartMode"].ToString() == "Auto" || managementObject["StartMode"].ToString() == "Manual")
                {
                    object[] parameters = new object[1] { "Disabled" };
                    managementObject.InvokeMethod("ChangeStartMode", parameters);
                    managementObject = new ManagementObject(managementPath);
                }
                UpdateText("Good , Windows Update has been Stopped", managementObject["StartMode"].ToString(), Color.Green , Color.Green);
                button1.Enabled = false;
                button1.Text = "Disabled";

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //List<string> list = ServiceController.GetServices().Select(a => a.DisplayName).ToList
            // Create Shortcut
            if (!System.IO.File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),curAssembly.GetName().Name +".lnk")))
            {
                CreateShortcut(curAssembly.GetName().Name, Environment.GetFolderPath(Environment.SpecialFolder.Desktop), curAssembly.Location);
            }

            service.MachineName = Environment.MachineName;
            managementPath = new ManagementPath(path);
            managementObject = new ManagementObject(managementPath);
            // Start Up
            if (key.GetValue(curAssembly.GetName().Name) != null)
            {
                if (key.GetValue(curAssembly.GetName().Name).ToString() != curAssembly.Location)
                {
                    key.SetValue(curAssembly.GetName().Name, curAssembly.Location);
                }
                Settings.Default.StartUp = true;
                Settings.Default.Save();
                checkBox1.CheckState = CheckState.Checked;
            }
            else
            {
                Settings.Default.StartUp = false;
                Settings.Default.Save();
                checkBox1.CheckState = CheckState.Unchecked;
            }

            // Check Service
            if (service.Status == ServiceControllerStatus.Running)
            {
                UpdateText("Windows Update is Running", managementObject["StartMode"].ToString(), Color.Red , Color.Red);
                button1.Enabled = true;
            }
            else if (service.Status == ServiceControllerStatus.Stopped)
            {
                if (managementObject["StartMode"].ToString() != "Disabled")
                {
                    UpdateText("Good , Windows Update has been Stopped",managementObject["StartMode"].ToString(), Color.Green , Color.Red);
                    button1.Text = "Change StartMode to Disabled";
                    button1.Click -= button1_Click;
                    button1.Click += button_StartModeClick;
                }
                else
                {
                    UpdateText("Good , Windows Update has been Stopped", managementObject["StartMode"].ToString(), Color.Green, Color.Green);
                    button1.Enabled = false;
                    button1.Text = "Disabled";
                }
            }
            else
            {
                UpdateText("ReRun the App","", Color.Gray,Color.Gray);
                button1.Enabled = false;
            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (!Settings.Default.StartUp)
            {
                try
                {
                    key.SetValue(curAssembly.GetName().Name, curAssembly.Location);
                    Settings.Default.StartUp = true;
                    Settings.Default.Save();
                    MessageBox.Show("Program Added To StartUp Successfully :D","StartUp",MessageBoxButtons.OK,MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                }
            }
            else if (Settings.Default.StartUp && !checkBox1.Checked)
            {
                try
                {
                    key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                    key.DeleteValue(curAssembly.GetName().Name);
                    Settings.Default.StartUp = false;
                    Settings.Default.Save();
                    MessageBox.Show("Program Removed From StartUp Successfully :D", "StartUp", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                }
            }
        }
    }
}
