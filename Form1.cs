using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Threading;
using System.Management;

namespace AutoUpdater
{
    public partial class Form1 : Form
    {
        string downoadfileName = @"\\10.0.85.129\Shared-MposFile\SedClientTools";
        string sourcefileName = Application.StartupPath;
        int myLastversion = 0;
        List<Version> versionList = new List<Version>();
        Updateconfig updateconfig = new Updateconfig();

        public Form1()
        {
            InitializeComponent();
            using (StreamReader r = new StreamReader(sourcefileName + @"\LastVersion.json"))
            {
                string json = r.ReadToEnd();
                myLastversion = JsonConvert.DeserializeObject<List<Version>>(json).First().VersionNumber;
            }
            using (StreamReader r = new StreamReader(sourcefileName + @"\UpdateConfig.json"))
            {
                string json = r.ReadToEnd();
                updateconfig = JsonConvert.DeserializeObject<List<Updateconfig>>(json).First();
                downoadfileName = updateconfig.UpdateUrl;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
        }
        private bool Is_App_Exist()
        {
            Process[] runingProcess = Process.GetProcesses();
            for (int i = 0; i < runingProcess.Length; i++)
            {
                // compare equivalent process by their name
                if (runingProcess[i].ProcessName == "SedClientTools")
                {
                    return true;
                }
            }
            return false;
        }
        private void Run_App()
        {
            if (Is_App_Exist())
            {
                MessageBox.Show("یک نسخه از نرم افزار در حال اجرا است");
                return;
            }
            // MessageBox.Show("5");
            //Application.StartupPath + @"\Update.exe"
            Process.Start(sourcefileName + @"\SedClientTools.exe");
            Application.Exit();
        }
        private void Online_Update_File()
        {            
            if (Is_App_Exist())
            {
                var dialogResult = MessageBox.Show("یک نسخه از نرم افزار در حال اجرا است،آیا مایلید آن را ببندید", "بروز رسانی نرم افزار", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.No)
                    Application.Exit();
                else
                {
                    var proccess = Process.GetProcessesByName("SedClientTools");
                    proccess.First().Kill();
                    Thread.Sleep(1000);
                    Online_Update_File();
                }
            }
            //Download zip file
            using (StreamReader r = new StreamReader(@"\\" + downoadfileName + @"\UpdateVersion.json"))
            {
                string json = r.ReadToEnd();
                versionList = JsonConvert.DeserializeObject<List<Version>>(json);
                versionList = versionList.OrderBy(t => t.VersionNumber).ToList();
            }
            // MessageBox.Show("1");
            for (int i = 0; i < versionList.Count; i++)
            {
                var version = versionList.Where(t => t.VersionNumber == myLastversion + 1).FirstOrDefault();
                if (version != null)
                {
                    var fileIsExist = File.Exists(sourcefileName + @"\Update\SedClientTools" + version.VersionNumber + ".zip");
                    if (!fileIsExist)
                        using (var client = new WebClient())
                        {
                            client.DownloadFile(@"\\" + version.Url, sourcefileName + @"\Update\SedClientTools" + version.VersionNumber + ".zip");
                        }

                    Extract_File(version);
                    myLastversion += 1;
                }
            }
            // MessageBox.Show("2");
            Run_App();
        }
        private void Offline_Update_File()
        {
            if (Is_App_Exist())
            {
                var proccess = Process.GetProcessesByName("SedClientTools");
                proccess.First().Kill();
                Thread.Sleep(1000);
                Offline_Update_File();
            }
            
            var version = new Version { VersionNumber = myLastversion + 1 };
            var fileIsExist = File.Exists(sourcefileName + @"\Update\SedClientTools" + version.VersionNumber + ".zip");
            if (fileIsExist)
                Extract_File(version);
            Run_App();
        }
        private void Extract_File(Version version)
        {
            using (ZipArchive archive = ZipFile.OpenRead(sourcefileName + @"\Update\SedClientTools" + version.VersionNumber + ".zip"))
            {
                var files = archive.Entries;
                foreach (var file in files)
                {
                    File.Delete(sourcefileName + @"\" + file.Name);
                }

                archive.ExtractToDirectory(sourcefileName);
            }
            var LastVersion = new List<Version>();
            LastVersion.Add(new Version
            {
                VersionNumber = version.VersionNumber,
            });

            string output = Newtonsoft.Json.JsonConvert.SerializeObject(LastVersion, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(sourcefileName + @"\LastVersion.json", output);

            File.Delete(sourcefileName + @"\Update\SedClientTools" + version.VersionNumber + ".zip");
        }
        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            var proccess1 = Process.GetProcessesByName("AutoUpdater");
            var result = GetCommandLine(proccess1.First());
            var state = result[result.Length - 1].ToString();
            if (state == "1")
                Offline_Update_File();
            else
                Online_Update_File();
        }
        private string GetCommandLine(Process process)
        {
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT CommandLine FROM Win32_Process WHERE ProcessId=" + process.Id))
            using (ManagementObjectCollection objects = searcher.Get())
            {
                return objects.Cast<ManagementBaseObject>().SingleOrDefault()?["CommandLine"]?.ToString();
            }

        }
    }

    public class Version
    {
        public int VersionNumber;
        public string Url;
    }
    public class Updateconfig
    {
        public string UpdateUrl;
        public int ForceUpdate;
    }

}
