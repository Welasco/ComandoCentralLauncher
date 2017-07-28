using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using System.Windows.Controls;

namespace ComandoCentralLauncher
{
    public class Log
    {
        private static string folderCCtempfolder = ConfigurationManager.AppSettings["folder:ComandoCentraltempfolder"];
        private static string fileLogConfigName = ConfigurationManager.AppSettings["file:Log"];
        private static string dstCCFolder = System.IO.Path.Combine(Environment.GetEnvironmentVariable("temp"), folderCCtempfolder);
        private static string dstlogfile = System.IO.Path.Combine(dstCCFolder, fileLogConfigName);

        private static bool writeHead = false;
        private static bool firstexec = true;
        public static async Task Add(MainWindow window,string msg)
        {
            //AddLog(msg);
            //Random r = new Random();
            //int start2 = r.Next(0, 8);
            //AddScreen(window, msg + " " + start2);
            //await Task.Delay(1000 * start2);
            AddScreen(window, msg);
            AddLog(msg);
            await Task.Delay(100);
            //return "";
        }

        private static void AddLog(string msg)
        {
            if (firstexec)
            {
                if (File.Exists(dstlogfile))
                {
                    File.Delete(dstlogfile);
                }
                firstexec = false;
            }
            if (!Directory.Exists(Path.GetDirectoryName(dstlogfile)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(dstlogfile));
            }
            StreamWriter file = new StreamWriter(dstlogfile, true);

            try
            {
                if (!writeHead)
                {
                    string fileheader = "###############################################################\r";
                    fileheader = fileheader + "############### Comando Central Launcher ######################\r";
                    fileheader = fileheader + "############### " + DateTime.Now + " #################\r";
                    fileheader = fileheader + "###############################################################\r";
                    fileheader = fileheader + msg;
                    file.WriteLine(fileheader);
                    writeHead = true;
                }
                else
                {
                    file.WriteLine(msg);
                }
            }
            finally
            {
                file.Close();
            }
        }

        private static void AddScreen(MainWindow window,string msg)
        {
            window.textBox.AppendText(msg + "\r");
            //return "";
        }
        

    }
}
