using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace ComandoCentralLauncher
{
    class LaunchAccess
    {

        private string GetMSAccessPathType
        {
            get
            {
                Type MSAccessType = Type.GetTypeFromProgID("Access.Application");
                dynamic MSAccessInst = Activator.CreateInstance(MSAccessType);
                var MSAccessPath = MSAccessInst.SysCmd(9);
                return MSAccessPath;
            }
        }

        public static void Launch()
        {
            string fileCCUsuario = ConfigurationManager.AppSettings["file:ComandoCentralUsuario"];
            string fileSeguranca = ConfigurationManager.AppSettings["file:Seguranca"];
            string folderCCtempfolder = ConfigurationManager.AppSettings["folder:ComandoCentraltempfolder"];
            string dstCCfolder = System.IO.Path.Combine(Environment.GetEnvironmentVariable("temp"), folderCCtempfolder);
            string dstfileCCseguranca = System.IO.Path.Combine(dstCCfolder, fileSeguranca);
            string dstfileCCusuario = System.IO.Path.Combine(dstCCfolder, fileCCUsuario);
            string cmd = System.IO.Path.Combine(GetMSAccessPath().ToString(), "MSAccess.exe");
            string cmdparam = dstfileCCusuario + " /wrkgrp " + dstfileCCseguranca;

            ProcessStartInfo MSAccessAPP = new ProcessStartInfo(cmd, cmdparam)
            {
            };

            Process.Start(MSAccessAPP);
        }

        private static string GetMSAccessPath()
        {
                Type MSAccessType = Type.GetTypeFromProgID("Access.Application");
                dynamic MSAccessInst = Activator.CreateInstance(MSAccessType);
                var MSAccessPath = MSAccessInst.SysCmd(9);
                return MSAccessPath;
        }
    }
}

