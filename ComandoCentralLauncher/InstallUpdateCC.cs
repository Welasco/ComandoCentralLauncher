using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using Microsoft.Win32;


namespace ComandoCentralLauncher
{
    class InstallUpdateCC
    {
        // Load config entries
        private static string fileAdView = ConfigurationManager.AppSettings["file:AdView"];
        private static string fileVText = ConfigurationManager.AppSettings["file:VText"];
        private static string fileMSCAL = ConfigurationManager.AppSettings["file:MSCAL"];
        private static string fileCCUsuario = ConfigurationManager.AppSettings["file:ComandoCentralUsuario"];
        private static string fileSeguranca = ConfigurationManager.AppSettings["file:Seguranca"];
        private static string fileVersion = ConfigurationManager.AppSettings["file:Version"];
        private static string folderCCtempfolder = ConfigurationManager.AppSettings["folder:ComandoCentraltempfolder"];
        private static string folderPathCCsharedfolder = ConfigurationManager.AppSettings["folder:ComandoCentralsourcesharedfolder"];
        
        //// System folders
        private static string sysWOW64 = System.IO.Path.Combine(Environment.GetEnvironmentVariable("windir"), "SysWOW64");
        private static string system32 = Environment.SystemDirectory;

        //// Defining source File Paths
        private static string srcfileAdView = System.IO.Path.Combine(folderPathCCsharedfolder, fileAdView);
        private static string srcfileVText = System.IO.Path.Combine(folderPathCCsharedfolder, fileVText);
        private static string srcfileMSCAL = System.IO.Path.Combine(folderPathCCsharedfolder, fileMSCAL);
        private static string srcfileCCUsuario = System.IO.Path.Combine(folderPathCCsharedfolder, fileCCUsuario);
        private static string srcfileSeguranca = System.IO.Path.Combine(folderPathCCsharedfolder, fileSeguranca);
        private static string srcfileVersion = System.IO.Path.Combine(folderPathCCsharedfolder, fileVersion);

        // Defining destination File Paths
        // x64
        private static string dstfile64AdView = System.IO.Path.Combine(sysWOW64, fileAdView);
        private static string dstfile64VText = System.IO.Path.Combine(sysWOW64, fileVText);
        private static string dstfile64MSCAL = System.IO.Path.Combine(sysWOW64, fileMSCAL);
        // x32         
        private static string dstfile32AdView = System.IO.Path.Combine(system32, fileAdView);
        private static string dstfile32VText = System.IO.Path.Combine(system32, fileVText);
        private static string dstfile32MSCAL = System.IO.Path.Combine(system32, fileMSCAL);
        // CC User Files
        private static string dstCCfolder = System.IO.Path.Combine(Environment.GetEnvironmentVariable("tmp"), folderCCtempfolder);
        private static string dstfileCCusuario = System.IO.Path.Combine(dstCCfolder, fileCCUsuario);
        private static string dstfileCCseguranca = System.IO.Path.Combine(dstCCfolder, fileSeguranca);
        private static string dstfileCCVersion = System.IO.Path.Combine(dstCCfolder, fileVersion);

        // Registry Key
        static Microsoft.Win32.RegistryKey Key;
        private MainWindow passWindow
        {
            get;
            set;
        }
        public InstallUpdateCC(MainWindow window)
        {
            passWindow = window;
        }

        // OS Version
        private static bool is64Bit = Environment.Is64BitOperatingSystem;
        public async Task installCCBin()
        {

            if (is64Bit)
            {
                await Log.Add(passWindow, "Instalando o Comando Central 64 bits...");
                await Log.Add(passWindow, "Ativando Macro no Microsoft Access...");
                if(EnableMacro())
                {
                    await Log.Add(passWindow, "Macro ativada com sucesso no Microsoft Access!");
                }
                else
                {
                    await Log.Add(passWindow, "Falha ao ativar a Macro no Microsoft Access!");
                }
                try
                {
                    await Log.Add(passWindow, "Copiando arquivos de sistema...");
                    await Task.Run(() => { 
                        File.Copy(srcfileAdView, dstfile64AdView, true);
                        File.Copy(srcfileAdView, dstfile32AdView, true);
                        File.Copy(srcfileVText, dstfile64VText, true);
                        File.Copy(srcfileVText, dstfile32VText, true);
                        File.Copy(srcfileMSCAL, dstfile64MSCAL, true);
                        File.Copy(srcfileMSCAL, dstfile32MSCAL, true);
                    });
                    await Log.Add(passWindow, "Arquivos de sistema copiados com sucesso!");
                }
                catch (Exception e)
                {
                    await Log.Add(passWindow, "Falha ao copiar os arquivos de sistema:");
                    await Log.Add(passWindow, e.Message.Trim());
                    await Log.Add(passWindow, e.StackTrace.Trim());
                    //throw e;
                }
            }
            else
            {
                await Log.Add(passWindow, "Instalando o Comando Central 32 bits...");
                try
                {
                    await Log.Add(passWindow, "Copiando arquivos de sistema...");
                    await Task.Run(() =>
                    {
                        File.Copy(srcfileVText, dstfile32VText, true);
                        File.Copy(srcfileMSCAL, dstfile64MSCAL, true);
                        File.Copy(srcfileMSCAL, dstfile32MSCAL, true);
                    });
                }
                catch (Exception e)
                {
                    await Log.Add(passWindow, "Falha ao copiar os arquivos de sistema:");
                    await Log.Add(passWindow, e.Message.Trim());
                    await Log.Add(passWindow, e.StackTrace.Trim());
                    //throw e;
                }
            }
        }

        public async Task<bool> isCCBinInstalled()
        {
            bool isfile64AdView = File.Exists(dstfile64AdView);
            bool isfile64VText = File.Exists(dstfile64VText);
            bool isfile64MSCAL = File.Exists(dstfile64MSCAL);

            bool isfile32AdView = File.Exists(dstfile32AdView);
            bool isfile32VText = File.Exists(dstfile32VText);
            bool isfile32MSCAL = File.Exists(dstfile32MSCAL);

            bool isvarMacroEnabled = isMacroEnabled();

            await Log.Add(passWindow,"Checando a instalação do Comando Central...");
            if (is64Bit)
            {
                await Log.Add(passWindow, "Detectado a versão 64bits do Sistema Operacional");
                if (!isfile64AdView || !isfile64VText || !isfile64MSCAL || !isfile32AdView || !isfile32VText || !isfile32MSCAL || !isvarMacroEnabled)
                {
                    await Log.Add(passWindow, "Comando Central não encontrado!");
                    return false;
                }
            }
            else
            {
                await Log.Add(passWindow, "Detectado a versão 32bits do Sistema Operacional");
                if (!isfile32AdView || !isfile32VText || !isfile32MSCAL || !isvarMacroEnabled)
                {
                    await Log.Add(passWindow, "Comando Central não encontrado!");
                    return false;
                }
            }
            await Log.Add(passWindow, "Comando Central está instalado corretamente.");
            return true;
        }
        public async Task UpdateCC()
        {
            await Log.Add(passWindow, "Checando se o Comando Central esta atualizado...");
            if (!File.Exists(dstfileCCusuario))
            {
                await Log.Add(passWindow, "Comando Central Usuario não encontrado");
                try
                {
                    await Log.Add(passWindow, "Copiando Comando Central Usuario...");
                    await Task.Run(() =>
                    {
                        if (!Directory.Exists(Path.GetDirectoryName(dstfileCCusuario)))
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(dstfileCCusuario));
                        }
                        File.Copy(srcfileCCUsuario, dstfileCCusuario, true);
                        File.Copy(srcfileSeguranca, dstfileCCseguranca, true);
                        File.Copy(srcfileVersion, dstfileCCVersion, true);
                    });
                    await Log.Add(passWindow, "Comando Central Usuario Copiado com sucesso!");
                }
                catch (Exception e)
                {
                    await Log.Add(passWindow, "Falha ao copiar Comando Central Usuario:");
                    await Log.Add(passWindow, e.Message.Trim());
                    await Log.Add(passWindow, e.StackTrace.Trim());
                }

            }
            else
            {
                await Log.Add(passWindow, "Comando Central encontrado");
                await Log.Add(passWindow, "Checando se existe uma no versão...");
                try
                {
                    using (StreamReader srvSR = new StreamReader(File.Open(srcfileVersion, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                    {
                        string strsrvVersion = srvSR.ReadLine();
                        int srvVersion = int.Parse(strsrvVersion);

                        using (StreamReader localSR = new StreamReader(File.Open(dstfileCCVersion, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                        {
                            string strlocalVersion = localSR.ReadLine();
                            int localVersion = int.Parse(strlocalVersion);
                            if (srvVersion > localVersion)
                            {
                                await Log.Add(passWindow, "Uma nova versão do Comando Cetral Usuario foi encontrada!");
                                await Log.Add(passWindow, "Versão atual: " + localVersion);
                                await Log.Add(passWindow, "Versão atualizada: " + srvVersion);
                                await Log.Add(passWindow, "Atualizando Comando Central Usuario...");
                                try
                                {
                                    await Task.Run(() =>
                                    {
                                        File.Copy(srcfileCCUsuario, dstfileCCusuario, true);
                                        File.Copy(srcfileSeguranca, dstfileCCseguranca, true);
                                        File.Copy(srcfileVersion, dstfileCCVersion, true);
                                    });
                                    await Log.Add(passWindow, "Comando Central Usuario atualizado com Sucesso!");
                                }
                                catch (Exception e)
                                {
                                    await Log.Add(passWindow, "Falha ao copiar Comando Central Usuario.");
                                    await Log.Add(passWindow, e.Message.Trim());
                                    await Log.Add(passWindow, e.StackTrace.Trim());
                                    throw;
                                }
                            }
                            else
                            {
                                await Log.Add(passWindow, "Comando Central já esta na versão atual!");
                            }
                        }

                    }
                }
                catch (Exception e)
                {
                    await Log.Add(passWindow, "Falha ao checar atualização Comando Central Usuario.");
                    await Log.Add(passWindow, e.Message.Trim());
                    await Log.Add(passWindow, e.StackTrace.Trim());
                    //throw;
                }
            }
        }

        public async Task<bool> isMSAccessInstalled()
        {
            await Log.Add(passWindow, "Checando se o Microsoft Access está instalado...");
            try
            {
                Type MSAccessType = Type.GetTypeFromProgID("Access.Application");
                dynamic MSAccessInst = Activator.CreateInstance(MSAccessType);
                var MSAccessVersion = MSAccessInst.Version;
                await Log.Add(passWindow, "Microsoft Access encontrado: " + MSAccessVersion);
                return true;
            }
            catch (Exception e)
            {
                await Log.Add(passWindow, "Nenhuma versão do Microsoft Access foi encontrada!");
                await Log.Add(passWindow, "O Comando Central não será iniciado!");
                await Log.Add(passWindow, e.Message.Trim());
                await Log.Add(passWindow, e.StackTrace.Trim());
                return false;
                //throw;
            }
        }

        private static bool isMacroEnabled()
        {
            try
            {
                Type MSAccessType = Type.GetTypeFromProgID("Access.Application");
                dynamic MSAccessInst = Activator.CreateInstance(MSAccessType);
                string MSAccessVersion = MSAccessInst.Version;

                Key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Microsoft\\Office\\" + MSAccessVersion + "\\Access\\Security");
                string VBAWarningsKey = Key.GetValue("VBAWarnings").ToString();
                int intVBAWarningsKey = int.Parse(VBAWarningsKey);
                if (intVBAWarningsKey == 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                return false;
            }
        }
        private static bool EnableMacro()
        {
            try { 
                Type MSAccessType = Type.GetTypeFromProgID("Access.Application");
                dynamic MSAccessInst = Activator.CreateInstance(MSAccessType);
                string MSAccessVersion = MSAccessInst.Version;
            
                Key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Microsoft\\Office\\" + MSAccessVersion + "\\Access\\Security");
                Key.SetValue("VBAWarnings", 1, RegistryValueKind.DWord);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
    }
}
