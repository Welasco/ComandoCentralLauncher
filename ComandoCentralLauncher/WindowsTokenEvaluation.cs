using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ComandoCentralLauncher
{
    class WindowsTokenEvaluation
    {
        public enum TOKEN_ELEVATION_TYPE
        {
            TokenElevationTypeDefault = 1,
            TokenElevationTypeFull,
            TokenElevationTypeLimited
            //TokenElevationTypeFull = 2,
            //TokenElevationTypeLimited = 3
        }

        public enum TOKEN_INFORMATION_CLASS
        {
            TokenUser = 1,
            TokenGroups,
            TokenPrivileges,
            TokenOwner,
            TokenPrimaryGroup,
            TokenDefaultDacl,
            TokenSource,
            TokenType,
            TokenImpersonationLevel,
            TokenStatistics,
            TokenRestrictedSids,
            TokenSessionId,
            TokenGroupsAndPrivileges,
            TokenSessionReference,
            TokenSandBoxInert,
            TokenAuditPolicy,
            TokenOrigin,
            TokenElevationType,
            TokenLinkedToken,
            TokenElevation,
            TokenHasRestrictions,
            TokenAccessInformation,
            TokenVirtualizationAllowed,
            TokenVirtualizationEnabled,
            TokenIntegrityLevel,
            TokenUIAccess,
            TokenMandatoryPolicy,
            TokenLogonSid,
            MaxTokenInfoClass

            //TokenGroups = 2,
            //TokenPrivileges = 3,
            //TokenOwner = 4,
            //TokenPrimaryGroup = 5,
            //TokenDefaultDacl = 6,
            //TokenSource = 7,
            //TokenType = 8,
            //TokenImpersonationLevel = 9,
            //TokenStatistics = 10,
            //TokenRestrictedSids = 11,
            //TokenSessionId = 12,
            //TokenGroupsAndPrivileges = 13,
            //TokenSessionReference = 14,
            //TokenSandBoxInert = 15,
            //TokenAuditPolicy = 16,
            //TokenOrigin = 17,
            //TokenElevationType = 18,
            //TokenLinkedToken = 19,
            //TokenElevation = 20,
            //TokenHasRestrictions = 21,
            //TokenAccessInformation = 22,
            //TokenVirtualizationAllowed = 23,
            //TokenVirtualizationEnabled = 24,
            //TokenIntegrityLevel = 25,
            //TokenUIAccess = 26,
            //TokenMandatoryPolicy = 27,
            //TokenLogonSid = 28,
            //MaxTokenInfoClass = 29
        }

        public static void CheckElevation()
        {
            TOKEN_ELEVATION_TYPE tokenElevationType = GetTokenElevationType();
            //if (tokenElevationType != TOKEN_ELEVATION_TYPE.TokenElevationTypeFull)
            if (tokenElevationType == TOKEN_ELEVATION_TYPE.TokenElevationTypeLimited)
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo(Assembly.GetExecutingAssembly().CodeBase);
                // I guess I have to change it to false.
                processStartInfo.UseShellExecute = true;
                processStartInfo.Verb = "runas";
                Process.Start(processStartInfo);
                Application.Current.Shutdown();
                return;
            }

        }
        // ExTracePlus.ExTracePlusUtility
        public static TOKEN_ELEVATION_TYPE GetTokenElevationType()
        {
            TOKEN_ELEVATION_TYPE result = TOKEN_ELEVATION_TYPE.TokenElevationTypeDefault;
            if (Environment.OSVersion.Platform != PlatformID.Win32NT || Environment.OSVersion.Version.Major < 6)
            {
                return result;
            }
            TOKEN_ELEVATION_TYPE tOKEN_ELEVATION_TYPE = TOKEN_ELEVATION_TYPE.TokenElevationTypeDefault;
            uint num = 0u;
            uint num2 = (uint)Marshal.SizeOf((int)tOKEN_ELEVATION_TYPE);
            IntPtr intPtr = Marshal.AllocHGlobal((int)num2);
            try
            {
                if (GetTokenInformation(WindowsIdentity.GetCurrent().Token, TOKEN_INFORMATION_CLASS.TokenElevationType, intPtr, num2, out num))
                {
                    result = (TOKEN_ELEVATION_TYPE)Marshal.ReadInt32(intPtr);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(intPtr);
            }
            return result;
        }


        // ExTracePlus.ExTracePlusUtility
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool GetTokenInformation(IntPtr TokenHandle, TOKEN_INFORMATION_CLASS TokenInformationClass, IntPtr TokenInformation, uint TokenInformationLength, out uint ReturnLength);
    }
}
