using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnlockBlackTheme
{
    class Program
    {
        static void Main(string[] args)
        {
            RegistryKey regKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\Common", true);
            if (regKey != null)
            {
                regKey.SetValue("PendingUITheme", "4", RegistryValueKind.DWord);
                regKey.SetValue("UI Theme", "4", RegistryValueKind.DWord);
                regKey.Close();
            }

            Process outlook = new Process();
            outlook.StartInfo.UseShellExecute = false;
            outlook.StartInfo.FileName = @"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.exe";
            outlook.Start();
        }
    }
}
