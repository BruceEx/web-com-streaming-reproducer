using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelRtdTest
{
    [ComVisible(true)]
    public interface ITestRtd
    {
        object TESTRTD(object arg1, object arg2);
    }

    [ComDefaultInterface(typeof(ITestRtd))]
    [Guid(Constants.AddinGuid)]
    [ComVisible(true)]
    [ProgId(Constants.AddinProgId)]
    public class Connect : IDTExtensibility2, ITestRtd
    {
        private Application _excel;
        private IRtdServer _server;

        public object TESTRTD(object topic, object arg /* = null */)
        {
            var wsfunc = _excel.WorksheetFunction;
            var result = wsfunc.RTD(Constants.ServerProgId, null, topic, arg);
            return result;
        }

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type, "Programmable"));
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(GetSubKeyName(type, "InprocServer32"), true);
            key.SetValue("", System.Environment.SystemDirectory + @"\mscoree.dll", RegistryValueKind.String);
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type, "Programmable"), false);
        }

        private static string GetSubKeyName(Type type, string subKeyName)
        {
            System.Text.StringBuilder s = new System.Text.StringBuilder();
            s.Append(@"CLSID\{");
            s.Append(type.GUID.ToString().ToUpper());
            s.Append(@"}\");
            s.Append(subKeyName);
            return s.ToString();
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Logger.Log("OnConnection method called");
            _excel = application as Application;
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Logger.Log("OnDisconnection method called");
            //_server.ServerTerminate();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            Logger.Log("OnAddInsUpdate method called");
        }

        public void OnStartupComplete(ref Array custom)
        {
            Logger.Log("OnStartupComplete method called");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Logger.Log("OnBeginShutdown method called");
        }
    }
}
