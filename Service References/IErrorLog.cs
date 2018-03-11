using System;
using System.Runtime.InteropServices;

namespace CSVWorker
{
    [Guid("3127CA40-446E-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IErrorLog
    {
        int AddError([MarshalAs(UnmanagedType.BStr)]string pszPropName, ref ExcepInfo pExepInfo);
    }
    //----------------------------------------------------------
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Pack = 8)]
    public struct ExcepInfo
    {
        public short wCode;
        public short wReserved;
        [MarshalAs(UnmanagedType.BStr)]
        public string bstrSource;
        [MarshalAs(UnmanagedType.BStr)]
        public string bstrDescription;
        [MarshalAs(UnmanagedType.BStr)]
        public string bstrHelpFile;
        public int dwHelpContext;
        public IntPtr pvReserved;
        public IntPtr pfnDereffered;
        public int scode;
    }
}
