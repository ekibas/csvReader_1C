﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.EnterpriseServices;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CSVWorker
{
    [Guid("AB634001-F13D-11d0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]    
    public interface IInitDone
    {
        /// <summary>
        /// Инициализация компонента
        /// </summary>
        /// <param name="connection">reference to IDispatch</param>
        void Init(
          [MarshalAs(UnmanagedType.IDispatch)]
    object connection);

        /// <summary>
        /// Вызывается перед уничтожением компонента
        /// </summary>
        void Done();

        /// <summary>
        /// Возвращается инициализационная информация
        /// </summary>
        /// <param name="info">Component information</param>
        void GetInfo(
          [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
    ref object[] info);
    }
}
