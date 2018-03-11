using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace CSVWorker
{
    [Guid("ab634005-f13d-11d0-a459-004095e1daea"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStatusLine
    {
        /// <summary>
        /// Задает текст статусной строки
        /// </summary>
        /// <param name="bstrStatusLine">Текст статусной строки</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT SetStatusLine(BSTR bstrStatusLine);
        /// </prototype>
        /// </remarks>
        void SetStatusLine(
          [MarshalAs(UnmanagedType.BStr)]String bstrStatusLine
          );

        /// <summary>
        /// Сброс статусной строки
        /// </summary>
        /// <remarks>
        /// <propotype>
        /// HRESULT ResetStatusLine();
        /// </propotype>
        /// </remarks>
        void ResetStatusLine();
    }
}
