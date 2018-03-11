using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSVWorker
{
    internal class V8Data
    {
        public static object V8Object
		{
			get
			{
				return m_V8Object;
			}
			set
			{
				m_V8Object = value;
				// Вызываем неявно QueryInterface
				m_ErrorInfo = (IErrorLog) value;
				m_AsyncEvent = (IAsyncEvent) value;
				m_StatusLine = (IStatusLine) value;
			}
		}
		public static IErrorLog ErrorLog
		{
			get
			{
				return m_ErrorInfo;
			}
		}
		public static IAsyncEvent AsyncEvent
		{
			get
			{
				return m_AsyncEvent;
			}
		}
		public static IStatusLine StatusLine
		{
			get
			{
				return m_StatusLine;
			}
		}
		private static object m_V8Object;
		private static IErrorLog m_ErrorInfo;
		private static IAsyncEvent m_AsyncEvent;
		private static IStatusLine m_StatusLine;
	}
}
