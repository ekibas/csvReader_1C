using System;
using System.IO;

namespace CSVWorker
{
    // TODO 1. Возможность считать время выполнения методов компоненты
    // TODO 2. Копить все ошибки в объекте
    
    class LogFile
    {
        const string FileName = "ExtComp.log";
        public string FilePath;
        public bool OnlyErrors = true;

        public LogFile()
        {
            string startupPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\";
            FilePath = String.Format("{0}{1}{2}", startupPath, "\\", FileName);
        }

        public void Add(string msg, bool error = false)
        {
            if (!error && OnlyErrors)
                return;

            if (String.IsNullOrWhiteSpace(FilePath))
                return;
    
            try
            {
                // Процедура добавляет в конец файла новую строку сообщения
                // Попытаемся сделать запись, через попытку/исключение, очень часто политика безопастности запрещает пользователю записывать файл 
                StreamWriter OutputFile;
                OutputFile = File.AppendText(FilePath);
                OutputFile.WriteLine(DateTime.Now.ToString() + " -- " + msg);
                OutputFile.Close();
            }
            catch (Exception)
            {
                //не удалось записать файл
            }

        }
    }
}
