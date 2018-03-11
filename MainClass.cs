using System;
using System.Runtime.InteropServices;
using System.EnterpriseServices;
using System.IO;

namespace CSVWorker
{

    public class MainClass : IInitDone, ILanguageExtender
    {
        protected string AddInName = "ExtVK";
        protected string AddInVer = "2";
        protected string m_errorMessage = "";
        protected int m_errorCode = 0;
        protected Boolean bEnableDebugMode = false;

        private LogFile logFile;

        protected void SetError(int code, string msg)
        {
            m_errorCode = code;
            m_errorMessage = msg;
        }

        public string ErrorMessage
        {
            get
            {
                return m_errorMessage;
            }
        }

        #region // Инициализация компонента
        public void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            V8Data.V8Object = pConnection;

            logFile = new LogFile();
            logFile.OnlyErrors = !bEnableDebugMode;

        }
        // Возвращается информация о компоненте
        public void GetInfo(ref object[] pInfo)
        {

            pInfo.SetValue("2000", 0);

            logFile.Add(String.Format("Версия компоненты: {0}, версия API 1С: 2.0", this.AddInVer));
        }
        public void Done()
        {
            try
            {
                logFile.Add("Завершение компоненты");
            }
            catch (Exception e)
            {
                logFile.Add("Ошибка во время завершения компоненты (" + e.Message + ")", true);
            }
        }
        public void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref String extensionName)
        {
            logFile.Add("Регистрация AddIn: " + AddInName);
            try
            {
                extensionName = AddInName;
            }
            catch (Exception e)
            {
                logFile.Add("Ошибка при регистрации AddIn (" + e.Message + ")", true);
            }
        }
        #endregion

        //Props
        public struct name_s
        {
            public string name_ru;
            public string name_en;
            public name_s(string en, string ru)
            {
                name_en = en;
                name_ru = ru;
            }
        }

        private name_s[] props;

        private name_s[] meths;

        protected void InitNames(name_s[] p, name_s[] m)
        {
            props = new name_s[p.Length];
            meths = new name_s[m.Length];
            int i;
            for (i = 0; i < p.Length; i++)
            {
                props[i].name_en = p[i].name_en;
                props[i].name_ru = p[i].name_ru;
            }
            for (i = 0; i < m.Length; i++)
            {
                meths[i].name_en = m[i].name_en;
                meths[i].name_ru = m[i].name_ru;
            }
        }


        #region Properties

        public void GetNProps(ref int plProps)
        {
            plProps = props.Length;
        }
        public void FindProp(string propName, ref int propNum)
        {
            propName = propName.ToLower();
            for (int i = 0; i < props.Length; i++)
            {
                if (propName == props[i].name_en.ToLower() || propName == props[i].name_ru.ToLower())
                {
                    propNum = i;
                    return;
                }
            }

            propNum = -1;
        }
        public void GetPropName(int propNum, int propAlias, ref string propName)
        {
            //Здесь 1С (теоретически) узнает имя свойства по его идентификатору. lPropAlias - номер псевдонима
            if (propNum >= 0 && propNum < props.Length)
            {
                try
                {
                    if (propAlias == 1) // Russian
                        propName = props[propNum].name_ru;
                    else // English
                        propName = props[propNum].name_en;
                }
                catch (Exception e)
                {
                    logFile.Add("Ошибка при чтении имени свойства компоненты \"" + props[propNum].name_ru + "\" (" + e.Message + ")", true);
                }
            }
        }
        public void GetPropVal(int propNum, ref object propVal)
        {
            if (propNum >= 0 && propNum < props.Length)
            {
                try
                {
                    propVal = this.GetType().GetProperty(props[propNum].name_en).GetValue(this, null);
                }
                catch (Exception e)
                {
                    logFile.Add("Ошибка при получении значения свойства \"" + props[propNum].name_ru + "\" (" + e.Message + ")", true);
                }

            }
        }
        public void SetPropVal(int propNum, ref object propVal)
        {
            if (propNum >= 0 && propNum < props.Length)
            {
                try
                {
                    this.GetType().GetProperty(props[propNum].name_en).SetValue(this, propVal, null);
                }
                catch (Exception e)
                {
                    logFile.Add("Ошибка при установке значения свойства \"" + props[propNum].name_ru + "\" (" + e.Message + ")", true);
                }
            }
        }
        public void IsPropReadable(int propNum, ref bool propRead)
        {
            if (propNum >= 0 && propNum < props.Length)
            {
                propRead = this.GetType().GetProperty(props[propNum].name_en).CanRead;
            }
        }
        public void IsPropWritable(int propNum, ref bool propWrite)
        {
            //Здесь 1С узнает, какие свойства доступны для записи
            if (propNum >= 0 && propNum < props.Length)
            {
                propWrite = this.GetType().GetProperty(props[propNum].name_en).CanWrite;
            }
        }

        #endregion

        #region Methods
        public void GetNMethods(ref int plMethods)
        {
            plMethods = meths.Length;
        }

        public void FindMethod(string bstrMethodName, ref int plMethodNum)
        {
            //Здесь 1С получает числовой идентификатор метода (процедуры или функции) по имени (названию) процедуры или функции
            string tmp = bstrMethodName.ToLower();

            for (int i = 0; i < meths.Length; i++)
            {
                if (tmp == meths[i].name_en.ToLower() || tmp == meths[i].name_ru.ToLower())
                {
                    plMethodNum = i;
                    logFile.Add(String.Format("Поиск метода \"{0}\". Метод найден. Индекс метода: {1}", bstrMethodName, i));
                    return;
                }
            }
            plMethodNum = -1;
            logFile.Add(String.Format("Поиск метода \"{0}\". Метод не найден.", bstrMethodName));
        }

        public void GetMethodName(int lMethodNum, int lMethodAlias, ref string pbstrMethodName)
        {
            //Здесь 1С (теоретически) получает имя метода по его идентификатору. lMethodAlias - номер синонима.
            if (lMethodNum >= 0 && lMethodNum < meths.Length)
            {
                try
                {
                    if (lMethodAlias == 1) // Russian
                        pbstrMethodName = meths[lMethodNum].name_ru;
                    else // English
                        pbstrMethodName = meths[lMethodNum].name_en;
                }
                catch (Exception e)
                {
                    logFile.Add(String.Format("Ошибка при чтении имени метода компоненты \"{0}\" ({1})", meths[lMethodNum].name_ru, e.Message), true);
                }
            }
            pbstrMethodName = "";
        }

        public void GetNParams(int lMethodNum, ref int plParams)
        {
            //Здесь 1С получает количество параметров у метода (процедуры или функции)
            if (lMethodNum >= 0 && lMethodNum < meths.Length)
            {
                try
                {
                    System.Reflection.MethodInfo mi = this.GetType().GetMethod(meths[lMethodNum].name_en);
                    if (mi == null)
                    {
                        logFile.Add(String.Format("Метод {0} не найден", meths[lMethodNum].name_en));
                        plParams = 0;
                    }
                    else
                    {
                        plParams = mi.GetParameters().Length;
                        logFile.Add(String.Format("Возвращаем количество параметров для метода \"{0}\" ({1}): {2}", meths[lMethodNum].name_ru, lMethodNum, plParams));
                    }
                }
                catch (Exception e)
                {
                    logFile.Add(String.Format("Ошибка при получении списка параметров для метода \"{0}\" ({1})\", парам:{2} ", meths[lMethodNum].name_ru, e.Message,plParams), true);
                }
            }
        }

        public void GetParamDefValue(int lMethodNum, int lParamNum, ref object pvarParamDefValue)
        {
            //Здесь 1С получает значения параметров процедуры или функции по умолчанию
            pvarParamDefValue = null; //Нет значений по умолчанию
        }

        public void HasRetVal(int lMethodNum, ref bool pboolRetValue)
        {
            //Здесь 1С узнает, возвращает ли метод значение (т.е. является процедурой или функцией)
            pboolRetValue = true; //Все методы у нас будут функциями (т.е. будут возвращать значение).
        }

        public void CallAsProc(int lMethodNum, ref System.Array paParams)
        {
            //Здесь внешняя компонента выполняет код процедур.
        }

        public void CallAsFunc(int lMethodNum, ref object pvarRetValue, ref System.Array paParams)
        {
            logFile.Add(String.Format("Процедура выполняющая наши команды. Вызываем функцию номер: {0} ", lMethodNum));
            //Здесь внешняя компонента выполняет код функций.
            try
            {
                pvarRetValue = this.GetType().GetMethod(meths[lMethodNum].name_en).Invoke(this, (object[])paParams);
            }
            catch (Exception e)
            {
                pvarRetValue = 0; //Возвращаемое значение метода для 1С
                logFile.Add(String.Format("Ошибка при вызове метода {0} ({1} {2})", meths[lMethodNum].name_ru, e.Message, e.InnerException.Message), true);
            }
        }
        #endregion

    }

    [ComVisible(true)]
    [Guid("51E44EBF-B4D4-4D36-9378-C289AD88C950")]
    [ProgId("AddIn.ExtVK")]
    [Description("ExtVK.Component class")]

    public class Component : MainClass
    {
        private CsvReader reader;
        private Table1C vt;
        private LogFile logFile;
        private char cFieldDelim = ',';

        public int Version
        {
            get
            {
                return 2;
            }
        }

        public void StatusLine(string s1_2)
        {
            V8Data.StatusLine.SetStatusLine(s1_2);
            System.Threading.Thread.Sleep(1000); // 1 сек
            V8Data.StatusLine.ResetStatusLine();
        }

        public void Pause(Int32 ms)
        {
            System.Threading.Thread.CurrentThread.Join(ms);
        }

        #region // Work with CSV

        public void CreateCSVreader(string FilePath, string sFieldDelim)
        {
            char cFieldDelim = Convert.ToChar(sFieldDelim);
            try
            {
                reader = new CsvReader(FilePath, CsvOptions.HasFieldHeaders, cFieldDelim);
            }
            catch (Exception e)
            {
                logFile.Add(String.Format("Ошибка чтения файла {0} ({1}) ", FilePath, e.ToString()),true);
            }
        }

        public Boolean ReadCSV()
        {
            return reader.Read();
        }
     
        public void DisposeCSV()
        {
            reader.Dispose();
        }

        public void CreateHeader(string[] saTypes, string[] saNames)
        {
            vt = new Table1C(logFile);
            try
            {
                vt.CreateColumns(saTypes, saNames);
                logFile.Add(String.Format("Созданы колонки в таблице ({0} шт.)", saTypes.Length));
            }
            catch
            {
                logFile.Add(String.Format("Не удалось создать колонки в таблице ({0} шт.)", saTypes.Length),true);
            }
        }

        public String ReflectCSVtoTable(string[] saFiles, string sPathTmp, string sFieldDelim, bool skipHeader)
        {
            string sResult = "";
            string sFilePath = "";
            if (vt == null)
            {
                logFile.Add(String.Format("Необходимо инициализировать таблицу значений (CreateHeader(string[], string[], int[])) "));
                return "";
            }
            for (int i = 0; i < saFiles.Length; i++)
            {
                sFilePath = saFiles[i];
                CreateCSVreader(sFilePath,sFieldDelim);
                if (reader != null)
               {
                   while (reader.Read()) {
                        vt.AddRow(reader);
                    }  
                    DisposeCSV();
                }
        }

        try
        {
            File.WriteAllText(sPathTmp, vt.ToString());
            sResult = sPathTmp;
            logFile.Add(String.Format("Файл таблицы значений с внутренним представлением 1С записан ({0})", sPathTmp));
        }
        catch (Exception e)
        {
            logFile.Add(String.Format("Ошибка записи в файл {0} ({1}) ", sPathTmp, e.ToString()),true);
        }

        return sResult;
        }

        public void CreateCSVreader_(string FilePath)
        {
            try
            {
                reader = new CsvReader(FilePath, CsvOptions.None, cFieldDelim);
            }
            catch (Exception e)
            {
                logFile.Add(String.Format("Ошибка чтения файла {0} ({1}) ", FilePath, e.Message),true);
            }
        }

        #endregion

        public void EnableDebugMode(Boolean bDebugMode)
        {
            logFile.OnlyErrors = !bDebugMode;
        }

        public Component() // Обязательно для COM инициализации
        {
            logFile = new LogFile();
      
            AddInName = "ExtVK";

            name_s[] p = {
                            new name_s("Version",  "Версия")
            };

            name_s[] m = {
                            //Остались из шаблона
                            new name_s("Pause",             "Пауза"),
                            new name_s("StatusLine",        "Состояние"),
                            //Методы по чтению *.csv (могут использоваться для отладки)
                            new name_s("CreateCSVreader",   "СоздатьКомпонентCSV"),
                            new name_s("ReadCSV",           "ПрочитатьCSV"),
                            new name_s("DisposeCSV",        "ЗакрытьCSV"),
                            //Основные методы компоненты
                            new name_s("CreateHeader",      "СоздатьЗаголовок"),
                            new name_s("ReflectCSVtoTable", "КонвертироватьCSVвТЗ"),
                            new name_s("EnableDebugMode",   "ВключитьРежимОтладки"),
            };
            InitNames(p, m);
        }
    }
}
