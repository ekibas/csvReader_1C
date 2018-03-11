using System;
using System.Collections;
using System.Linq;
using System.Text;

namespace CSVWorker
{
   
    class Table1C
    {
        private StringBuilder sbTable;
        private int nRCountPlace;   // индекс символа StringBuilder'a после создания шапки
        private int nCCount;
        private int nRCount;
        private string[] saColumns;
        private Hashtable hTypes;
        private LogFile logFile;

        public Table1C(LogFile log)
        {
            sbTable = new StringBuilder();
            sbTable.Clear();
            hTypes  = new Hashtable();
            hTypes.Add("S", "{\"Pattern\", {\"S\"}}");
            hTypes.Add("D", "{\"Pattern\", {\"D\"}}");
            hTypes.Add("N", "{\"Pattern\", {\"N\"}}");
            hTypes.Add("B", "{\"Pattern\", {\"B\"}}");

            InputHeader();
            logFile = log;
            logFile.Add("Компонента запущена из - " + System.Reflection.Assembly.GetExecutingAssembly().Location.ToString());

        }
           
        public void CreateColumns(string[] saTypes, string[] saNames){

            string sType = "";
            nCCount = saTypes.Length;
            saColumns = saTypes;

            sbTable.Append('{');
            sbTable.Append(nCCount);
            for (int i = 0; i < nCCount; i++)
            {
                sbTable.Append(',');
                sbTable.Append('{');
                sbTable.Append(i);
                sbTable.Append(',');

                // Имя 
                sbTable.Append('"');
                // todo: проверять на пробелы
                sbTable.Append(saNames[i]);
                sbTable.Append('"');
                sbTable.Append(',');

                // Тип
                sType = hTypes[saTypes[i]].ToString();
                sbTable.Append(sType);
                sbTable.Append(',');

                // Заголовок
                sbTable.Append('"');
                sbTable.Append(saNames[i]);
                sbTable.Append('"');
                sbTable.Append(',');

                // Ширина поля
                sbTable.Append(0); 

                sbTable.Append('}');
            }

            sbTable.Append('}');
            sbTable.Append(new char[] {',', '{', '2', ','});
            sbTable.Append(nCCount);
            sbTable.Append(',');

            for (int i = 0; i < nCCount; i++)
            {
                sbTable.Append(i);
                sbTable.Append(',');
                sbTable.Append(i);
                sbTable.Append(',');
            }

            sbTable.Append(new char[] { '{', '1', ',' });
            nRCountPlace = sbTable.Length;
        }
        
        public void AddRow(CsvReader reader)
        {
            Object[] oaFields = reader.GetValues();
  
            sbTable.Append(",{2,");
            sbTable.Append(nRCount);
            sbTable.Append(',');

            int nFieldsCount = oaFields.Length;
            if (nFieldsCount >= nCCount) {
                sbTable.Append(nCCount);
                for (int i = 0; i < nCCount; i++) {
                    sbTable.Append(',');
                    sbTable.Append('{');
                    
                    // Тип
                    sbTable.Append('"');
                    sbTable.Append(saColumns[i]);
                    sbTable.Append('"');
                    sbTable.Append(',');
                  
                    // Значение
                    string sValue = "";
                    try
                    {
                        switch (saColumns[i])
                        {
                            case "S":
                                {
                                    sbTable.Append('"');
                                    sValue = reader.GetValue(i).ToString();
                                    sValue = sValue.Replace("\"", "\"\"");
                                    sbTable.Append(sValue);
                                    sbTable.Append('"');
                                    break;
                                }
                            case "N":
                                {
                                    sValue = reader.GetString(i);
                                    if (String.IsNullOrEmpty(sValue))
                                        sValue = "0";
                                    else
                                        sValue = sValue.Replace(',', '.');
                                    sbTable.Append(sValue);
                                    break;
                                }
                            case "D":
                                {
                                    try
                                    {
                                        DateTime dtValue = reader.GetDateTime(i);
                                        sValue = dtValue.ToString("yyyyMMddhhmmss");
                                        sbTable.Append(sValue);
                                    }
                                    catch
                                    {
                                        sbTable.Append('"');
                                        sValue = reader.GetValue(i).ToString();
                                        sValue = sValue.Replace("\"", "\"\"");
                                        sbTable.Append(sValue);
                                        sbTable.Append('"');
                                        logFile.Add(String.Format("Не удалось преобразовать строку ({0}) в дату. Строка {1}, колонка {2}", sValue, nRCount + 1, i), true);
                                    }
                                    break;
                                }
                            case "B":
                                {
                                    sValue = reader.GetString(i);
                                    if (String.IsNullOrEmpty(sValue))
                                        sValue = "0";
                                    else
                                        sValue = sValue.Replace(',', '.');
                                    sValue = sValue.ToUpper();
                                    switch (sValue)
                                    {
                                        case "TRUE": sValue = "1"; break;
                                        case "FALSE":sValue = "0"; break;
                                        case "ИСТИНА":sValue = "1";break;
                                        case "ЛОЖЬ":sValue = "0";break;
                                        case "1": break;
                                        default: sValue = "0"; break;
                                    }
                                    sbTable.Append(sValue);
                                    break;
                                }
                        }
                    }

                    catch
                    {
                        // TODO: Корректная запись необработаной строки (иначе не отрабатывает ЗначениеИзФайла)
                        logFile.Add(String.Format("Ошибка преобразования значение {0} в тип {1}. Строка {2}, колонка {3}", sValue, saColumns[i], nRCount + 1, i), true);
            
                    }        
    
                    sbTable.Append('}');
                }
            }

            else
            {   // В случае если количество колонок не совпало с исходным - пишем пустую строку
                sbTable.Append(0);
                logFile.Add(String.Format("Различное количество колонок в таблице ({0}) и в файле ({1}), добавлена пустая строка ({2})", nCCount, nFieldsCount, nRCount + 1), true);
            }

            sbTable.Append(",0}");
            nRCount++;
        }
        public override string ToString()
        {
            sbTable.Insert(nRCountPlace, nRCount);
            InputEnd();
            return sbTable.ToString();
        }

        private void InputHeader()
        {
            sbTable.Append("{\"#\",acf6192e-81ca-46ef-93a6-5a6968b78663,{8,");
        }

        private void InputEnd()
        {
            sbTable.Append('}');
            sbTable.Append(',');
            sbTable.Append(nCCount-1);
            sbTable.Append(',');
            sbTable.Append(nRCount-1);
            sbTable.Append(new char[] { '}', '}', '}' } );
        }

    }
}
