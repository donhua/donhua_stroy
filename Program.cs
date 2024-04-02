using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Xml;
using Application = Microsoft.Office.Interop.Word.Application;

namespace Donhua
{
    class PrintAct
    ///TODO добавить логирование
    ///TODO переработать перерменные
    {
        public List<string> actList = new List<string>();
        Dictionary<int, string> data_act = new Dictionary<int, string>();
        public string actName = "aosr";

        /// <summary>
        /// Распарсивает xml
        /// </summary>
        /// HACK временно провисан путь к xml
        /// TODO вывести путь к xml в параметр метода
        /// TODO добавить проверку файла
        public void GetConfig()
        {
            PathCollection path_c = new();
            XmlDocument xDoc = new();
            xDoc.Load(@"config/config.xml");
            XmlElement? xRoot = xDoc.DocumentElement;
            if (xRoot != null)
            {
                foreach (XmlElement xnode in xRoot)
                {
                    if (xnode.Name == "pathes")
                    {
                        foreach (XmlNode childnode in xnode.ChildNodes)
                        {
                            switch (childnode.Name)
                            {
                                case "path_act_base_folder":
                                    path_c.path_act_base_folder = childnode.InnerText;
                                    break;
                                case "path_config":
                                    path_c.path_act_base_folder = childnode.InnerText;
                                    break;
                                case "path_result":
                                    path_c.path_act_base_folder = childnode.InnerText;
                                    break;
                            }
                        }
                    }

                }
            }
        }

        /// <summary>
        /// Формирует список имен полей документа для последующей вставки значений в поле
        /// </summary>
        /// <param name="nameAct">название акта для которого формируем список</param>
        /// <returns>список полей string</returns>
        /// HACK прописан путь к файлу
        /// TODO вывести путь в параметр и добавить проверку наличия файла
        /// TODO После отладку убрать Console.WriteLine()
        public void GetTabletValueOutConfig(string nameAct)
        {
            XmlDocument xDoc = new();
            xDoc.Load(@"config/config.xml");
            XmlElement? xRoot = xDoc.DocumentElement;
            if (xRoot != null)
            {
                foreach (XmlElement xnode in xRoot)
                {
                    if (xnode.Name == nameAct)
                    {
                        Console.WriteLine($"Элемент {xnode.Name} найден");
                        foreach (XmlNode childnode in xnode)
                        {
                            Console.WriteLine(childnode.Name);
                            string a = childnode.Name;
                            actList.Add(a);
                        }
                        Console.WriteLine(actList.Count);
                        Console.WriteLine(string.Join(", ", actList.ToArray()));
                    }
                    else { Console.WriteLine("Отчет в xml не найден!!!"); }
                }
            }
        }

        /// <summary>
        /// генерирует из настроек словарь ключ - поле для замены с пустыми значениями
        /// </summary>
        /// <returns>Dictionary<string, string></returns>
        /// TODO реализовать метод
        /// HACK выводит пустой словарь
        public Dictionary<string, string> CreateDitionaryForActValue()
        {
            //создать dict с парой значений "имя поля для вставки значения в акт" и "значение из базы"
            Dictionary<string, string> dict = new Dictionary<string, string>();
            return dict;
        }

        /// <summary>
        /// принимает значения из листов Exceld в словарь
        /// </summary>
        /// <param name="filename">путь к файлу эксель</param>
        /// <return>Dictionary<string, string></return>
        /// UNDONE
        public void GetInExelValue(string filename)
        {
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = xlsApp.Workbooks.Open(filename, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true);
            Sheets sheets = wb.Worksheets;
            Worksheet ws = (Worksheet)sheets.get_Item(1);

            Microsoft.Office.Interop.Excel.Range firstColumn = ws.UsedRange.Columns[1];
            Array myvalues = (Array)firstColumn.Cells.Value;
            string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();
        }

        /// <summary>
        /// в шаблоне word производит замену слова
        /// </summary>
        /// <param name="text_find">слово которое заменим</param>
        /// <param name="text_replace">слово на которое заменим</param>
        /// <param name="app">передаем приложение (word)</param>
        /// <param name="doc">передаем документ в котором проводим правки</param>
        private void Replaser(string text_find, string text_replace, Application app, Document doc)
        {
            //создаем акт по шаблону, подставляем значения в поля акта, сохраняем в выходную папку
            doc.Activate(); //TODO вывести из метода
            Find findObject = app.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = text_find;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = text_replace;

            object findtext = findObject.Text;
            object reptext = findObject.Replacement.Text;
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            app.Selection.Find.Execute(ref findtext, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref reptext, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            Console.WriteLine($"Замена {text_find} на {text_replace} прошла успешно!");
        }

        /// <summary>
        /// Генерирует имя результирующего файла
        /// </summary>
        /// <returns>string</returns>
        /// HACK возвращает строку за раннее прописанную в методе
        /// TODO проработать механизм генерации имени файлов!
        private string GenerateName()
        {
            string name = "ку";
            return name;
        }

        /// <summary>
        /// Генерирует документы по списку и сохраняет их.
        /// </summary>
        /// <param name="list_info">список с набором словарей 
        /// где ключ - заменяемое слово, значение - то, чем будет заменено</param>
        /// <param name="path_act">массив с путями к шаблонам документов</param>
        /// TODO доработать распарсивание словарей из списка
        /// TODO добавить цикл по генерации документов по количеству видов документов
        /// TODO добавить цикл по генерации документов по количеству словарей в списке
        /// UNDONE !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        public void CreateAktDoc(List<Dictionary<string, string>> list_info, string[] path_act)
        {
             Application app = new Application();/*
             for (int j = 0; j < path_act.Length; j++)
             {

                 Document doc = app.Documents.Open(path_act[j], ReadOnly: false);
                 app.Visible = true;
                 doc.Activate();
                 for (int i = 0; i < arrFind.Length; i++)
                 {
                     Replaser(arrFind[i], arrReplace[i], app, doc);
                 }
                 doc.SaveAs2($@"C:\{GenerateName()}{j}.docx");

                 Console.WriteLine($"сохранение {GenerateName()}{j}.docx после замены прошло успешно!"); ;
             }*/
             app.Quit();

        }

    }


    class PathCollection
    {
        public string? path_act_base_folder { get; set; }
        public string? path_config { get; set; }
        public string? path_result { get; set; }
    }

    class Programm
    {
        static void Main()
        {
            PrintAct act = new PrintAct();
            Dictionary<string, string> act_info1 = new Dictionary<string, string>()
            {
                ["alfa"] = "альфа1",
                ["beta"] = "бета1",
            };
            Dictionary<string, string> act_info2 = new Dictionary<string, string>()
            {
                ["alfa"] = "альфа2",
                ["beta"] = "бета2",
            };
            List<Dictionary<string, string>> list_info = new List<Dictionary<string, string>>();
            list_info.Add(act_info1);
            list_info.Add(act_info2);

            string[] path_act = { @"C:\Users\Yakunin\source\repos\donhua_stroy\template\aosr1.dotx" };

            act.CreateAktDoc(list_info, path_act);
            Console.ReadLine();
        }

    }

}



