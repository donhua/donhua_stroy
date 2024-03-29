using System;
using System.Security.Cryptography;
using System.Xml;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Word.Application;

namespace Donhua
{
    class PrintAct
    {
        public List<string> actList = new List<string>();
        Dictionary<int, string> data_act = new Dictionary<int, string>();
        public string actName = "aosr";


        public void GetConfig()
        {
            //получение значений полей класса из конфиг файла
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
        public void GetTabletValueOutConfig(string nameAct)
        {
            XmlDocument xDoc = new();
            xDoc.Load(@"config/config.xml");
            XmlElement? xRoot = xDoc.DocumentElement;
            if (xRoot != null)
            {
                foreach (XmlElement xnode in xRoot)
                {
                    if (xnode.Name == nameAct) {
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

        public static void CreateDitionaryForActValue()
        {
            //создать dict с парой значений "имя поля для вставки значения в акт" и "значение из базы"


        }


        public  void GetInExelValue() 
        {
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = xlsApp.Workbooks.Open(filename, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true);
            Sheets sheets = wb.Worksheets;
            Worksheet ws = (Worksheet)sheets.get_Item(1);

            Microsoft.Office.Interop.Excel.Range firstColumn = ws.UsedRange.Columns[1];
            Array myvalues = (Array)firstColumn.Cells.Value;
            string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();
        }

        public void Create_act()
        {
            //создаем акт по шаблону, подставляем значения в поля акта, сохраняем в выходную папку
            Application fileOpen = new Application();
            Document document = fileOpen.Documents.Open(@"C:\Users\Yakunin\source\repos\donhua_stroy\template\aosr.dotx", ReadOnly: false);
            fileOpen.Visible = true;
            document.Activate();
            Find findObject = fileOpen.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = "data_inform";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = "sdsdff";

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
            fileOpen.Selection.Find.Execute(ref findtext, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref reptext, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            document.SaveAs2(@"C:\Test.doc");
            fileOpen.Quit();
            Console.WriteLine("Операция сохранения прошла успешно!");
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
            PrintAct act = new PrintAct ();
            act.Create_act();
            Console.ReadLine();
        }

    }

}



