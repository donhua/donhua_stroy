using System;
using System.IO;
using System.Xml;

namespace Donhua
{
    class PrintAct
    {
        //тут конструктор класса

        public void GetConfig()
        {
            //получение значений полей класса из конфиг файла
            PathCollection path_c = new PathCollection();
            XmlDocument xDoc = new XmlDocument();
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

        public void GetTabletValueOutConfig(string nameAct)
        {
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(@"config/config.xml");
            XmlElement? xRoot = xDoc.DocumentElement;
            if (xRoot != null)
            {
                foreach (XmlElement xnode in xRoot)
                {
                    if (xnode.Name == nameAct)
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

        public void CreateDitionaryForActValue()
        {
            //создать dict с парой значений "имя поля для вставки значения в акт" и "значение из базы"

        }

        public void create_act()
        {
            //создаем акт по шаблону, подставляем значения в поля акта, сохраняем в выходную папку
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
            act.GetConfig();
        }

    }

}



