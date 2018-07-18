using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace XmlToSql
{
    class Program
    {
        //static string shablon = string.Format("{0,-5} | {1,-70} | {2,-5} | {3,-5} | {4,-5}", "AreaId", "Name", "IP", "PavilionId", "TypeId");
        static Model1 db = new Model1();
        static void Main(string[] args)
        {

            task4();
            //task56();
            //task7();
            //task8();
            //task9();
            //task10();
            //task11();

            return;

            ExcelPackage exp = new ExcelPackage();
            ExcelWorksheet worksheet = exp.Workbook.Worksheets.Add("List1");



            int row = 2;
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Column(2).Width = 50;
            worksheet.Cells[1, 3].Value = "IP";
            worksheet.Column(3).Width = 11;

            foreach (Area area in db.Area)
            {
                worksheet.Cells[row, 1].Value = area.AreaId;
                worksheet.Cells[row, 2].Value = area.FullName;
                worksheet.Cells[row, 3].Value = area.IP;
                row++;
            }


            /*4.	Создать метод, который возвращает данные в виде List<Area> */
            public static void task4()
            {
                List<Area> areas = db.Area.ToList();

                foreach (Area area in areas)
                    Console.WriteLine(area);
            }

            /*5.	Реализовать справочник, который возвращает ID зоны/участка,
             * и IP адрес данной зоны/участка. Так же необходимо исключить
             * зоны/участки у которых не заполнено поле IP
             6.	Реализовать справочник, который возвращает IP адрес и касс Area,
             исключить все зоны/участки, у которых отсутствует IPадрес, а так же
             исключить все дочерние зоны/участки (ParentId!=0)*/

            public static void task56()
            {
                Dictionary<string, Area> dicIP = db.Area.
                    Where(w => !string.IsNullOrEmpty(w.IP)
                    && w.ParentId != 0)
                    .Select(s => new { s.IP }).Distinct()
                    .Select(s => new
                    {
                        ip = s.IP,
                        area = db.Area.FirstOrDefault(f => f.IP == s.IP)
                    })
                    .ToDictionary(d => d.ip, d => d.area);

                foreach (var item in dicIP)
                {
                    Console.WriteLine(item.Key + "\t" + item.Value);
                }
            }

            /*7.	Используя коллекцию Lookup, вернуть следующие данные. В качестве
             * ключа использовать IP адрес, в качестве значения использовать класс Area*/

            public static void task7()
            {
                ILookup<string, Area> lkp = db.Area.Where(w => !string.IsNullOrEmpty(w.IP)).ToLookup(l => l.IP, l => l);

                foreach (var item in lkp)
                {
                    Console.Write(item.Key + "\t\t");
                    foreach (Area i in item)
                    {
                        Console.Write(i); break;
                    }
                    Console.WriteLine();
                }
            }


            /*8.	Вернуть первую запись из последовательности, где HiddenArea=1*/
            public static void task8()
            {
                Console.WriteLine(db.Area.ToList().Where(o => o.HiddenArea == true).FirstOrDefault());
                Console.WriteLine();
            }

            /*9.	Вернуть последнюю запись из таблицы Area, указав следующий фильтр – PavilionId = 1*/

            public static void task9()
            {
                //Console.WriteLine(shablon);
                Console.WriteLine(db.Area.ToList().Where(o => o.PavilionId == 1).LastOrDefault());
            }


            /*10.	Используя квантификаторы, вывесит на экран значения следующих фильтров:
             a.	Есть ли в таблице зоны/участки для PavilionId = 1 и IP = 10.53.34.85, 10.53.34.77, 10.53.34.53
             b.	Содержатся ли данные в таблице Area с наименованием зон/участков - PT disassembly, Engine testing*/

            public static void task10()
            {
                List<Area> areas = db.Area.Where(o => o.PavilionId == 1 && o.IP.Contains("10.53.34.85") || o.IP.Contains("10.53.34.77") || o.IP.Contains("10.53.34.53")).ToList();

                if (areas.Count > 0)
                    Console.WriteLine("В таблице Ареа содержатся участки с такими айпи\n");
                else
                    Console.WriteLine("В таблице Ареа не содержатся участки с такими айпи\n");

                foreach (Area item in areas)
                {
                    Console.WriteLine(item);
                }
                Console.WriteLine("\n------------------------------------------------------------------\n");

                List<Area> areas1 = db.Area.Where(o => o.Name == "PT disassembly" || o.Name == "Engine testing").ToList();

                if (areas1.Count > 0)
                    Console.WriteLine("В таблице Ареа содержатся участки, содержаюшие PT disassembly или  Engine testing\n");
                else
                    Console.WriteLine("В таблице Ареа не содержатся участки, содержаюшие PT disassembly или  Engine testing\n");

                foreach (Area item in areas1)
                {
                    Console.WriteLine(item);
                }
            }

        /*11.	Вывести сумму всех работающих работников (WorkingPeople) на зонах*/

            public static void task11()
            {
                Console.WriteLine(db.Area.Sum(o => o.WorkingPeople));
            }

            ExcelWorksheet worksheet2 = exp.Workbook.Worksheets.Add("List2");
            row = 2;

            //foreach (var item in dicIP)
            //{
            //    worksheet2.Cells[row, 1].Value = item.Key;
            //    worksheet.Cells[row, 2].Value = item.Value.FullName;
            //    row++;
            //}

            FileStream fs = File.Create("Excl.xlsx");
            fs.Close();
            exp.SaveAs(fs);
        }
    }
}
