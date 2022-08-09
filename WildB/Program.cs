using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace WildB
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application eApp = new Excel.Application();
            Excel.Workbook wb = eApp.Workbooks.Add();
            Excel.Worksheet wsOne = wb.ActiveSheet;
            wsOne.Name = "Игрушки";
            Excel.Worksheet wsTwo = wb.Sheets.Add(After: wb.ActiveSheet);
            wsTwo.Name = "Настолки";
            Excel.Worksheet wsThree = wb.Sheets.Add(After: wsTwo);
            wsThree.Name = "Телефоны";

            string[] str = {"Title","Brand","Id","Feedbacks","Price" };
            for (int i = 1; i < 6; i++)
            {
                wsOne.Cells[1, i] = str[i-1];
                wsTwo.Cells[1, i] = str[i - 1];
                wsThree.Cells[1, i] = str[i - 1];
            }

            List<string> keys = new List<string>();
            
            using (StreamReader st = new StreamReader("Keys.txt"))
            {
                string line = string.Empty;
                while ((line = st.ReadLine()) != null)
                {
                    keys.Add(line);
                }
            }           
            

            foreach (var k in keys)
            {
                Console.WriteLine($"\n{k}\n");
                var req = new GetRequest($"https://search.wb.ru/exactmatch/sng/common/v4/search?appType=1&couponsGeo=12,7,3,21&curr=&dest=12358386,12358403,-70563,-8139704&emp=0&lang=ru&locale=by&pricemarginCoeff=1&query={k}&reg=0&regions=68,83,4,80,33,70,82,86,30,69,22,66,31,40,1,48&resultset=catalog&sort=popular&spp=0&suppressSpellcheck=false");
                
                req.Run();
                
                var json = JObject.Parse(req.Response);
                var data = json["data"];
                var products = data["products"];
                
                int i = 0;
                foreach(var prod in products)
                {
                    i++;
                    var tmp = new List<string>();
                    tmp.Add(prod["name"].ToString());
                    tmp.Add(prod["brand"].ToString());
                    tmp.Add(prod["id"].ToString());
                    tmp.Add(prod["feedbacks"].ToString());
                    tmp.Add(((Int32)prod["priceU"] / 100M).ToString());
                    
                    for (int j = 1; j < 6; j++) 
                    {
                        switch (k)
                        {
                            case "Игрушки": wsOne.Rows[i + 1].Columns[j] = tmp[j - 1]; break;
                            case "Настолки": wsTwo.Rows[i + 1].Columns[j] = tmp[j - 1]; break;
                            case "Телефоны": wsThree.Rows[i + 1].Columns[j] = tmp[j - 1]; break;
                        }
                                               
                        
                    }
                }            
                
            }
            eApp.AlertBeforeOverwriting = false;
            wb.SaveAs("out.xlsx");
            eApp.Quit();
            Console.Read();
        }        
    }
}
