using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DbHelpers;

 class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("NPOI-Template：");
            var DataSet = DbHelperSQL.Query(@"select 1 as [index],'good' b,'testc' c,'testd' d,'50%' p,'Successful' s,'remark content' r");

            if (DataSet.Tables[0].Rows.Count > 0)
            {
                var TemplatePath = Environment.CurrentDirectory + @"\Template\Template.xlsx";
                var Dictionary = new Dictionary<string, string>();
                Dictionary["no"] = "TEST NPOI NO.001 TEMPLATE";
                Dictionary["key1"] = "key1value";
                Dictionary["key2"] = "key2value";
                Dictionary["key3"] = "key3value";
                Dictionary["key4"] = "key4value";
                NPOItemplate.IworkSave(NPOItemplate.GenerateIWorkbook(TemplatePath, Dictionary, DataSet.Tables[0]), Environment.CurrentDirectory + @"..\..\..\OutPut\" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx");
                
                //if you want to return MemoryStream 
                //NPOItemplate.IworkToMemoryStream(NPOItemplate.GenerateIWorkbook(TemplatePath, Dictionary, DataSet.Tables[0]));
            }
            Console.WriteLine("Successful.");
        }
    }
