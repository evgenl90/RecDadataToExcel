using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Net;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;



namespace RecDadataToExcel
{  
    class Program
    {      
        static void Main(string[] args)
        {
  
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();

            ex.Workbooks.Open(@"*****.xls",
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                false, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
 
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ex.Sheets[1];

            // Указываем номер столбца
            int numCol = 2;

            Excel.Range usedColumn = ObjWorkSheet.UsedRange.Columns[numCol];      
            System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
           
            var j = 0;
            string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();
            Console.WriteLine(strArray.Length);
            
            foreach (var i in strArray)
            {
              
                var str = strArray[j];

                var url = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party";

                //генерация запроса
                HttpWebRequest req = WebRequest.Create(url) as HttpWebRequest;
                req.Method = "POST";
                req.Timeout = 10000;
                req.Headers.Add("Authorization", "Token *******");
                req.ContentType = "application/json";
                req.Accept = "application/json";

                //данные для отправки
                var sentData = Encoding.UTF8.GetBytes("{ \"query\": \"" + strArray[j] + "\" }");
                req.ContentLength = sentData.Length;
                Stream sendStream = req.GetRequestStream();
                sendStream.Write(sentData, 0, sentData.Length);

                //получение ответа
                var res = req.GetResponse() as HttpWebResponse;
                var resStream = res.GetResponseStream();
                var sr = new StreamReader(resStream, Encoding.UTF8);

                var response = sr.ReadToEnd();

                ApiResponse apiResponse = JsonConvert.DeserializeObject<ApiResponse>(response);
                var x = 0;
                                          
                    foreach (var z in apiResponse.Suggestions)
                    {

                    // ОПФ
                    if (apiResponse.Suggestions[x].data.opf != null)
                    {
                        ObjWorkSheet.Cells[1, "Q"] = "ОПФ";
                        ObjWorkSheet.Cells[j + 1, "Q"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.opf.@short) ? "" : apiResponse.Suggestions[x].data.opf.@short;
                    }

                    //Полное Наименование
                    if (apiResponse.Suggestions[x].data.name != null)
                    {
                        ObjWorkSheet.Cells[1, "S"] = "Полное Наименование";
                        ObjWorkSheet.Cells[j + 1, "S"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.name.full_with_opf) ? "" : apiResponse.Suggestions[x].data.name.full_with_opf;
                    }                   

                    if (apiResponse.Suggestions[x].data != null)
                    { 
                        ObjWorkSheet.Cells[1, "R"] = "КПП";
                        ObjWorkSheet.Cells[j + 1, "R"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.kpp) ? "" : apiResponse.Suggestions[x].data.kpp;                 
                    }

                    if (apiResponse.Suggestions[x].data.address.data != null)
                    {
                        ObjWorkSheet.Cells[1, "T"] = "Регион";
                        ObjWorkSheet.Cells[j + 1, "T"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.address.data.region_with_type) ? "" : apiResponse.Suggestions[x].data.address.data.region_with_type;

                        ObjWorkSheet.Cells[1, "U"] = "Район";
                        ObjWorkSheet.Cells[j + 1, "U"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.address.data.city_district_with_type) ? "" : apiResponse.Suggestions[x].data.address.data.city_district_with_type;

                        ObjWorkSheet.Cells[1, "V"] = "Город";
                        ObjWorkSheet.Cells[j + 1, "V"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.address.data.city_with_type) ? "" : apiResponse.Suggestions[x].data.address.data.city_with_type;

                        ObjWorkSheet.Cells[1, "W"] = "Улица";
                        ObjWorkSheet.Cells[j + 1, "W"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.address.data.street_with_type) ? "" : apiResponse.Suggestions[x].data.address.data.street_with_type;

                        ObjWorkSheet.Cells[1, "X"] = "Дом";
                        ObjWorkSheet.Cells[j + 1, "X"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.address.data.house) ? "" : apiResponse.Suggestions[x].data.address.data.house;

                        ObjWorkSheet.Cells[1, "Y"] = "Корпус";
                        ObjWorkSheet.Cells[j + 1, "Y"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.address.data.block) ? "" : apiResponse.Suggestions[x].data.address.data.block;

                        ObjWorkSheet.Cells[1, "Z"] = "Квартира";
                        ObjWorkSheet.Cells[j + 1, "Z"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.address.data.flat) ? "" : apiResponse.Suggestions[x].data.address.data.flat;

                        ObjWorkSheet.Cells[1, "AA"] = "Индекс";
                        ObjWorkSheet.Cells[j + 1, "AA"] = String.IsNullOrEmpty(apiResponse.Suggestions[x].data.address.data.postal_code) ? "" : apiResponse.Suggestions[x].data.address.data.postal_code;

                    }          
                    x++;                
                }
               j++;
            }                  
           
        ex.Visible = true; //Отобразить Excel

        ex.DisplayAlerts = false; //Отключить отображение окон с сообщениями
        ex.Application.ActiveWorkbook.SaveAs("*****.xls", Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, false, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        ex.Quit();         
        }
    }
}

