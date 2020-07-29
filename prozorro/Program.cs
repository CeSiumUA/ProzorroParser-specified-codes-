using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace prozorro
{
    class Program
    {

        public const string webURL = "https://public.api.openprocurement.org";

        public const string listOfTenders = "/api/2.5/tenders?limit=1000&offset=";
        /*https://public.api.openprocurement.org/api/2.5/tenders?limit=1000&offset=2019-08-25*/

        static List<string> Codes = new List<string>();

        static List<string> EDRPOUCodes = new List<string>();

        static string ExcelFileName = "";

        static string YEAR;
        static string MONTH;
        static string DAY;
        static void Main(string[] args)
        {
            //   string a1 = "\u0410\u0441\u0444\u0430\u043b\u044c\u0442";

            //   Console.WriteLine(System.Text.RegularExpressions.Regex.Unescape(a1));

            // Console.ReadLine();
            try
            {
                //Console.WriteLine(DateTime.Now.ToString());
                
                for (int x = 0; x < codesList.Length; x++)
                {
                    if (!Codes.Contains(codesList[x]))
                    {
                        Codes.Add(codesList[x]);
                    }
                }
                Directory.CreateDirectory("ParsedExcels");
                DateTime dt = DateTime.Now;

                ExcelFileName = Directory.CreateDirectory("ParsedExcels").FullName + @"\ParsedExcel " + dt.Day.ToString() + ", " + dt.Month.ToString() + ", " + dt.Year.ToString() + " " + dt.Hour + "," + dt.Minute + "," + dt.Second + ".xlsx";

                using (StreamReader sr = new StreamReader("NumbersFilter.txt"))
                {
                    string line;

                    while ((line = sr.ReadLine()) != null)
                    {
                        if (!Codes.Contains(line))
                        {
                            Codes.Add(line);
                        }
                    }
                }
                Console.WriteLine("Використовувати ЄДРПОУ? (y/n)");
                bool useEDRPOU = (Console.ReadLine().ToLower() == "y") ? true : false;
                if (useEDRPOU)
                {
                    using (StreamReader sr = new StreamReader("EDRPOU.txt"))
                    {
                        string line;

                        while ((line = sr.ReadLine()) != null)
                        {
                            if (!EDRPOUCodes.Contains(line))
                            {
                                EDRPOUCodes.Add(line);
                            }
                        }
                    }
                }
                else
                {
                    EDRPOUCodes.Clear();
                    EDRPOUCodes.Add("");
                }
                List<string> columnsnames = new List<string>();
                using (StreamReader sr = new StreamReader("ColumnNames.txt"))
                {
                    string line;

                    while ((line = sr.ReadLine()) != null)
                    {
                        columnsnames.Add(line);
                    }
                }

                int year = DateTime.Now.Year;

                int month = DateTime.Now.Month;

                int day = DateTime.Now.Day;//- 1;


                YEAR = year.ToString();

                MONTH = month.ToString();

                DAY = day.ToString();

                if (month < 10)
                {
                    MONTH = '0' + MONTH;
                }

                if (day < 10)
                {
                    DAY = '0' + DAY;
                }

                string resp = "";

                string date = YEAR + "-" + MONTH + "-" + DAY;

                Console.WriteLine("Введіть початкову дату (YYYY-MM-DD), (натисніть Enter, якщо треба парсити починаючи з сьогоднішньої дати):");
                string vb = "";
                if ((vb = Console.ReadLine()) != "")
                {
                    date = vb;
                }
                Console.WriteLine("Почати з вказаного часу? (y/n)");
                if (Console.ReadLine().ToLower() == "y")
                {
                    Console.WriteLine("Введіть дату у форматі: yyyy-MM-dd hh:mm:ss");
                    var targetTime = DateTime.ParseExact(Console.ReadLine(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    while (DateTime.Now.ToString() != targetTime.ToString())
                    {

                    }

                }
                sheet = wb.CreateSheet(YEAR + "-" + MONTH + "-" + DAY);

                IRow row = sheet.CreateRow(0);

                int cellNameNo = 0;

                foreach (string x in columnsnames)
                {
                    row.CreateCell(cellNameNo).SetCellValue(x);
                    cellNameNo++;
                }

                WebRequest request0 = WebRequest.Create(webURL + listOfTenders + date);

                request0.Credentials = CredentialCache.DefaultCredentials;

                HttpWebResponse response0 = (HttpWebResponse)request0.GetResponse();

                Console.WriteLine(response0.StatusDescription);

                Stream dataStream0 = response0.GetResponseStream();

                StreamReader reader0 = new StreamReader(dataStream0);

                string responseFromServer0 = reader0.ReadToEnd();

                resp = responseFromServer0;


                Console.WriteLine(responseFromServer0);

                StartPage sp0 = JsonConvert.DeserializeObject<StartPage>(responseFromServer0);

                for (int x = 0; x < sp0.data.Length; x++)
                {
                    //     CheckID(webURL + "/api/2.5/tenders/" + sp.data[x].id);
                    Console.WriteLine(sp0.data[x].dateModified);
                }

                // Cleanup the streams and the response. 
                reader0.Close();
                dataStream0.Close();
                response0.Close();

                string nextPage = sp0.next_page.path;
                List<Task> RunningTasks = new List<Task>();
                while (!resp.Contains("\"data\": [], \"prev_page\""))
                {

                    WebRequest request = WebRequest.Create(webURL + nextPage);

                    request.Credentials = CredentialCache.DefaultCredentials;

                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                    Console.WriteLine(response.StatusDescription);

                    Stream dataStream = response.GetResponseStream();

                    StreamReader reader = new StreamReader(dataStream);

                    string responseFromServer = reader.ReadToEnd();

                    resp = responseFromServer;

                    reader.Close();
                    dataStream.Close();
                    response.Close();

                    Task parsingTask = new Task(() =>
                    {


                        StartPage sp = JsonConvert.DeserializeObject<StartPage>(responseFromServer);

                        nextPage = sp.next_page.path;

                        for (int x = 0; x < sp.data.Length; x++)
                        {
                            CheckID(webURL + "/api/2.5/tenders/" + sp.data[x].id, useEDRPOU);
                            //Console.WriteLine(sp.data[x].dateModified);
                        }
                    });
                    RunningTasks.Add(parsingTask);
                    parsingTask.Start();
                    
                }
                Task.WaitAll(RunningTasks.ToArray());
                using (FileStream fs = new FileStream(ExcelFileName, FileMode.Create, FileAccess.Write))
                {
                    wb.Write(fs);
                }
            }
            catch(Exception er)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(er.ToString());
                Console.ForegroundColor = ConsoleColor.White;
            }
            Console.WriteLine("Done!");
            Console.ReadLine();
        }

        private static void CheckID(string addons, bool UseEdrpou)
        {
            WebRequest request = WebRequest.Create(addons);

            request.Credentials = CredentialCache.DefaultCredentials;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            Console.WriteLine(response.StatusDescription);

            Stream dataStream = response.GetResponseStream();

            StreamReader reader = new StreamReader(dataStream);

            string responseFromServer = reader.ReadToEnd();

            //string testString = "\"startDate\": \"" + YEAR + "-" + MONTH + "-" + DAY;

            if (strinContainsElement(responseFromServer))
            {

              //  Console.Beep(8000, 2000);

                Console.ForegroundColor = ConsoleColor.Red;

                Console.WriteLine("//--------------------------------------------------------------------//");
                //Console.WriteLine(/*System.Text.RegularExpressions.Regex.Unescape(*/responseFromServer/*).Replace('?', 'І')*/);

                // JObject obj = JObject.Parse(responseFromServer);

                // Console.WriteLine(obj["id"]);

                Parser(responseFromServer, UseEdrpou);

                Console.ForegroundColor = ConsoleColor.White;
            }


        }

        static bool strinContainsElement(string strin)
        {

             // if (strin.Contains("\"id\": \"48329000-0\"") || strin.Contains("\"id\": \"72252000-6\"") || strin.Contains("\"id\": \"79995100-6\"") || strin.Contains("\"id\": \"92500000-6\"") || strin.Contains("\"id\": \"92510000-9\"") || strin.Contains("\"id\": \"92512000-3\"") || strin.Contains("\"id\": \"64216200-5\"") || strin.Contains("\"id\": \"79999000-3\"") || strin.Contains("\"id\": \"79999100-4\"") /*|| strin.Contains("\"id\": \"72252000-6\"")*/)
             // {
             //     return true;
            //  }

            JObject jo = JObject.Parse(strin);

            for(int x = 0; x < Codes.Count; x++)
            {
                JArray ja = (JArray)jo["data"]["items"];
                for(int y = 0; y < ja.Count; y++)
                {
                    if(ja[y]["classification"]["id"].ToString() == Codes[x])
                    {
                        return true;
                    }
                }
            }


            return false;
        }

        //public static string[] Codes = new string[] { "\"id\": \"72252000-6\"", "\"id\": \"48329000-0\"", "\"id\": \"79995100-6\"", "\"id\": \"92500000-6\"", "\"id\": \"92510000-9\"", "\"id\": \"92512000-3\"", "\"id\": \"64216200-5\"", "\"id\": \"79999000-3\"", "\"id\": \"79999100-4\"" };

        public static string[] codesList = new string[] { "72252000-6", "48329000-0", "79995100-6", "92500000-6", "92510000-9", "92512000-3", "64216200-5", "79999000-3", "79999100-4" };

        static int incrementInExcel = 1;

      /*  static void Parser0(string response)
        {
            Data dat = new Data();

            JObject jo = JObject.Parse(response);

            dat.AddDate = YEAR + "-" + MONTH + "-" + DAY;

            dat.Number = jo["data"]["tenderID"].ToString();

            Console.WriteLine(response);

            Data[] data = null;

            if (response.Contains("\"lots\"")) {
                JArray ja = (JArray)jo["data"]["lots"];

                data = new Data[ja.Count];

                for (int x = 0; x < ja.Count; x++)
                {


                    Data dt = new Data();
                    dt.Budget = ja[x]["value"]["amount"].ToString(); //+ " " + ja[x]["value"]["currency"] + ", " + "Tax included: " + ja[x]["value"]["valueAddedTaxIncluded"];
                    dt.AddDate = dat.AddDate;
                    dt.Number = dat.Number;
                    dt.EndDate = jo["data"]["tenderPeriod"]["endDate"].ToString();//ja[x]["auctionPeriod"]["endDate"].ToString();
                    dt.EDRPOU = jo["data"]["procuringEntity"]["identifier"]["id"].ToString();
                    dt.Client = jo["data"]["procuringEntity"]["name"].ToString();
                    dt.URL = @"https://www.prozorro.gov.ua/tender/" + dat.Number;
                    dt.Comments = ja[x]["title"].ToString();
                    dt.Item = ja[x]["title"].ToString();

                    data[x] = dt;
                }
            }
            else
            {
                JArray ja = (JArray)jo["data"]["awards"];

                data = new Data[ja.Count];

                string amount = "";

                for(int x = 0; x < ja.Count; x++)
                {
                    if(ja[x]["status"].ToString() != "unsuccessful")
                    {
                        amount = ja[x]["value"]["amount"].ToString();
                    }

                }

               // for (int x = 0; x < ja.Count; x++)
               // {


                    Data dt = new Data();
                    dt.Budget = amount; //+ " " + ja[x]["value"]["currency"] + ", " + "Tax included: " + ja[x]["value"]["valueAddedTaxIncluded"];
                    dt.AddDate = dat.AddDate;
                    dt.Number = dat.Number;
                    dt.EndDate = jo["data"]["tenderPeriod"]["endDate"].ToString();//ja[x]["auctionPeriod"]["endDate"].ToString();
                    dt.EDRPOU = jo["data"]["procuringEntity"]["identifier"]["id"].ToString();
                    dt.Client = jo["data"]["procuringEntity"]["name"].ToString();
                    dt.URL = @"https://www.prozorro.gov.ua/tender/" + dat.Number;
                //dt.Comments = jo["title"].ToString();
                    Console.WriteLine(jo["title"]);
                    //dt.Item = ja[x]["title"].ToString();
                    dt.Item = "";

                    data[0] = dt;
                //}
            }

            JArray ja0 = (JArray)jo["data"]["items"];
            for (int y = 0; y < ja0.Count; y++)
            {


                Data dt = data[y];

                dt.Item = ja0[y]["classification"]["description"].ToString() + ", " + ja0[y]["quantity"] + " шт.";

                data[y] = dt;
            }

            for (int a = 0; a < data.Length; a++)
            {

                WriteToExcel(data[a]);
            }

            incrementInExcel += data.Length;


            using (FileStream fs = new FileStream("res.xlsx", FileMode.Create, FileAccess.Write))
            {
                wb.Write(fs);
            }

           
        }*/

        static void Parser(string responseFromServer0, bool UseEdrpou)
        {
            Data dat = new Data();

            JObject jo = JObject.Parse(responseFromServer0);

            JToken ji = jo["data"];
            dat.ProzorroLink = "https://prozorro.gov.ua/tender/" + ji["tenderID"].ToString();
            try
            {
                dat.AddDate = ji["date"].ToString();//ji["tenderPeriod"]["startDate"].ToString();
            }
            catch
            {
                dat.AddDate = "0";
            }

            try
            {
                dat.ExpireDate = ji["tenderPeriod"]["endDate"].ToString();
            }
            catch
            {
                dat.ExpireDate = "0";    
            }
            try
            {
                dat.enqueryPeriod = ji["enquiryPeriod"]["startDate"].ToString();
            }
            catch
            {
                dat.enqueryPeriod = "0";
            }

            try
            {
                dat.clarificationUntil = ji["enquiryPeriod"]["clarificationsUntil"].ToString();
            }
            catch
            {
                dat.clarificationUntil = "0";
            }
           

            try
            {
                dat.awardPeriodStartDate = ji["awardPeriod"]["startDate"].ToString();
            }
            catch
            {
                dat.awardPeriodStartDate = "0";
            }

            try
            {
                dat.Code = ji["id"].ToString();
            }
            catch
            {
                dat.Code = "0";
            }

            try
            {
                dat.value_amount = ji["value"]["amount"].ToString();
            }
            catch
            {
                dat.value_amount = "0";
            }
            JArray ja = (JArray)ji["items"];

            Data[] dt = new Data[ja.Count];

            for (int x = 0; x < dt.Length; x++)
            {
                Data d1 = new Data();
                d1.ProzorroLink = dat.ProzorroLink;
                d1.AddDate = dat.AddDate;

                d1.ExpireDate = dat.ExpireDate;

                d1.enqueryPeriod = dat.enqueryPeriod;

                d1.clarificationUntil = dat.clarificationUntil;

                d1.awardPeriodStartDate = dat.awardPeriodStartDate;

                d1.Code = dat.Code;

                d1.value_amount = dat.value_amount;

                try
                {
                    d1.Class_ID = ja[x]["classification"]["id"].ToString();
                }
                catch
                {
                    d1.Class_ID = "0";
                }

                try
                {
                    d1.classification_description = ja[x]["classification"]["description"].ToString();
                }
                catch
                {
                    d1.classification_description = "0";
                }

                try
                {
                    d1.description = ja[x]["description"].ToString();
                }
                catch
                {
                    d1.description = "0";
                }

                try
                {
                    d1.quantity = ja[x]["quantity"].ToString();
                }
                catch
                {
                    d1.quantity = "0";
                }

                try
                {
                    d1.CodePostpayment = ji["milestones"][0]["code"].ToString();
                }
                catch
                {
                    d1.CodePostpayment = "0";
                }

                try
                {
                    d1.deliveryDate = ja[x]["deliveryDate"]["endDate"].ToString();
                }
                catch
                {
                    d1.deliveryDate = "0";
                }

                try
                {
                    d1.Ident_ID = ji["procuringEntity"]["identifier"]["id"].ToString();
                }
                catch
                {
                    d1.Ident_ID = "0";
                }

                try
                {
                    d1.Ident_LegalName = ji["procuringEntity"]["identifier"]["legalName"].ToString();
                }
                catch
                {
                    d1.Ident_LegalName = "0";
                }

                try
                {
                    d1.contactPoint = ji["procuringEntity"]["contactPoint"]["name"].ToString() + ", " + ji["procuringEntity"]["contactPoint"]["telephone"].ToString() + ", " + ji["procuringEntity"]["contactPoint"]["email"].ToString();
                }
                catch
                {
                    d1.contactPoint = "0";
                }

                try
                {
                    d1.Address = ji["procuringEntity"]["address"]["countryName"].ToString() + ", " + ji["procuringEntity"]["address"]["locality"].ToString() + ", " + ji["procuringEntity"]["address"]["region"].ToString() + ", " + ji["procuringEntity"]["address"]["postalCode"].ToString() + ", " + ji["procuringEntity"]["address"]["streetAddress"].ToString();
                }
                catch
                {
                    d1.Address = "0";
                }
                dt[x] = d1;
            }

            for (int a = 0; a < dt.Length; a++)
            {
                if (UseEdrpou)
                {
                    if (EDRPOUCodes.Contains(dt[a].Ident_ID))
                    {
                        WriteToExcel(dt[a]);
                    }
                }
                else
                {
                    //if (Convert.ToDouble(dt[a].value_amount) > 500000)
                    //{
                        WriteToExcel(dt[a]);
                    //}
                }
            }

            //using (FileStream fs = new FileStream(ExcelFileName, FileMode.Create, FileAccess.Write))
            //{
            //    wb.Write(fs);
            //}
        }

        static void WriteToExcel(Data dat)
        {
            lock (sheet)
            {
                int num = incrementInExcel;

                sheet.CreateRow(num);
                for (int i = 0; i < 18; i++)
                {
                    sheet.GetRow(num).CreateCell(i);
                }

                sheet.GetRow(num).GetCell(0).SetCellValue(dat.AddDate);

                sheet.GetRow(num).GetCell(1).SetCellValue(dat.ExpireDate);

                sheet.GetRow(num).GetCell(2).SetCellValue(dat.enqueryPeriod);

                sheet.GetRow(num).GetCell(3).SetCellValue(dat.clarificationUntil);

                sheet.GetRow(num).GetCell(4).SetCellValue(dat.awardPeriodStartDate);

                sheet.GetRow(num).GetCell(5).SetCellValue(dat.Code);

                sheet.GetRow(num).GetCell(6).SetCellValue(dat.value_amount);

                sheet.GetRow(num).GetCell(7).SetCellValue(dat.Class_ID);

                sheet.GetRow(num).GetCell(8).SetCellValue(dat.classification_description);

                sheet.GetRow(num).GetCell(9).SetCellValue(dat.description);

                sheet.GetRow(num).GetCell(10).SetCellValue(dat.quantity);

                sheet.GetRow(num).GetCell(11).SetCellValue(dat.CodePostpayment);

                sheet.GetRow(num).GetCell(12).SetCellValue(dat.deliveryDate);

                sheet.GetRow(num).GetCell(13).SetCellValue(dat.Ident_ID);

                sheet.GetRow(num).GetCell(14).SetCellValue(dat.Ident_LegalName);

                sheet.GetRow(num).GetCell(15).SetCellValue(dat.contactPoint);

                sheet.GetRow(num).GetCell(16).SetCellValue(dat.Address);

                sheet.GetRow(num).GetCell(17).SetCellValue(dat.ProzorroLink);

                incrementInExcel++;
            }
        }

        static XSSFWorkbook wb = new XSSFWorkbook();

        static ISheet sheet;

        public struct StartPage
        {
            public NextPage next_page;

            public DataSample[] data;
            public struct NextPage
            {
                public string path;

                public string uri;

                public string offset;
            }

            public struct DataSample
            {
                public string id;

                public string dateModified;
            }
        }

        public struct Data
        {
            public string AddDate;

            public string ExpireDate;

            public string enqueryPeriod;

            public string clarificationUntil;

            public string awardPeriodStartDate;

            public string Code;

            public string value_amount;

            public string Class_ID;

            public string classification_description;

            public string description;

            public string quantity;

            public string CodePostpayment;

            public string deliveryDate;

            public string Ident_ID;

            public string Ident_LegalName;

            public string contactPoint;

            public string Address;

            public string ProzorroLink;
        }

    }
}
