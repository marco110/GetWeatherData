using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace GetWeatherData
{
    class Program
    {
        private const string baseURL = "http://tianqi.2345.com/t/wea_history/js/";

        static void Main(string[] args)
        {
            List<List<string>> citiesData = CountiesData();

            List<CountyInfo> countiesList = AddCountyToCountyInfo(citiesData);
            List<string> names = new List<string>();
            foreach (var item in countiesList)
            {
                names.Add(item.Name);
            }
            File.WriteAllLines(@"C:\Users\Administrator\Desktop\citynames.json", names);
            citiesData.Clear();

            bool needContinue = true;
            int year = 2011;

            while (needContinue)
            {
                Console.WriteLine("开始获取" + year + "年数据！");

                List<string> stringData = GetWeatherDataByYear(countiesList, year);

                string output = Environment.CurrentDirectory + "\\results\\WeatherData_" + year + ".json";

                File.WriteAllLines(output, stringData);

                Console.WriteLine(year + "年数据获取完毕!");
                stringData.Clear();
                year++;

                if (year == 2017)
                    needContinue = false;

            }

            Process.Start(@"C:\Users\Administrator\Desktop\SaveWeatherData");

            Console.ReadKey();
        }

        private static List<string> GetWeatherDataByYear(List<CountyInfo> countiesList, int Year)
        {
            List<string> noDataCountyList = new List<string>();
            Dictionary<string, string> noDataCounty = new Dictionary<string, string>();
            List<string> stringData = new List<string>();
            int count = 0;

            if (Year < 2016)
            {
                for (int i = 0; i < countiesList.Count; i++)
                {
                    count++;

                    for (int year = Year; year < Year + 1; year++)
                    {
                        for (int month = 1; month < 13; month++)
                        {
                            string getDataURL = baseURL + countiesList[i].ID + "_" + year + month + ".js";

                            getDataFromWeb(countiesList, noDataCountyList, stringData, count, i, year, month, getDataURL);
                        }
                    }
                }
            }
            else if (Year == 2016)
            {
                for (int i = 0; i < countiesList.Count; i++)
                {
                    count++;

                    for (int year = Year; year < Year + 1; year++)
                    {
                        for (int month = 1; month < 13; month++)
                        {
                            string getDataURL = baseURL + countiesList[i].ID + "_" + year + month + ".js";

                            if (month < 10 && month > 3)
                            {
                                getDataURL = baseURL + year + "0" + month + "/" + countiesList[i].ID + "_" + year + "0" + month + ".js";
                            }
                            else if (month >= 10)
                            {
                                getDataURL = baseURL + year + month + "/" + countiesList[i].ID + "_" + year + month + ".js";
                            }

                            getDataFromWeb(countiesList, noDataCountyList, stringData, count, i, year, month, getDataURL);
                        }
                    }
                }
            }
            noDataCountyList.Add("总条数：" + noDataCountyList.Count);

            File.WriteAllLines(Environment.CurrentDirectory + "\\results\\log_" + Year + ".txt", noDataCountyList);
            return stringData;
        }

        private static void getDataFromWeb(List<CountyInfo> countiesList, List<string> noDataCountyList, List<string> stringData, int count, int i, int year, int month, string getDataURL)
        {
            using (var httpClient = new HttpClient())
            {
                httpClient.BaseAddress = new Uri(getDataURL);
                httpClient.Timeout = new TimeSpan(0, 0, 10);
                httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36");

                try
                {
                    var statusCode = httpClient.GetAsync(getDataURL).Result.StatusCode;

                    if (statusCode.ToString() == "OK")
                    {
                        var temResult = httpClient.GetByteArrayAsync(getDataURL).Result;

                        var result = Encoding.Default.GetString(temResult);

                        var temRegexResult = Regex.Replace(result, "var weather_str=", "");

                        result = Regex.Replace(temRegexResult, @"(avgyWendu:'-?\d+'\});", @"${1}");
                        result = Regex.Replace(result, ";", "");

                        stringData.Add(result);

                        Console.WriteLine(count + "/" + 374 + countiesList[i].Name + year + "年" + month + "月");
                    }

                    else
                    {
                        string log = countiesList[i].Name + " 于 " + year.ToString() + "年" + month.ToString() + "月 无数据！\n";
                        Console.WriteLine(count + "/" + 374);
                        Console.WriteLine(countiesList[i].Name + year + month + "无数据！");
                        noDataCountyList.Add(log);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(countiesList[i].Name + " 于 " + year + month + " " + ex.Message);
                    noDataCountyList.Add(countiesList[i].Name + " 于 " + year + month + " " + ex.Message);
                }
            }
        }

        private static List<CountyInfo> AddCountyToCountyInfo(List<List<string>> countiesData)
        {
            List<CountyInfo> countiesList = new List<CountyInfo>();
            int length = 0;

            for (var i = 0; i < countiesData.Count; i++)
            {
                for (var j = 0; j < countiesData[i].Count; j++)
                {
                    var countyInfoString = countiesData[i][j];

                    MatchCollection matched = Regex.Matches(countyInfoString, @"(a?\d+)-\w\s([^-]+)");

                    CountyInfo countyInfo = new CountyInfo(matched[0].Groups[2].Value, matched[0].Groups[1].Value);
                    countiesList.Add(countyInfo);
                    length++;
                    Console.WriteLine(length + "/374" + countyInfo.Name + " " + countyInfo.ID);
                }
            }
            return countiesList;
        }

        private static List<List<string>> CountiesData()
        {
            List<List<string>> countiesData = new List<List<string>>();

            using (var httpClient = new HttpClient())
            {
                string getDataURL = "http://tianqi.2345.com/js/citySelectData.js";
                httpClient.BaseAddress = new Uri(getDataURL);

                var temp = httpClient.GetByteArrayAsync(getDataURL).Result;

                var tempCountiesList = Encoding.Default.GetString(temp);

                var countiesList = Regex.Replace(tempCountiesList, @"var\sprovqx=new[\s\S]*", "");

                var counties = countiesList.Split('\n').ToList();
                counties.RemoveAt(0);
                counties.RemoveAt(counties.Count - 1);
                counties.RemoveAt(counties.Count - 1);
                counties[0] = "54511-B 北京-54511";
                counties[1] = "54527-T 天津-54527";
                counties[2] = "58362-S 上海-58362";
                counties[3] = "57516-C 重庆-57516";
                counties[4] = "45007-X 香港-45007";
                counties[5] = "'45011-A 澳门-45011";
                counties[6] = "59554-G 高雄-59554|71301-F 屏东-59554|71298-J 嘉义-59554|71299-T 台南-59554|71300-T 台东-59554|71295-T 桃园-71294|71294-T 台北-71294|71296-X 新竹-71294|71297-Y 宜兰-71294|71305-H 花莲-71082|71302-M 苗栗-71082|71304-N 南投-71082|71306-Y 云林-71082|71303-Z 彰化-71082|71082-T 台中-71082";

                foreach (var county in counties)
                {
                    List<string> groupData = new List<string>();

                    if (!string.IsNullOrEmpty(county))
                    {
                        //var realCounty = Regex.Replace(county, @"(provqx\[\d+\]=\[)(.*)\][\s\S]", @"${2}");
                        //realCounty = Regex.Replace(realCounty, ",", "|");
                        var oneCounty = county.Split('|');

                        foreach (var item in oneCounty)
                        {
                            groupData.Add(item);
                        }
                        countiesData.Add(groupData);
                    }

                }
            }
            return countiesData;
        }

        private static List<WeatherInfo> ConverJObjectDataToWeatherInfo(JObject jObject)
        {
            //var jObject = (JObject)JsonConvert.DeserializeObject(result);
            List<WeatherInfo> weatherList = new List<WeatherInfo>();

            int weatherCount = jObject["tqInfo"].Count();

            for (int i = 0; i < weatherCount; i++)
            {
                var result = jObject["tqInfo"][i];

                if (result.First != null)
                {
                    WeatherInfo weatherInfo = new WeatherInfo(jObject["city"].ToString(), result["ymd"].ToString(), result["bWendu"].ToString(), result["yWendu"].ToString(), result["tianqi"].ToString(), result["fengxiang"].ToString());
                    weatherList.Add(weatherInfo);
                }
            }
            return weatherList;
        }

        public static bool WriteXls(string filename)
        {
            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            _Workbook book = xls.Workbooks.Add(Missing.Value); //创建一张表，一张表可以包含多个sheet

            //如果表已经存在，可以用下面的命令打开
            //_Workbook book = xls.Workbooks.Open(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            _Worksheet sheet;
            xls.Visible = false;//设置Excel后台运行
            xls.DisplayAlerts = false;//设置不显示确认修改提示

            for (int i = 1; i < 4; i++)//循环创建并写入数据到sheet
            {
                try
                {
                    sheet = (_Worksheet)book.Worksheets.get_Item(i);
                }
                catch (Exception ex)//不存在就增加一个sheet
                {
                    Console.WriteLine(ex.Message);
                    sheet = (_Worksheet)book.Worksheets.Add(Missing.Value, book.Worksheets[book.Sheets.Count], 1, Missing.Value);
                    Console.WriteLine("已添加sheet");
                }
                sheet.Name = "第" + i.ToString() + "页";
                for (int row = 1; row < 20; row++)
                {
                    for (int offset = 1; offset < 10; offset++)
                        sheet.Cells[row, offset] = "( " + row.ToString() + "," + offset.ToString() + " )";
                }
            }

            book.SaveAs(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //book.Save();

            book.Close(false, Missing.Value, Missing.Value);
            xls.Quit();
            sheet = null;
            book = null;
            xls = null;
            GC.Collect();
            return true;
        }

        public static Array ReadXls(string filename, int index)
        {

            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();

            _Workbook book = xls.Workbooks.Open(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            _Worksheet sheet;
            xls.Visible = false;
            xls.DisplayAlerts = false;

            try
            {
                sheet = (_Worksheet)book.Worksheets.get_Item(index);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

            Console.WriteLine(sheet.Name);
            int row = sheet.UsedRange.Rows.Count;
            int col = sheet.UsedRange.Columns.Count;
            var value = (Array)sheet.Range[sheet.Cells[1, 1], sheet.Cells[row, col]].Cells.Value2;//get_Range(sheet.Cells[1, 1], sheet.Cells[row, col]).Cells.Value2;

            book.Save();
            book.Close(false, Missing.Value, Missing.Value);
            xls.Quit();
            sheet = null;
            book = null;
            xls = null;
            GC.Collect();
            return value;
        }

        public static bool WriteXls(string filename, List<string> lat, List<string> lon, List<string> cityNames)
        {

            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            _Workbook book = xls.Workbooks.Add(Missing.Value); //创建一张表，一张表可以包含多个sheet

            //如果表已经存在，可以用下面的命令打开
            //_Workbook book = xls.Workbooks.Open(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            _Worksheet sheet;
            xls.Visible = false;//设置Excel后台运行
            xls.DisplayAlerts = false;//设置不显示确认修改提示

            //for (int i = 1; i < 2; i++)//循环创建并写入数据到sheet
            int row = 1;
            {
                try
                {
                    sheet = (_Worksheet)book.Worksheets.get_Item(1);
                }
                catch (Exception ex)//不存在就增加一个sheet
                {
                    Console.WriteLine(ex.Message);
                    sheet = (_Worksheet)book.Worksheets.Add(Missing.Value, book.Worksheets[book.Sheets.Count], 1, Missing.Value);
                    Console.WriteLine("已添加sheet");
                }
                //sheet.Name = "第" + 1 + "页";
                for (int i = 0; i < lat.Count; i++)
                {
                    sheet.Cells[i + 1, 1] = cityNames[i].ToString();
                    sheet.Cells[i + 1, 2] = lat[i].ToString();
                    sheet.Cells[i + 1, 3] = lon[i].ToString();

                }
            }

            book.SaveAs(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //book.Save();

            book.Close(false, Missing.Value, Missing.Value);
            xls.Quit();
            sheet = null;
            book = null;
            xls = null;
            GC.Collect();
            return true;
        }

    }
}
