using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetWeatherData
{
    public class WeatherInfo
    {
        public string CountyName { get; set; }
        public string Date { get; set; }
        public string MaxTemperature { get; set; }
        public string MinTemperature { get; set; }
        public string Weather { get; set; }
        public string WindDirection { get; set; }

        public WeatherInfo() { }

        public WeatherInfo(string countyName,string date,string maxTemperature,string minTemperature,string weather,string windDirection)
        {
            this.CountyName = countyName;
            this.Date = date;
            this.MaxTemperature = maxTemperature;
            this.MinTemperature = minTemperature;
            this.Weather = weather;
            this.WindDirection = windDirection;
        }                 
    }
}
