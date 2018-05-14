using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetWeatherData
{
    public class CountyInfo
    {
        public string Name { get; set; }
        public string ID { get; set; }

        public CountyInfo() { }
        public CountyInfo(string countyName,string countyID)
        {
            this.Name = countyName;
            this.ID = countyID;
        }
    }
}
