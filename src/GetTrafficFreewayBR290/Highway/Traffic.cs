using GetTrafficFreewayBR290.Text;
using GetTrafficFreewayBR290.TypeFile;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace GetTrafficFreewayBR290.Highway
{
    /// <summary>
    /// This class represents the object that will fetch the traffic data on the 
    /// highway BR290, located in the south of Brazil, and export in the desired format. 
    /// </summary>
    public class Traffic
    {
        private readonly List<Flow> _listFlow = new List<Flow>();
        private readonly string urlXmlDataTotal = "http://189.45.47.195/fluxorodovia/XMLData/fluxoacumulado.aspx?dtIni={0}&cdLId={1}";
        private readonly string urlXmlHoraTotal = "http://189.45.47.195/fluxorodovia/XMLData/fluxomedio.aspx?dtIni={0}&cdLId={1}";
        private readonly string _file;
        private readonly DateTime _firstPeriod;
        private readonly DateTime _finalPeriod;

        /// <summary>
        /// Load the Traffic object by passing the following parameters
        /// </summary>
        /// <param name="file">Name of the file with its path on the machine</param>
        /// <param name="firstPeriod">Initial period of traffic search</param>
        /// <param name="finalPeriod">Final traffic search period</param>
        public Traffic(string file, DateTime firstPeriod, DateTime finalPeriod)
        {
            _file = file;
            _firstPeriod = firstPeriod;
            _finalPeriod = finalPeriod;
        }

        /// <summary>
        /// Starts the process of obtaining the data
        /// </summary>
		public void Run(ITypeFile file)
        {
            ExtractData();
            CreateFile(file);
        }

        /// <summary>
        /// This method extract the data of traffic
        /// </summary>
        private void ExtractData()
        {
            string urlDayTotal = string.Empty;
            string urlDayHour = string.Empty;
            for (DateTime date = _firstPeriod; date < _finalPeriod; date = date.AddDays(1.0))
            {
                urlDayTotal = string.Format(
                    urlXmlDataTotal,
                    string.Concat(date.Year, date.Month.ToString().PadLeft(2, '0'), date.Day.ToString().PadLeft(2, '0')),
                    "12");
                XDocument xmlDadosDiarioTotal = XDocument.Load(urlDayTotal);
                Flow totalFlow = (from total in xmlDadosDiarioTotal.Descendants(ItensData.FlowData)
                                  select new Flow
                                  {
                                      FlowAccum12 = Convert.ToInt32(total.Descendants(ItensData.Accumulated).FirstOrDefault().Attribute(ItensData.FlowAccum12).Value),
                                      FlowAccum18 = Convert.ToInt32(total.Descendants(ItensData.Accumulated).FirstOrDefault().Attribute(ItensData.FlowAccum18).Value),
                                      FlowAccum24 = Convert.ToInt32(total.Descendants(ItensData.Accumulated).FirstOrDefault().Attribute(ItensData.FlowAccum24).Value)
                                  }).FirstOrDefault();
                totalFlow.Date = date;


                urlDayHour = string.Format(
                    urlXmlHoraTotal,
                    string.Concat(date.Year, date.Month.ToString().PadLeft(2, '0'), date.Day.ToString().PadLeft(2, '0')),
                    "12");
                XDocument xmlDataDayHour = XDocument.Load(urlDayHour);
                List<FlowByHour> listFlowHourDate = new List<FlowByHour>();
                foreach (var item in xmlDataDayHour.Descendants(ItensData.FlowData).ToList())
                {
                    foreach (var subItem in item.Descendants(ItensData.Time).ToList())
                    {
                        FlowByHour flowHour = new FlowByHour();
                        flowHour.Flow = Convert.ToInt32(subItem.Attribute(ItensData.Flow).Value);
                        flowHour.AverageFlow = Convert.ToInt32(subItem.Attribute(ItensData.AverageFlow).Value);
                        flowHour.Hour = Convert.ToInt32(subItem.Attribute(ItensData.Hour).Value);
                        listFlowHourDate.Add(flowHour);
                    }
                }
                totalFlow.ListFlowHour = listFlowHourDate;
                _listFlow.Add(totalFlow);
                date.AddDays(1);
            }
        }

        /// <summary>
        /// This method create the file that has the informations
        /// </summary>
        /// <param name="file"></param>
        private void CreateFile(ITypeFile file)
        {
            var dataset = file.GetFile(_listFlow);
            BinaryWriter bw = new BinaryWriter(File.Open(_file, FileMode.OpenOrCreate));
            bw.Write(dataset);
        }
    }
}
