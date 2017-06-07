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
            string urlDayTotal = string.Empty;
            string urlDayHour = string.Empty;
            for (DateTime data = _firstPeriod; data < _finalPeriod; data = data.AddDays(1.0))
            {
                urlDayTotal = String.Format(
                    urlXmlDataTotal,
                    String.Concat(data.Year, data.Month.ToString().PadLeft(2, '0'), data.Day.ToString().PadLeft(2, '0')),
                    "12");
                XDocument xmlDadosDiarioTotal = XDocument.Load(urlDayTotal);
                Flow fluxoTotal = (from total in xmlDadosDiarioTotal.Descendants("FluxoData")
                                    select new Flow
                                    {
                                        FluxoAcum12 = Convert.ToInt32(total.Descendants("Acumulado").FirstOrDefault().Attribute("FluxoAcum12").Value),
                                        FluxoAcum18 = Convert.ToInt32(total.Descendants("Acumulado").FirstOrDefault().Attribute("FluxoAcum18").Value),
                                        FluxoAcum24 = Convert.ToInt32(total.Descendants("Acumulado").FirstOrDefault().Attribute("FluxoAcum24").Value)
                                    }).FirstOrDefault();
                fluxoTotal.Data = data;
	

                urlDayHour = String.Format(
                    urlXmlHoraTotal,
                    String.Concat(data.Year, data.Month.ToString().PadLeft(2, '0'), data.Day.ToString().PadLeft(2, '0')),
                    "12");
                XDocument xmlDadosDiarioHora = XDocument.Load(urlDayHour);

                List<FlowByHour> listFluxoHorasDados = new List<FlowByHour>();
                foreach (var item in xmlDadosDiarioHora.Descendants("FluxoData").ToList())
                {
                    foreach (var itemFilho in item.Descendants("Time").ToList())
                    {
                        FlowByHour fluxoHora = new FlowByHour();
                        fluxoHora.Fluxo = Convert.ToInt32(itemFilho.Attribute("fluxo").Value);
                        fluxoHora.FluxoMedio = Convert.ToInt32(itemFilho.Attribute("fluxomedio").Value);
                        fluxoHora.Hora = Convert.ToInt32(itemFilho.Attribute("hora").Value);
                        listFluxoHorasDados.Add(fluxoHora);
                    }
                }
                fluxoTotal.ListFluxoHora = listFluxoHorasDados;
                _listFlow.Add(fluxoTotal);
                data.AddDays(1);
            }
			var dataset = file.GetFile(_listFlow);
			BinaryWriter bw = new BinaryWriter(File.Open(_file, FileMode.OpenOrCreate));
			bw.Write(dataset);
        }
    }
}
