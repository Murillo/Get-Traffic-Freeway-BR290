using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetTrafficFreewayBR290.Highway
{
	public class Flow
    {
        public int FluxoAcum12 { get; set; }
        public int FluxoAcum18 { get; set; }
        public int FluxoAcum24 { get; set; }
        public List<FlowByHour> ListFluxoHora { get; set; }
        public DateTime Data { get; set; }
    }
}
