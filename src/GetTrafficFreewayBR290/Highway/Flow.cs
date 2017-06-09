using System;
using System.Collections.Generic;

namespace GetTrafficFreewayBR290.Highway
{
    public class Flow
    {
        public int FlowAccum12 { get; set; }
        public int FlowAccum18 { get; set; }
        public int FlowAccum24 { get; set; }
        public List<FlowByHour> ListFlowHour { get; set; }
        public DateTime Date { get; set; }
    }
}
