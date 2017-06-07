using GetTrafficFreewayBR290.Highway;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GetTrafficFreewayBR290.TypeFile;

namespace GetTrafficFreewayBR290
{
    class Program
    {
        static void Main(string[] args)
        {
#if DEBUG
			Traffic traffic = new Traffic(
				@"/home/murillo/Desktop/GetTrafficFreewayBR290/arq.xlsx", 
				new DateTime(2016,05,01), 
				new DateTime(2016,05,03));
			traffic.Run(new Excel());
#else
            Traffic traffic = new Traffic();
            traffic.Run();
#endif
        }
    }
}
