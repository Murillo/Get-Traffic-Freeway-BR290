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
				@"/home/murillo/anaconda3/git/Get-Traffic-Freeway-BR290/data/arq.xlsx", 
				new DateTime(2016,01,01), 
				new DateTime(2016,01,03));
			traffic.Run(new Excel());
#else
            string path = args[0];
            DateTime start = Convert.ToDateTime(args[1]);
            DateTime end = Convert.ToDateTime(args[2]);
            Traffic traffic = new Traffic(path, start, end);
            traffic.Run(new Excel());
#endif
        }
    }
}
