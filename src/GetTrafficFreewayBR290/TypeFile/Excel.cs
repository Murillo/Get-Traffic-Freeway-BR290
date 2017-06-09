using GetTrafficFreewayBR290.Highway;
using GetTrafficFreewayBR290.Text;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace GetTrafficFreewayBR290.TypeFile
{
	internal class Excel : ITypeFile
    {
		private void GenerateHeader(ExcelWorksheet worksheet)
		{
			if (worksheet != null)
			{
				worksheet.Cells[1, 1].Value     = Header.Date;
				worksheet.Cells[1, 2].Value     = Header.FlowAccum12;
				worksheet.Cells[1, 3].Value     = Header.FlowAccum18;
				worksheet.Cells[1, 4].Value     = Header.FlowAccum24;
				worksheet.Cells[1, 5].Value     = Header.AverageFlowHour01;
				worksheet.Cells[1, 6].Value     = Header.AverageFlowHour02;
				worksheet.Cells[1, 7].Value     = Header.AverageFlowHour03;
				worksheet.Cells[1, 8].Value     = Header.AverageFlowHour04;
				worksheet.Cells[1, 9].Value     = Header.AverageFlowHour05;
				worksheet.Cells[1, 10].Value    = Header.AverageFlowHour06;
				worksheet.Cells[1, 11].Value    = Header.AverageFlowHour07;
				worksheet.Cells[1, 12].Value    = Header.AverageFlowHour08;
				worksheet.Cells[1, 13].Value    = Header.AverageFlowHour09;
				worksheet.Cells[1, 14].Value    = Header.AverageFlowHour10;
				worksheet.Cells[1, 15].Value    = Header.AverageFlowHour11;
				worksheet.Cells[1, 16].Value    = Header.AverageFlowHour12;
				worksheet.Cells[1, 17].Value    = Header.AverageFlowHour13;
				worksheet.Cells[1, 18].Value    = Header.AverageFlowHour14;
				worksheet.Cells[1, 19].Value    = Header.AverageFlowHour15;
				worksheet.Cells[1, 20].Value    = Header.AverageFlowHour16;
				worksheet.Cells[1, 21].Value    = Header.AverageFlowHour17;
				worksheet.Cells[1, 22].Value    = Header.AverageFlowHour18;
				worksheet.Cells[1, 23].Value    = Header.AverageFlowHour19;
				worksheet.Cells[1, 24].Value    = Header.AverageFlowHour20;
				worksheet.Cells[1, 25].Value    = Header.AverageFlowHour21;
				worksheet.Cells[1, 26].Value    = Header.AverageFlowHour22;
				worksheet.Cells[1, 27].Value    = Header.AverageFlowHour23;
				worksheet.Cells[1, 28].Value    = Header.AverageFlowHour24;
				worksheet.Cells[1, 29].Value    = Header.TotalFlowHour01;
				worksheet.Cells[1, 30].Value    = Header.TotalFlowHour02;
				worksheet.Cells[1, 31].Value    = Header.TotalFlowHour03;
				worksheet.Cells[1, 32].Value    = Header.TotalFlowHour04;
				worksheet.Cells[1, 33].Value    = Header.TotalFlowHour05;
				worksheet.Cells[1, 34].Value    = Header.TotalFlowHour06;
				worksheet.Cells[1, 35].Value    = Header.TotalFlowHour07;
				worksheet.Cells[1, 36].Value    = Header.TotalFlowHour08;
				worksheet.Cells[1, 37].Value    = Header.TotalFlowHour09;
				worksheet.Cells[1, 38].Value    = Header.TotalFlowHour10;
				worksheet.Cells[1, 39].Value    = Header.TotalFlowHour11;
				worksheet.Cells[1, 40].Value    = Header.TotalFlowHour12;
				worksheet.Cells[1, 41].Value    = Header.TotalFlowHour13;
				worksheet.Cells[1, 42].Value    = Header.TotalFlowHour14;
				worksheet.Cells[1, 43].Value    = Header.TotalFlowHour15;
				worksheet.Cells[1, 44].Value    = Header.TotalFlowHour16;
				worksheet.Cells[1, 45].Value    = Header.TotalFlowHour17;
				worksheet.Cells[1, 46].Value    = Header.TotalFlowHour18;
				worksheet.Cells[1, 47].Value    = Header.TotalFlowHour19;
				worksheet.Cells[1, 48].Value    = Header.TotalFlowHour20;
				worksheet.Cells[1, 49].Value    = Header.TotalFlowHour21;
				worksheet.Cells[1, 50].Value    = Header.TotalFlowHour22;
				worksheet.Cells[1, 51].Value    = Header.TotalFlowHour23;
				worksheet.Cells[1, 52].Value    = Header.TotalFlowHour24;
				worksheet.Cells[1, 1, 1, 52].Style.Font.Bold = true;
				worksheet.Cells[1, 1, 1, 52].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
			}
		}

		private void GenerateFile(ExcelWorksheet worksheet, List<Flow> listFlow)
		{
			int x = 2;
			foreach (Flow item in listFlow)
			{
				worksheet.Cells[x, 1].Value = item.Date.ToString("dd/MM/yyyy");
				worksheet.Cells[x, 2].Value = item.FlowAccum12;
				worksheet.Cells[x, 3].Value = item.FlowAccum18;
				worksheet.Cells[x, 4].Value = item.FlowAccum24;

				int y = 5;
				foreach (var itemListaFluxo in item.ListFlowHour)
				{
					worksheet.Cells[x, y].Value = itemListaFluxo.AverageFlow;
					y++;
				}

				int z = 29;
				foreach (var itemListaFluxo in item.ListFlowHour)
				{
					worksheet.Cells[x, z].Value = itemListaFluxo.Flow;
					z++;
				}
				x++;
			}
		}

		public byte[] GetFile(List<Flow> flow)
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(Header.Report);
				this.GenerateHeader(worksheet);
				this.GenerateFile(worksheet, flow);
				worksheet.Cells.AutoFitColumns();
				MemoryStream stream = new MemoryStream();
				package.SaveAs(stream);
				return stream.GetBuffer();
			}
		}
    }
}
