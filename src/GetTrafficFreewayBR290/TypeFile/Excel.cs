using GetTrafficFreewayBR290.Highway;
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
				worksheet.Cells[1, 1].Value     = "Data";
				worksheet.Cells[1, 2].Value     = "Fluxo Acumulado 12";
				worksheet.Cells[1, 3].Value     = "Fluxo Acumulado 18";
				worksheet.Cells[1, 4].Value     = "Fluxo Acumulado 24";
				worksheet.Cells[1, 5].Value     = "Hora Fluxo Médio: 01";
				worksheet.Cells[1, 6].Value     = "Hora Fluxo Médio: 02";
				worksheet.Cells[1, 7].Value     = "Hora Fluxo Médio: 03";
				worksheet.Cells[1, 8].Value     = "Hora Fluxo Médio: 04";
				worksheet.Cells[1, 9].Value     = "Hora Fluxo Médio: 05";
				worksheet.Cells[1, 10].Value    = "Hora Fluxo Médio: 06";
				worksheet.Cells[1, 11].Value    = "Hora Fluxo Médio: 07";
				worksheet.Cells[1, 12].Value    = "Hora Fluxo Médio: 08";
				worksheet.Cells[1, 13].Value    = "Hora Fluxo Médio: 09";
				worksheet.Cells[1, 14].Value    = "Hora Fluxo Médio: 10";
				worksheet.Cells[1, 15].Value    = "Hora Fluxo Médio: 11";
				worksheet.Cells[1, 16].Value    = "Hora Fluxo Médio: 12";
				worksheet.Cells[1, 17].Value    = "Hora Fluxo Médio: 13";
				worksheet.Cells[1, 18].Value    = "Hora Fluxo Médio: 14";
				worksheet.Cells[1, 19].Value    = "Hora Fluxo Médio: 15";
				worksheet.Cells[1, 20].Value    = "Hora Fluxo Médio: 16";
				worksheet.Cells[1, 21].Value    = "Hora Fluxo Médio: 17";
				worksheet.Cells[1, 22].Value    = "Hora Fluxo Médio: 18";
				worksheet.Cells[1, 23].Value    = "Hora Fluxo Médio: 19";
				worksheet.Cells[1, 24].Value    = "Hora Fluxo Médio: 20";
				worksheet.Cells[1, 25].Value    = "Hora Fluxo Médio: 21";
				worksheet.Cells[1, 26].Value    = "Hora Fluxo Médio: 22";
				worksheet.Cells[1, 27].Value    = "Hora Fluxo Médio: 23";
				worksheet.Cells[1, 28].Value    = "Hora Fluxo Médio: 24";
				worksheet.Cells[1, 29].Value    = "Hora Fluxo Total: 01";
				worksheet.Cells[1, 30].Value    = "Hora Fluxo Total: 02";
				worksheet.Cells[1, 31].Value    = "Hora Fluxo Total: 03";
				worksheet.Cells[1, 32].Value    = "Hora Fluxo Total: 04";
				worksheet.Cells[1, 33].Value    = "Hora Fluxo Total: 05";
				worksheet.Cells[1, 34].Value    = "Hora Fluxo Total: 06";
				worksheet.Cells[1, 35].Value    = "Hora Fluxo Total: 07";
				worksheet.Cells[1, 36].Value    = "Hora Fluxo Total: 08";
				worksheet.Cells[1, 37].Value    = "Hora Fluxo Total: 09";
				worksheet.Cells[1, 38].Value    = "Hora Fluxo Total: 10";
				worksheet.Cells[1, 39].Value    = "Hora Fluxo Total: 11";
				worksheet.Cells[1, 40].Value    = "Hora Fluxo Total: 12";
				worksheet.Cells[1, 41].Value    = "Hora Fluxo Total: 13";
				worksheet.Cells[1, 42].Value    = "Hora Fluxo Total: 14";
				worksheet.Cells[1, 43].Value    = "Hora Fluxo Total: 15";
				worksheet.Cells[1, 44].Value    = "Hora Fluxo Total: 16";
				worksheet.Cells[1, 45].Value    = "Hora Fluxo Total: 17";
				worksheet.Cells[1, 46].Value    = "Hora Fluxo Total: 18";
				worksheet.Cells[1, 47].Value    = "Hora Fluxo Total: 19";
				worksheet.Cells[1, 48].Value    = "Hora Fluxo Total: 20";
				worksheet.Cells[1, 49].Value    = "Hora Fluxo Total: 21";
				worksheet.Cells[1, 50].Value    = "Hora Fluxo Total: 22";
				worksheet.Cells[1, 51].Value    = "Hora Fluxo Total: 23";
				worksheet.Cells[1, 52].Value    = "Hora Fluxo Total: 24";
				worksheet.Cells[1, 1, 1, 52].Style.Font.Bold = true;
				worksheet.Cells[1, 1, 1, 52].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
			}
		}

		private void GenerateFile(ExcelWorksheet worksheet, List<Flow> listFlow)
		{
			int x = 2;
			foreach (Flow item in listFlow)
			{
				worksheet.Cells[x, 1].Value = item.Data.ToString("dd/MM/yyyy");
				worksheet.Cells[x, 2].Value = item.FluxoAcum12;
				worksheet.Cells[x, 3].Value = item.FluxoAcum18;
				worksheet.Cells[x, 4].Value = item.FluxoAcum24;

				int y = 5;
				foreach (var itemListaFluxo in item.ListFluxoHora)
				{
					worksheet.Cells[x, y].Value = itemListaFluxo.FluxoMedio;
					y++;
				}

				int z = 29;
				foreach (var itemListaFluxo in item.ListFluxoHora)
				{
					worksheet.Cells[x, z].Value = itemListaFluxo.Fluxo;
					z++;
				}
				x++;
			}
		}

		public byte[] GetFile(List<Flow> flow)
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Relatório");
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
