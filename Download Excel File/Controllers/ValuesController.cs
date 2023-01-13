using ClosedXML.Excel;
using Dapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;

namespace Download_Excel_File.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class ValuesController : ControllerBase
	{
		private readonly IConfiguration config;

		public ValuesController(IConfiguration config)
		{
			this.config = config;
		}

		[HttpGet]
		public async Task<IActionResult> Get()
		{
			var connection = new SqlConnection(config.GetConnectionString("defaultconnection"));
			var data = await connection.QueryAsync<Model>("select * from closexmltbl");
			
			XLWorkbook workbook = new XLWorkbook();
			IXLWorksheet worksheet = workbook.Worksheets.Add();

			
			var currentrow = 1;

			//Excel File ma Header Set Gareko
			worksheet.Cell(currentrow, 1).Value = "Customercode";
			worksheet.Cell(currentrow, 2).Value = "FirstName";
			worksheet.Cell(currentrow, 3).Value = "LastName";
			worksheet.Cell(currentrow, 4).Value = "gender";
			worksheet.Cell(currentrow, 5).Value = "Country";
			worksheet.Cell(currentrow, 6).Value = "Age";
			
			//Excel File ma Value Set Gareko
			foreach(Model datum in data)
			{
				currentrow++;
				//worksheet.Cell(currentrow, 1).Value = datum.Id;
				worksheet.Cell(currentrow, 1).Value = datum.customercode;
				worksheet.Cell(currentrow, 2).Value = datum.firstname;
				worksheet.Cell(currentrow, 3).Value = datum.lastname;
				worksheet.Cell(currentrow, 4).Value = datum.gender;
				worksheet.Cell(currentrow, 5).Value = datum.country;
				worksheet.Cell(currentrow, 6).Value = datum.age;

			}
			
			//Workbook lai stream ma save gareko ani FileStream 
			MemoryStream stream = new MemoryStream();

			workbook.SaveAs(stream);
			stream.Position = 0;

			return new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") { FileDownloadName = "ssb.xlsx" };


			//Tala ko code chai excel file create garera download garne tarika
			
			//XLWorkbook workbook= new XLWorkbook();
			//IXLWorksheet worksheet = workbook.Worksheets.Add(1);
			//worksheet.Cell(1, 1).SetValue("Hello");

			//MemoryStream stream= new MemoryStream();

			//workbook.SaveAs(stream);
			//stream.Position = 0;

			//return new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") { FileDownloadName = "ssb.xlsx" };

		}

	}
}
