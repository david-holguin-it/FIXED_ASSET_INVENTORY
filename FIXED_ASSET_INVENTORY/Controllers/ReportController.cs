using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Numerics;

namespace FIXED_ASSET_INVENTORY.Controllers
{
    [Authorize]
    public class ReportController : Controller
    {
        private readonly string _connStr;
        public ReportController(IConfiguration configuration)
        {
            if(configuration.GetConnectionString("PSGDbConnStr") != null)
                _connStr = configuration.GetConnectionString("PSGDbConnStr"); // Connection string from appsettings.json
        }

        public IActionResult Index()
        {
            var username = User.Identity.Name; // DOMAIN\usuario 
            ViewData["username"] = username;
            return View(); 
        }
        public JsonResult getReportData()
        {
            var dr = new List<object>();
            SqlConnection con = new SqlConnection(_connStr);
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT * FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]", con);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var record = new
                {
                    id                      = reader["id"]                      == DBNull.Value ?"0" : reader["id"].ToString(),
                    manufacturerName        = reader["manufacturerName"]        == DBNull.Value ? "" : reader["manufacturerName"].ToString(),
                    partyManufacturerName   = reader["partyManufacturerName"]   == DBNull.Value ? "" : reader["partyManufacturerName"].ToString(),  // TBD
                    materialNumber          = reader["materialNumber"]          == DBNull.Value ? "" : reader["materialNumber"].ToString(), 
                    description             = reader["description"]             == DBNull.Value ? "" : reader["description"].ToString(),
                    purchaseValue           = reader["purchaseValue"]           == DBNull.Value ? "" : reader["purchaseValue"].ToString(),
                    accumulatedDepreciation = reader["accumulatedDepreciation"] == DBNull.Value ? "" : reader["accumulatedDepreciation"].ToString(),
                    netBookValue            = reader["netBookValue"]            == DBNull.Value ? "" : reader["netBookValue"].ToString(),
                    purchaseOrderNo         = reader["purchaseOrderNo"]         == DBNull.Value ? "" : reader["purchaseOrderNo"].ToString(),
                    department              = reader["department"]              == DBNull.Value ? "" : reader["department"].ToString(),
                    fixedAssetNumber        = reader["fixedAssetNumber"]        == DBNull.Value ? "" : reader["fixedAssetNumber"].ToString(),
                    serialNumber            = reader["serialNumber"]            == DBNull.Value ? "" : reader["serialNumber"].ToString(),
                    location                = reader["location"]                == DBNull.Value ? "" : reader["location"].ToString(),
                    PIC                     = reader["PIC"]                     == DBNull.Value ? "" : reader["PIC"].ToString(),
                    glAccount               = reader["glAccount"]               == DBNull.Value ? "" : reader["glAccount"].ToString()
                }; 
                dr.Add(record);
            }
            return Json(new
            {
                data = dr
            }); 
        }

        public IActionResult DownloadReport()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Report"); 
                // 2. Agregar encabezados
                worksheet.Cell(1, 1).Value = "Manufacturer Name";
                worksheet.Cell(1, 2).Value = "Third Party Manufacturer Name";
                worksheet.Cell(1, 3).Value = "Material number";
                worksheet.Cell(1, 4).Value = "Description";
                worksheet.Cell(1, 5).Value = "Purchase Value (USD)";
                worksheet.Cell(1, 5).Value = "Accumulated Depreciation";
                worksheet.Cell(1, 12).Value = "Net Book Value";
                worksheet.Cell(1, 6).Value = "Purchase Order";
                worksheet.Cell(1, 7).Value = "Department";
                worksheet.Cell(1, 8).Value = "Fixed Asset Number";
                worksheet.Cell(1, 9).Value = "Serial Number";
                worksheet.Cell(1,10).Value = "Location";
                worksheet.Cell(1,11).Value = "PIC";
                worksheet.Range("A1:L1").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("A1:L1").Style.Font.Bold = true;
                // 3. Agregar algunas filas de ejemplo

                var dr = new List<object>();
                SqlConnection con = new SqlConnection(_connStr);
                con.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]", con);
                SqlDataReader reader = cmd.ExecuteReader();
                int row = 2; // Empezar desde la segunda fila, ya que la primera es para los encabezados
                while (reader.Read())
                {
                    //  worksheet.Cell(row, 1).Value  = reader["id"] == DBNull.Value ? "0" : reader["id"].ToString(),
                    worksheet.Cell(row, 1).Value = reader["manufacturerName"]       == DBNull.Value ? "" : reader["manufacturerName"].ToString(); 
                    worksheet.Cell(row, 2).Value = reader["partyManufacturerName"]  == DBNull.Value ? "" : reader["partyManufacturerName"].ToString();   // TBD
                    worksheet.Cell(row, 3).Value = reader["materialNumber"]         == DBNull.Value ? "" : reader["materialNumber"].ToString();
                    worksheet.Cell(row, 4).Value = reader["description"]            == DBNull.Value ? "" : reader["description"].ToString();
                    worksheet.Cell(row, 5).Value = reader["purchaseValue"]          == DBNull.Value ? "" : reader["purchaseValue"].ToString();
                    worksheet.Cell(row, 5).Value = reader["accumulatedDepreciation"]== DBNull.Value ? "" : reader["accumulatedDepreciation"].ToString();
                    worksheet.Cell(row, 12).Value= reader["netBookValue"]           == DBNull.Value ? "" : reader["netBookValue"].ToString();
                    worksheet.Cell(row, 6).Value = reader["purchaseOrderNo"]        == DBNull.Value ? "" : reader["purchaseOrderNo"].ToString();
                    worksheet.Cell(row, 7).Value = reader["department"]             == DBNull.Value ? "" : reader["department"].ToString();
                    worksheet.Cell(row, 8).Value = reader["fixedAssetNumber"]       == DBNull.Value ? "" : reader["fixedAssetNumber"].ToString();
                    worksheet.Cell(row, 9).Value = reader["serialNumber"]           == DBNull.Value ? "" : reader["serialNumber"].ToString();
                    worksheet.Cell(row,10).Value = reader["location"]               == DBNull.Value ? "" : reader["location"].ToString();
                    worksheet.Cell(row,11).Value = reader["PIC"]                    == DBNull.Value ? "" : reader["PIC"].ToString();
                    row++;
                } 

                // 4. Guardar el Excel en un MemoryStream
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    // 5. Retornar el archivo para descarga
                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Fixed Asset Report.xlsx"
                    );
                }
            }
        }
    }
}
