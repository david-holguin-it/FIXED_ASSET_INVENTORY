using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet; 
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using System.Numerics;
using System.Text;

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
         //   SqlCommand cmd = new SqlCommand("SELECT * FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]", con);
            SqlCommand cmd = new SqlCommand(@"SELECT id,
                                                    manufacturerName,
                                                    partyManufacturerName,
                                                    materialNumber,
                                                    description,
                                                    purchaseValue,
                                                    accumulatedDepreciation,
                                                    netBookValue,
                                                    purchaseDate,
                                                    purchaseOrderNo,
                                                    department,
                                                    fixedAssetNumber,
                                                    serialNumber,
                                                    location,
                                                    PIC,
                                                    glAccount       
                                                    FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] ", con);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var record = new
                {
                    id                      = reader[0]                      == DBNull.Value ?"0" : reader[0].ToString(),
                    manufacturerName        = reader[1]        == DBNull.Value ? "" : reader[1].ToString(),
                    partyManufacturerName   = reader[2]   == DBNull.Value ? "" : reader[2].ToString(),  // TBD
                    materialNumber          = reader[3]          == DBNull.Value ? "" : reader[3].ToString(), 
                    description             = reader[4]             == DBNull.Value ? "" : reader[4].ToString(),
                    purchaseValue           = reader[5]           == DBNull.Value ? "" : reader[5].ToString(),
                    accumulatedDepreciation = reader[6] == DBNull.Value ? "" : reader[6].ToString(),
                    netBookValue            = reader[7]            == DBNull.Value ? "" : reader[7].ToString(),
                    purchaseDate            = reader[8]            == DBNull.Value ? "" : reader[8].ToString(),
                    purchaseOrderNo         = reader[9]         == DBNull.Value ? "" : reader[9].ToString(),
                    department              = reader[10]              == DBNull.Value ? "" : reader[10].ToString(),
                    fixedAssetNumber        = reader[11]        == DBNull.Value ? "" : reader[11].ToString(),
                    serialNumber            = reader[12]            == DBNull.Value ? "" : reader[12].ToString(),
                    location                = reader[13]                == DBNull.Value ? "" : reader[13].ToString(),
                    PIC                     = reader[14]                     == DBNull.Value ? "" : reader[14].ToString(),
                    glAccount               = reader[15]               == DBNull.Value ? "" : reader[15].ToString()
                }; 
                dr.Add(record);
            }
            return Json(new
            {
                data = dr
            }); 
        }

        public IActionResult DownloadReport1()
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Report");
                    using (var con = new SqlConnection(_connStr))
                    {    
                        // 2. Agregar encabezados
                        worksheet.Cell(1, 1).Value = "Manufacturer Name";
                        worksheet.Cell(1, 2).Value = "Third Party Manufacturer Name";
                        worksheet.Cell(1, 3).Value = "Material number";
                        worksheet.Cell(1, 4).Value = "Description";
                        worksheet.Cell(1, 5).Value = "Purchase Value (USD)";
                        worksheet.Cell(1, 6).Value = "Accumulated Depreciation";
                        worksheet.Cell(1, 7).Value = "Net Book Value";
                        worksheet.Cell(1, 8).Value = "Purchase Date";
                        worksheet.Cell(1, 9).Value = "Purchase Order";
                        worksheet.Cell(1, 10).Value = "Department";
                        worksheet.Cell(1, 11).Value = "Fixed Asset Number";
                        worksheet.Cell(1, 12).Value = "Serial Number";
                        worksheet.Cell(1, 13).Value = "Location";
                        worksheet.Cell(1, 14).Value = "PIC";
                        worksheet.Range("A1:N1").Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Range("A1:N1").Style.Font.Bold = true;
                        // 3. Agregar algunas filas de ejemplo

                        var dr = new List<object>(); 
                        con.Open();
                        SqlCommand cmd = new SqlCommand(@"
                                                    SELECT 
                                                        manufacturerName,
                                                        partyManufacturerName,
                                                        materialNumber,
                                                        description,
                                                        purchaseValue,
                                                        accumulatedDepreciation,
                                                        netBookValue,
                                                        purchaseDate,
                                                        purchaseOrderNo,
                                                        department,
                                                        fixedAssetNumber,
                                                        serialNumber,
                                                        location,
                                                        PIC 
                                                    FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]", con);
                        SqlDataReader reader = cmd.ExecuteReader();
                        int row = 2; // Empezar desde la segunda fila, ya que la primera es para los encabezados
                        while (reader.Read())
                        {
                            for (int col = 0; col < reader.FieldCount; col++)
                            {
                                if (reader.IsDBNull(col))
                                {
                                    worksheet.Cell(row, col + 1).Value = "";
                                    continue;
                                }

                                var value = reader.GetValue(col);
                                switch (value)
                                {
                                    case int i: worksheet.Cell(row, col + 1).Value = i; break;
                                    case decimal d: worksheet.Cell(row, col + 1).Value = d; break;
                                    case double db: worksheet.Cell(row, col + 1).Value = db; break;
                                    case DateTime dt: worksheet.Cell(row, col + 1).Value = dt; break;
                                    case bool b: worksheet.Cell(row, col + 1).Value = b; break;
                                    default: worksheet.Cell(row, col + 1).Value = value.ToString(); break;
                                }
                            } 
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
         

    // Método en ASP.NET Core
        public IActionResult DownloadReport2()
            { // es mas legible que el siguiente 
                var ms = new MemoryStream();
                byte[] content;
                using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "FixedAssets"
                    };
                    sheets.Append(sheet);

                    using var con = new SqlConnection(_connStr);
                    con.Open();
                    using var cmd = new SqlCommand(@"
                    SELECT 
                        manufacturerName,
                        partyManufacturerName,
                        materialNumber,
                        description,
                        purchaseValue,
                        accumulatedDepreciation,
                        netBookValue,
                        purchaseDate,
                        purchaseOrderNo,
                        department,
                        fixedAssetNumber,
                        serialNumber,
                        location,
                        PIC, 
                        glAccount
                    FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]", con);
                    using var reader = cmd.ExecuteReader();

                    // Encabezados
                    var headerRow = new Row();
                    for (int col = 0; col < reader.FieldCount; col++)
                        headerRow.Append(new Cell { CellValue = new CellValue(reader.GetName(col)), DataType = CellValues.String });
                    sheetData.AppendChild(headerRow);

                    // Filas de datos
                    while (reader.Read())
                    {
                        var dataRow = new Row();
                        for (int col = 0; col < reader.FieldCount; col++)
                        {
                            Cell cell = new Cell();

                            if (reader.IsDBNull(col))
                            {
                                cell.CellValue = new CellValue("");
                                cell.DataType = CellValues.String;
                            }
                            else
                            {
                                var value = reader.GetValue(col);
                                switch (value)
                                {
                                    case int i:
                                        cell.CellValue = new CellValue(i.ToString());
                                        cell.DataType = CellValues.Number;
                                        break;
                                    case decimal d:
                                        cell.CellValue = new CellValue(d.ToString());
                                        cell.DataType = CellValues.Number;
                                        break;
                                    case double db:
                                        cell.CellValue = new CellValue(db.ToString());
                                        cell.DataType = CellValues.Number;
                                        break;
                                    case DateTime dt:
                                        cell.CellValue = new CellValue(dt.ToOADate().ToString());
                                        cell.DataType = CellValues.Number; // Excel fecha como número
                                        cell.StyleIndex = 1; // opcional: usar estilo de fecha si defines uno
                                        break;
                                    default:
                                        cell.CellValue = new CellValue(value.ToString());
                                        cell.DataType = CellValues.String;
                                        break;
                                }
                            }

                            dataRow.Append(cell);
                        }
                        sheetData.AppendChild(dataRow);
                    }
                }
              
                // escribir datos...  
                content = ms.ToArray(); 
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FixedAssetReport.xlsx");
            }


        public IActionResult DownloadReport3()
    {  // codigo orrible
            var stopwatch = Stopwatch.StartNew();  // inicia el cronómetro
            TimeSpan elapsedMs  ;  // tiempo en milisegundos
            var ms = new MemoryStream(); // NO usar using, lo devuelve ASP.NET Core
            byte[] c;
            // Crear documento Excel
            using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {


                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                // Abrimos conexión SQL
                stopwatch.Start();
                using var con = new SqlConnection(_connStr);

                stopwatch.Stop();  // detén el cronómetro
                elapsedMs = stopwatch.Elapsed;  // tiempo en milisegundos
                Debug.WriteLine($"Crear sl conn BD Tiempo transcurrido: {elapsedMs} ms");
                stopwatch.Restart();

                con.Open();

                stopwatch.Stop();  // detén el cronómetro
                elapsedMs = stopwatch.Elapsed;  // tiempo en milisegundos
                Debug.WriteLine($"Abrir conexion BD Tiempo transcurrido: {elapsedMs} ms");
                stopwatch.Restart();


                using var cmd = new SqlCommand(@"
                SELECT 
                    manufacturerName,
                    partyManufacturerName,
                    materialNumber,
                    description,
                    purchaseValue,
                    accumulatedDepreciation,
                    netBookValue,
                    purchaseDate,
                    purchaseOrderNo,
                    department,
                    fixedAssetNumber,
                    serialNumber,
                    location,
                    PIC, 
                    glAccount
                FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]", con);
                using var reader = cmd.ExecuteReader();

                stopwatch.Stop();  // detén el cronómetro
                elapsedMs = stopwatch.Elapsed;  // tiempo en milisegundos
                Debug.WriteLine($"Leer de BD Tiempo transcurrido: {elapsedMs} ms");
                stopwatch.Restart();

                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());


                    // Encabezados
                    writer.WriteStartElement(new Row());
                    for (int col = 0; col < reader.FieldCount; col++)
                    {
                        writer.WriteElement(new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(reader.GetName(col))
                        });
                    }
                    writer.WriteEndElement(); // Row

                    // Filas de datos
                    while (reader.Read())
                    {
                        writer.WriteStartElement(new Row());
                        for (int col = 0; col < reader.FieldCount; col++)
                        {
                            Cell cell = new Cell();

                            if (reader.IsDBNull(col))
                            {
                                cell.CellValue = new CellValue("");
                                cell.DataType = CellValues.String;
                            }
                            else
                            {
                                var value = reader.GetValue(col);
                                switch (value)
                                {
                                    case int i:
                                    case decimal d:
                                    case double db:
                                        cell.CellValue = new CellValue(value.ToString());
                                        cell.DataType = CellValues.Number;
                                        break;
                                    case DateTime dt:
                                        cell.CellValue = new CellValue(dt.ToOADate().ToString());
                                        cell.DataType = CellValues.Number;
                                        break;
                                    default:
                                        cell.CellValue = new CellValue(value.ToString());
                                        cell.DataType = CellValues.String;
                                        break;
                                }
                            }

                            writer.WriteElement(cell);
                        }
                        writer.WriteEndElement(); // Row
                    }

                    writer.WriteEndElement(); // SheetData
                    writer.WriteEndElement(); // Worksheet
                }

                // Hojas del workbook
                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "FixedAssets"
                };
                sheets.Append(sheet);

                stopwatch.Stop();  // detén el cronómetro
                elapsedMs = stopwatch.Elapsed;  // tiempo en milisegundos
                Debug.WriteLine($"write to excel Tiempo transcurrido: {elapsedMs} ms");

                stopwatch.Restart();

            }
            stopwatch.Stop();  // detén el cronómetro
            elapsedMs = stopwatch.Elapsed;  // tiempo en milisegundos
            Debug.WriteLine($"save excel Tiempo transcurrido: {elapsedMs} ms");
            ms.Position = 0;
            c= ms.ToArray();
            return File(c, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FixedAssetReport.xlsx");
        }

        public IActionResult DownloadReport()
        {   // este guarda toda la todo el datatable en la primera fila y columna
            var dt = new System.Data.DataTable();

            using (var con = new SqlConnection(_connStr))
            using (var cmd = new SqlCommand(@"
            SELECT 
                manufacturerName,
                partyManufacturerName,
                materialNumber,
                description,
                purchaseValue,
                accumulatedDepreciation,
                netBookValue,
                purchaseDate,
                purchaseOrderNo,
                department,
                fixedAssetNumber,
                serialNumber,
                location,
                PIC, 
                glAccount
            FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]", con))
            using (var adapter = new SqlDataAdapter(cmd))
            {
                adapter.Fill(dt);
            }

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("FixedAssets");

            // Inserta toda la DataTable empezando en A1
            worksheet.Cell(1, 1).InsertTable(dt, "FixedAssets", true); // true = encabezados

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);

            ms.Position = 0;
            return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FixedAssetReport.xlsx");
        }


    }
}
