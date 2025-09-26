using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using FIXED_ASSET_INVENTORY.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Net;
using System.Reflection.PortableExecutable;
using System.Text.Json;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace FIXED_ASSET_INVENTORY.Controllers
{

    /// <summary>
    /// Los fixed assets se son bienes que se usan para producir bienes y servicios, y no se destinan para venta 
    /// TODOS 
    ///     Calcular depreciacion mensualmente antes del cierre del mes (TBP by RENANDO)
    ///     1    tomar la fecha de capitalizacion para empezar a depreciar
    ///     2    despues de la vida util se haya vencido, el valor en libros debe ser 0
    ///     3    llevar un registro de la depreciacion acumulada ?
    ///     4    llevar un registro del valor en libros (net book value) ??
    ///     -Cargar campos faltantes del formato
    ///     -Agregar  campos faltantes a
    ///     -Cambiar  locacion y llevar el registro de quien lo registro, cuando y a donde 
    ///     -Restricted acces by AD
    /// CONCEPTS
    /// Capitalization Day es la fecha en que un activo fijo se registra en los libros contables de la empresa, esto es a partir de que ya se puede usar y  a partir de esta fecha se empieza a depreciar el activo
    /// </summary>
    [Authorize]
    public class FAIController : Controller
    { 
        private readonly string _connStr;
        List<string> lstCols = new List<string>() {
            "manufacturerName",
            "partyManufacturerName",
            "materialNumber",
            "productName",
            "description", 
            "purchaseValue", 
            "purchaseValueUSD", 
            "paymentTerms",
            "purchaseOrderNo",
            "contractNo",
            "signOff",
            "remark",
            "materialsSent",
            "department",
            "manager",
            "fixedAssetNumber",
            "serialNumber",
            "location",
            "PIC",
            "createdBy"
        };
        public FAIController(IConfiguration configuration)
        {
            _connStr=configuration.GetConnectionString("PSGDbConnStr");
        }

        public IActionResult Index()
        {
            var username = User.Identity.Name; // DOMAIN\usuario 
            ViewData["username"] = username;
            return View();
        } 

        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file, string page, string userName)
        {
            if (file == null  || file.Length == 0)
            {
                ViewBag.Message = "No se seleccionó ningún archivo."; 
                return View("Index");
            }
            // Carpeta donde se guardarán los archivos
            //  var path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/uploads", file.FileName);

            var path = Directory.GetCurrentDirectory() + file.FileName ;
            using (var stream = new FileStream(path, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }
            ClosedXML.Excel.XLWorkbook workbook = new ClosedXML.Excel.XLWorkbook(path);
            var worksheet = workbook.Worksheet(1); // Selecciona la primera hoja
            if(page == "2")
            { 
                worksheet = workbook.Worksheet(2); // Selecciona la primera hoja
            }
                //   var rows = worksheet.RowsUsed().Skip(1); // Omite la fila de encabezado

            List<string> readedHeaderNames = new List<string>(); 
             
            var headersInFirstRow = worksheet.Rows()
                 .Where(row => !row.IsEmpty())
                 .FirstOrDefault();
            foreach (var cell in headersInFirstRow.Cells())
            {
                readedHeaderNames.Add(cell.GetValue<string>());
                Debug.WriteLine(cell.GetValue<string>());
            }
            var rows = worksheet.Rows()
                        .Where(row => !row.IsEmpty())
                        .Skip(2);
            SqlConnection c = new SqlConnection(_connStr);
            c.Open();
            int cnt = 0;
            List<string> lstErrors = new List<string>();
            using (SqlTransaction tran = c.BeginTransaction())
            {
                lstCols.Add("");
                foreach (var row in rows)
                {
                    var cells = row.Cells(1, 29).Select(c => c.Value).ToArray();
                    //   if (cnt >= 10) break;
                    if (row.IsEmpty()) continue;

                    string insertQuery = string.Format("INSERT INTO [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] "
                       + "         ( {0} )"
                       + " VALUES  ( {1} )", string.Join(",", lstCols), "@" + string.Join(",@", lstCols)); 
                /* 
                    category, netBookValue, usefulLife, capitalizationDate, accumulatedDepreciation
                */
                    SqlCommand cmd = new SqlCommand(insertQuery, c, tran);     
      

                    try
                    {
                      
                        var cellValue1  = cells[0].ToString(); // row.Cell( 1).GetValue<string>(); // no

                        var cellValue2  = cells[1].ToString(); // row.Cell( 2).GetValue<string>(); // manufacturer name
                        var cellValue3  = cells[2].ToString(); // row.Cell( 3).GetValue<string>(); // third party manuf name
                        var cellValue4  = cells[3].ToString(); // row.Cell( 4).GetValue<string>(); // material number
                        var cellValue5  = cells[4].ToString(); // row.Cell( 5).GetValue<string>(); // product name
                        var cellValue6  = cells[5].ToString(); // row.Cell( 6).GetValue<string>(); // description
                         
                        var cellValue8  = Convert.ToDecimal(cells[7].ToString()); // row.Cell( 8).GetValue<Decimal>(); // purchaseValue
                      //  var cellValue9  = Convert.ToDecimal(cells[8].ToString()); // row.Cell( 9).GetValue<Decimal>(); // totalPrice
                        var cellValue10 = Convert.ToDecimal(cells[9].ToString()); // row.Cell(10).GetValue<Decimal>(); // purchaseValueUSD
                     //   var cellValue11 = Convert.ToDecimal(cells[10].ToString()); // row.Cell(11).GetValue<Decimal>(); // totalUSD

                        var cellValue12 = cells[11].ToString(); //row.Cell(12).GetValue<string>(); // paymentTerms
                        var cellValue13 = cells[12].ToString(); //row.Cell(13).GetValue<string>(); // purchaseOrderNo

                        var cellValue14 = cells[13].ToString(); //row.Cell(14).GetValue<string>(); // contractNo
                        var cellValue15 = cells[14].ToString(); //row.Cell(15).GetValue<string>(); // signOff
                        var cellValue16 = cells[15].ToString(); //row.Cell(16).GetValue<string>(); // remark
                        var cellValue17 = cells[16].ToString(); //row.Cell(17).GetValue<string>(); // materialsSent
                        var cellValue18 = cells[17].ToString(); //row.Cell(18).GetValue<string>(); // department
                        var cellValue19 = cells[18].ToString(); //row.Cell(19).GetValue<string>(); // manager

                        var cellValue20 = cells[19].ToString(); //row.Cell(20).GetValue<string>(); // fixedAssetNumber
                        var cellValue21 = cells[20].ToString(); //row.Cell(21).GetValue<string>(); // serialNumber
                        var cellValue22 = cells[21].ToString(); //row.Cell(22).GetValue<string>(); // location
                        var cellValue23 = cells[22].ToString(); //row.Cell(23).GetValue<string>(); // PIC
                        // var cellValue24 = cells[23].ToString(); //row.Cell(24).GetValue<string>(); // NOTE

                        /*
                            var cellValue25 = cells[24].ToString(); // category
                            var cellValue26 = cells[25].ToString(); // netBookValue
                            var cellValue27 = cells[26].ToString(); // usefulLife
                            var cellValue28 = cells[27].ToString(); // capitalizationDate
                            var cellValue29 = cells[28].ToString(); // accumulatedDepreciation
                        */

                        cmd.Parameters.AddWithValue("@manufacturerName", cellValue2);
                        cmd.Parameters.AddWithValue("@partyManufacturerName", cellValue3);  // TBD se elimina?
                        cmd.Parameters.AddWithValue("@materialNumber", cellValue4);
                        cmd.Parameters.AddWithValue("@productName", cellValue5);
                        cmd.Parameters.AddWithValue("@description", cellValue6);
                         
                        cmd.Parameters.AddWithValue("@purchaseValue", cellValue8);
                    //    cmd.Parameters.AddWithValue("@totalPrice", cellValue9);
                        cmd.Parameters.AddWithValue("@purchaseValueUSD", cellValue10);
                       // cmd.Parameters.AddWithValue("@totalUSD", cellValue11);

                        cmd.Parameters.AddWithValue("@paymentTerms", cellValue12);
                        cmd.Parameters.AddWithValue("@purchaseOrderNo", cellValue13);

                        cmd.Parameters.AddWithValue("@contractNo", cellValue14);
                        cmd.Parameters.AddWithValue("@signOff", cellValue15);
                        cmd.Parameters.AddWithValue("@remark", cellValue16);
                        cmd.Parameters.AddWithValue("@materialsSent", cellValue17);
                        cmd.Parameters.AddWithValue("@department", cellValue18);
                        cmd.Parameters.AddWithValue("@manager", cellValue19);

                        cmd.Parameters.AddWithValue("@fixedAssetNumber", cellValue20);
                        cmd.Parameters.AddWithValue("@serialNumber", cellValue21);
                        cmd.Parameters.AddWithValue("@location", cellValue22);
                        cmd.Parameters.AddWithValue("@PIC", cellValue23);
                         
                        cmd.Parameters.AddWithValue("@createdBy", userName);
                    //    cmd.Parameters.AddWithValue("@NOTE", cellValue24);
                    /*
                        cmd.Parameters.AddWithValue("@category", cellValue25);
                        cmd.Parameters.AddWithValue("@netBookValue", cellValue26);
                        cmd.Parameters.AddWithValue("@usefulLife", cellValue27);
                        cmd.Parameters.AddWithValue("@capitalizationDate", cellValue28);
                        cmd.Parameters.AddWithValue("@accumulatedDepreciation", cellValue29);
                    */
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        lstErrors.Add("Error in line "+ cnt  +":"+ex.Message);
                    }
                    cnt++;
                    Debug.WriteLine(cnt);
                }
                tran.Commit();
            }
            if (lstErrors.Count > 0)
            {
                ViewBag.Message = "Ocurrieron errores al intentar cargar el archivo: " + string.Join(Environment.NewLine, lstErrors);
            }else
                ViewBag.Message = "Archivo subido correctamente!";
            

            //return View();
            return View("Index");
        }

        [HttpGet]
        public IActionResult GetTableData()
        { 
            var dr = new List<object>();
            SqlConnection con = new SqlConnection(_connStr);
            try
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]", con);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var record = new
                    {
                        id                  = reader.IsDBNull(0) ?"0" : reader["id"].ToString(),
                        manufacturerName    = reader.IsDBNull(1) ? "" : reader["manufacturerName"].ToString(),
                        materialNumber      = reader.IsDBNull(2) ? "" : reader["materialNumber"].ToString(),
                        productName         = reader.IsDBNull(3) ? "" : reader["productName"].ToString(),
                        description         = reader.IsDBNull(4) ? "" : reader["description"].ToString(),
                        purchaseValue           = reader.IsDBNull(6) ? "" : reader["purchaseValue"].ToString(),
                        netBookValue        = reader.IsDBNull(7) ? "" : reader["netBookValue"].ToString()
                    };
                    dr.Add(record);
                }
            }
            catch(Exception ex)
            {
            }
            return Json(new
            {
                data = dr
            });
        }
          
        [HttpPost]
        public JsonResult CreateItem( FixedAssetItem item)
        {
            string message = "";    
            SqlConnection con = new SqlConnection(_connStr);
            con.Open();
            try
            {

                using (SqlTransaction tran = con.BeginTransaction())
                {

                    string insertQuery = string.Format("INSERT INTO [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] "
                       + "         ( {0} )"
                       + " VALUES  ( {1} )", string.Join(",", lstCols), "@"+string.Join(",@", lstCols));

                    SqlCommand cmd = new SqlCommand(insertQuery, con, tran);
                    //    cmd.Parameters.AddWithValue("@id", item.id); 
                    cmd.Parameters.AddWithValue("@manufacturerName", item.manufacturerName);
                    cmd.Parameters.AddWithValue("@partyManufacturerName", item.partyManufacturerName);      // TBD se elimina?
                    cmd.Parameters.AddWithValue("@materialNumber", item.materialNumber);
                    cmd.Parameters.AddWithValue("@productName", item.productName);
                    cmd.Parameters.AddWithValue("@description", item.description);
                     
                    cmd.Parameters.AddWithValue("@purchaseValue", item.purchaseValue);
            //        cmd.Parameters.AddWithValue("@totalPrice", item.totalPrice);
            //        cmd.Parameters.AddWithValue("@purchaseValueUSD", item.purchaseValueUSD);
             //       cmd.Parameters.AddWithValue("@totalUSD", item.totalUSD);

                    cmd.Parameters.AddWithValue("@paymentTerms", item.paymentTerms);
                    cmd.Parameters.AddWithValue("@purchaseOrderNo", item.purchaseOrderNo);

                    cmd.Parameters.AddWithValue("@contractNo", item.contractNo);
                    cmd.Parameters.AddWithValue("@signOff", item.signOff);
                    cmd.Parameters.AddWithValue("@remark", item.remark);
                    cmd.Parameters.AddWithValue("@materialsSent", item.materialsSent);
                    cmd.Parameters.AddWithValue("@department", item.department);
                    cmd.Parameters.AddWithValue("@manager", item.manager);

                    cmd.Parameters.AddWithValue("@fixedAssetNumber", item.fixedAssetNumber);
                    cmd.Parameters.AddWithValue("@serialNumber", item.serialNumber);
                    cmd.Parameters.AddWithValue("@location", item.location);
                    cmd.Parameters.AddWithValue("@PIC", item.PIC);


                    cmd.Parameters.AddWithValue("@createdBy", item.updatedBy);
                    cmd.ExecuteNonQuery();
                    tran.Commit();
                }
                message = "Item was created";
            }
            catch (Exception ex)
            {
                message = "Couldn't create item";
            }
            return Json(new { msg = message });
        } 

        [HttpPost]
        public JsonResult EditItem(FixedAssetItem item)
        {
            string message = "";
            string identity = item.updatedBy;
            if( item.id == 0)
                return Json(new { msg = "Item ID is required" });
            SqlConnection con = new SqlConnection(_connStr);
            con.Open();
            try
            { 
                using (SqlTransaction tran = con.BeginTransaction())
                {
                    string insertQuery = "UPDATE [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] " +
                    "   SET manufacturerName=@manufacturerName, " +
                    "partyManufacturerName=@partyManufacturerName, " +
                    "materialNumber=@materialNumber, " +
                    "productName=@productName, " +
                    "description=@description, " +
                    "purchaseValue=@purchaseValue, " + 
                    "paymentTerms=@paymentTerms, " +
                    "purchaseOrderNo=@purchaseOrderNo, " +
                    "contractNo=@contractNo, " +
                    "signOff=@signOff, " +
                    "remark=@remark, " +
                    "materialsSent=@materialsSent, " +
                    "department=@department, " +
                    "manager=@manager, " +
                    "fixedAssetNumber=@fixedAssetNumber, " +
                    "serialNumber=@serialNumber, " +
                    "location=@location, " +
                    "PIC=@PIC, " +
                    "netBookValue=@netBookValue, " +
                    "usefulLife=@usefulLife, " + 
                    "capitalizationDate=@capitalizationDate " +
                    "WHERE Id=@id";

                    SqlCommand cmd = new SqlCommand(insertQuery, con, tran);
                    cmd.Parameters.AddWithValue("@id", item.id);

                    cmd.Parameters.AddWithValue("@manufacturerName", string.IsNullOrEmpty(item.manufacturerName) ? "" : item.manufacturerName);
                    cmd.Parameters.AddWithValue("@partyManufacturerName", string.IsNullOrEmpty(item.partyManufacturerName) ? "" : item.partyManufacturerName);      // TBD se elimina?
                    cmd.Parameters.AddWithValue("@materialNumber", string.IsNullOrEmpty(item.materialNumber) ? "" : item.materialNumber);
                    cmd.Parameters.AddWithValue("@productName", string.IsNullOrEmpty(item.productName) ? "" : item.productName);
                    cmd.Parameters.AddWithValue("@description", string.IsNullOrEmpty(item.description) ? "" : item.description);

                    cmd.Parameters.AddWithValue("@purchaseValue", item.purchaseValue);
              //      cmd.Parameters.AddWithValue("@purchaseValueUSD", item.purchaseValueUSD);

                    cmd.Parameters.AddWithValue("@paymentTerms", string.IsNullOrEmpty(item.paymentTerms) ? "" : item.paymentTerms);
                    cmd.Parameters.AddWithValue("@purchaseOrderNo", string.IsNullOrEmpty(item.purchaseOrderNo) ? "" : item.purchaseOrderNo);

                    cmd.Parameters.AddWithValue("@contractNo", string.IsNullOrEmpty(item.contractNo) ? "" : item.contractNo);
                    cmd.Parameters.AddWithValue("@signOff", string.IsNullOrEmpty(item.signOff) ? "" : item.signOff);
                    cmd.Parameters.AddWithValue("@remark", string.IsNullOrEmpty(item.remark) ? "" : item.remark);
                    cmd.Parameters.AddWithValue("@materialsSent", string.IsNullOrEmpty(item.materialsSent) ? "" : item.materialsSent);
                    cmd.Parameters.AddWithValue("@department", string.IsNullOrEmpty(item.department) ? "" : item.department);
                    cmd.Parameters.AddWithValue("@manager", string.IsNullOrEmpty(item.manager) ? "" : item.manager); 

                    cmd.Parameters.AddWithValue("@fixedAssetNumber", string.IsNullOrEmpty(item.fixedAssetNumber) ? "" : item.fixedAssetNumber);
                    cmd.Parameters.AddWithValue("@serialNumber", string.IsNullOrEmpty(item.serialNumber) ? "" : item.serialNumber);

                    // IF NEW LOCATION IS DIFFERENT THAN OLD LOCATION, SAVE REGISTER
                    string insertLog = "INSERT INTO LocationChangeLog (location, idAsset, updatedBy, dateOfChange, description) "
                    + "SELECT @location, @idAsset, @updatedBy, getdate(), '' "
                    + "WHERE (SELECT COUNT(*) cnt FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] "
                    + "WHERE id = @idAsset AND location <> @location) > 0";
                    SqlCommand cmd2 = new SqlCommand(insertLog, con, tran);
                    cmd2.Parameters.AddWithValue("@location", item.location);
                    cmd2.Parameters.AddWithValue("@idAsset", item.id);
                    cmd2.Parameters.AddWithValue("@updatedBy", identity);
                    cmd2.ExecuteNonQuery();

                    cmd.Parameters.AddWithValue("@location", string.IsNullOrEmpty(item.location) ? "" : item.location);
                    cmd.Parameters.AddWithValue("@PIC", string.IsNullOrEmpty(item.PIC) ? "" : item.PIC);

                    cmd.Parameters.AddWithValue("@accumulatedDepreciation", item.accumulatedDepreciation);
                    cmd.Parameters.AddWithValue("@netBookValue", item.netBookValue);
                    cmd.Parameters.AddWithValue("@usefulLife", item.usefulLife);
                    cmd.Parameters.AddWithValue("@capitalizationDate", item.capitalizationDate);

                    cmd.ExecuteNonQuery();
                    tran.Commit();
                    message = "Item was modified";
                }
            }
            catch(Exception ex)
            {
                message = "An error has ocurred, data couldn't be saved";
            } 
            return Json(new { msg = message });
        }

        [HttpGet]
        public IActionResult GetItem(int id)
        {
            var d = new FixedAssetItem();
            SqlConnection c = new SqlConnection(_connStr);
            try
            {
                c.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] WHERE id=@id", c);
                cmd.Parameters.AddWithValue("@id", id);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    d = new FixedAssetItem
                    {
                        id = reader["Id"] == DBNull.Value ? 0 : Convert.ToInt32(reader["Id"]),
                        manufacturerName = reader["manufacturerName"] == DBNull.Value ? "" : reader["manufacturerName"].ToString(),
                        partyManufacturerName = reader["partyManufacturerName"] == DBNull.Value ? "" : reader["partyManufacturerName"].ToString(),      // TBD se elimina?
                        materialNumber = reader["materialNumber"] == DBNull.Value ? "" : reader["materialNumber"].ToString(),
                        productName = reader["productName"] == DBNull.Value ? "" : reader["productName"].ToString(),
                        description = reader["description"] == DBNull.Value ? "" : reader["description"].ToString(),
                        purchaseValue = reader["manufacturerName"] == DBNull.Value ? 0 : (float)Convert.ToDouble(reader["purchaseValue"]), 
                        paymentTerms = reader["paymentTerms"] == DBNull.Value ? "" : reader["paymentTerms"].ToString(),
                        purchaseOrderNo = reader["purchaseOrderNo"] == DBNull.Value ? "" : reader["purchaseOrderNo"].ToString(),
                        contractNo = reader["contractNo"] == DBNull.Value ? "" : reader["contractNo"].ToString(),
                        signOff = reader["signOff"] == DBNull.Value ? "" : reader["signOff"].ToString(),
                        remark = reader["remark"] == DBNull.Value ? "" : reader["remark"].ToString(),
                        materialsSent = reader["materialsSent"] == DBNull.Value ? "" : reader["materialsSent"].ToString(),
                        department = reader["department"] == DBNull.Value ? "" : reader["department"].ToString(),
                        manager = reader["manager"] == DBNull.Value ? "" : reader["manager"].ToString(),
                        fixedAssetNumber = reader["fixedAssetNumber"] == DBNull.Value ? "" : reader["fixedAssetNumber"].ToString(),
                        serialNumber = reader["serialNumber"] == DBNull.Value ? "" : reader["serialNumber"].ToString(),
                        location = reader["location"] == DBNull.Value ? "" : reader["location"].ToString(),
                        PIC = reader["PIC"] == DBNull.Value ? "" : reader["PIC"].ToString(),
                        accumulatedDepreciation = reader["accumulatedDepreciation"] == DBNull.Value ? 0 : (float)Convert.ToDecimal(reader["accumulatedDepreciation"].ToString()),
                        netBookValue = reader["netBookValue"] == DBNull.Value ? 0 : (float)Convert.ToDecimal(reader["netBookValue"].ToString()),
                        usefulLife = reader["usefulLife"] == DBNull.Value ? 0 : Convert.ToInt32(reader["usefulLife"].ToString()),
                        capitalizationDate = reader["capitalizationDate"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["capitalizationDate"].ToString())
                    };
                }
            }
            catch(Exception ex)
            {
                d = null;
            }

            var json = JsonSerializer.Serialize(new { data = d });
            return Content(json, "application/json");
        }

        [HttpGet]
        public JsonResult DeleteItem(int id)
        {
            string message = "";
            SqlConnection c = new SqlConnection(_connStr);
            c.Open();
            try
            {
                using (SqlTransaction tran = c.BeginTransaction())
                {
                    string deleteQuery = "DELETE [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]  " +
                       "WHERE Id=@id";

                    SqlCommand cmd = new SqlCommand(deleteQuery, c, tran);
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                    tran.Commit();
                }
                message = "ITEM WAS REMOVED";
            }
            catch(Exception ex)
            {
                message = "Couldn't remove item";
            }
            return Json(new {msg=message});
        }
 
         
    }
}
