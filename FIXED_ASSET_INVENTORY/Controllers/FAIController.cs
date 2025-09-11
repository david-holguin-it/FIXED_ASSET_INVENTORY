using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using FIXED_ASSET_INVENTORY.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Diagnostics.Contracts;
using System.Net;

namespace FIXED_ASSET_INVENTORY.Controllers
{
    public class FAIController : Controller
    { 
        private readonly string _connStr;
        public FAIController(IConfiguration configuration)
        {
            _connStr=configuration.GetConnectionString("PSGDbConnStr");
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
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
             //   var rows = worksheet.RowsUsed().Skip(1); // Omite la fila de encabezado
                var rows = worksheet.Rows()
                                        .Where(row => !row.IsEmpty())
                                        .Skip(2);
                SqlConnection c = new SqlConnection(_connStr);
                c.Open();
                   int cnt = 0;
                List<string> lstErrors = new List<string>();
                using (SqlTransaction tran = c.BeginTransaction())
                {
                    foreach (var row in rows)
                    {
                     //   if (cnt >= 10) break;
                        if (row.IsEmpty()) continue;
                        string insertQuery = "INSERT INTO [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] " 
                           + "(         manufacturerName, partyManufacturerName, materialNumber, productName, description, quantity,  unitPrice,  totalPrice,  unitPriceUSD,  totalUSD,  paymentTerms,  purchaseOrderNo,  contractNo,  signOff,  remark,  materialsSent,  department,  manager,  fixedAssetNumber,  serialNumber,  location,  PIC,  NOTE )"
                           + " VALUES (@manufacturerName,@partyManufacturerName,@materialNumber,@productName,@description,@quantity, @unitPrice, @totalPrice, @unitPriceUSD, @totalUSD, @paymentTerms, @purchaseOrderNo, @contractNo, @signOff, @remark, @materialsSent, @department, @manager, @fixedAssetNumber, @serialNumber, @location, @PIC, @NOTE )";

                        SqlCommand cmd = new SqlCommand(insertQuery, c, tran);
                        try
                        {

                            var cellValue1  = row.Cell( 1).GetValue<string>(); // no

                            var cellValue2  = row.Cell( 2).GetValue<string>(); // manufacturer name
                            var cellValue3  = row.Cell( 3).GetValue<string>(); // third party manuf name
                            var cellValue4  = row.Cell( 4).GetValue<string>(); // material number
                            var cellValue5  = row.Cell( 5).GetValue<string>(); // product name
                            var cellValue6  = row.Cell( 6).GetValue<string>(); // description

                            var cellValue7  = row.Cell( 7).GetValue<Decimal>(); // quantity
                            var cellValue8  = row.Cell( 8).GetValue<Decimal>(); // unitPrice
                            var cellValue9  = row.Cell( 9).GetValue<Decimal>(); // totalPrice
                            var cellValue10 = row.Cell(10).GetValue<Decimal>(); // unitPriceUSD
                            var cellValue11 = row.Cell(11).GetValue<Decimal>(); // totalUSD

                            var cellValue12 = row.Cell(12).GetValue<string>(); // paymentTerms
                            var cellValue13 = row.Cell(13).GetValue<string>(); // purchaseOrderNo

                            var cellValue14 = row.Cell(14).GetValue<string>(); // contractNo
                            var cellValue15 = row.Cell(15).GetValue<string>(); // signOff
                            var cellValue16 = row.Cell(16).GetValue<string>(); // remark
                            var cellValue17 = row.Cell(17).GetValue<string>(); // materialsSent
                            var cellValue18 = row.Cell(18).GetValue<string>(); // department
                            var cellValue19 = row.Cell(19).GetValue<string>(); // manager

                            var cellValue20 = row.Cell(20).GetValue<string>(); // fixedAssetNumber
                            var cellValue21 = row.Cell(21).GetValue<string>(); // serialNumber
                            var cellValue22 = row.Cell(22).GetValue<string>(); // location
                            var cellValue23 = row.Cell(23).GetValue<string>(); // PIC
                            var cellValue24 = row.Cell(24).GetValue<string>(); // NOTE

                            cmd.Parameters.AddWithValue("@manufacturerName", cellValue2);
                            cmd.Parameters.AddWithValue("@partyManufacturerName", cellValue3);
                            cmd.Parameters.AddWithValue("@materialNumber", cellValue4);
                            cmd.Parameters.AddWithValue("@productName", cellValue5);
                            cmd.Parameters.AddWithValue("@description", cellValue6);

                            cmd.Parameters.AddWithValue("@quantity", cellValue7);
                            cmd.Parameters.AddWithValue("@unitPrice", cellValue8);
                            cmd.Parameters.AddWithValue("@totalPrice", cellValue9);
                            cmd.Parameters.AddWithValue("@unitPriceUSD", cellValue10);
                            cmd.Parameters.AddWithValue("@totalUSD", cellValue11);

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
                            cmd.Parameters.AddWithValue("@NOTE", cellValue24);
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            lstErrors.Add("Error in line "+ cnt  +":"+ex.Message);
                        }
                        cnt++;
                    }
                    tran.Commit();
                }
                if (lstErrors.Count > 0)
                {
                    ViewBag.Message = "Ocurrieron errores al intentar cargar el archivo: " + String.Join(Environment.NewLine, lstErrors);
                }else
                    ViewBag.Message = "Archivo subido correctamente!";
            

            //return View();
            return View("Index");
        }

        [HttpGet]
        public IActionResult GetTableData()
        { 
            var dr = new List<Object>();
            SqlConnection c = new SqlConnection("Server=10.95.2.52; Database=FIXED_ASSET_INVENTORY;User Id=IMPIT;Password=PSG+123.;TrustServerCertificate=True");
            c.Open();
            SqlCommand cmd = new SqlCommand("SELECT * FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]", c);
            SqlDataReader reader = cmd.ExecuteReader(); 
            while(reader.Read())
            {
                var d = new
                {  
                    id=  reader.IsDBNull(0) ? 0.ToString() : reader["id"].ToString(), 
                    manufactureName =  reader.IsDBNull(1) ? "" : reader["manufacturerName"].ToString(),
                    materialNumber = reader.IsDBNull(2) ? "" : reader["materialNumber"].ToString(),
                    productName = reader.IsDBNull(3) ? "" : reader["productName"].ToString(),
                    description = reader.IsDBNull(4) ? "" : reader["description"].ToString(),
                    quantity = reader.IsDBNull(5) ? "" : reader["quantity"].ToString(),
                    unitPrice = reader.IsDBNull(6) ? "" : reader["unitPrice"].ToString(),
                    totalPrice = reader.IsDBNull(7) ? "" : reader["totalPrice"].ToString()
                };
                dr.Add(d);
            }
            return Json(new
            {
                data = dr
            });
        }
        
          
        [HttpPost]
        public JsonResult CreateItem([FromBody] string s)
        {

            //SqlConnection c = new SqlConnection("Server=10.95.2.52; Database=FIXED_ASSET_INVENTORY;User Id=IMPIT;Password=PSG+123.;TrustServerCertificate=True");
            //c.Open();
            //try
            //{

            //    using (SqlTransaction tran = c.BeginTransaction())
            //    {

            //        string insertQuery = "INSERT INTO [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] "
            //           + "         (manufacturerName, partyManufacturerName, materialNumber, productName, description, quantity,  unitPrice,  totalPrice,  unitPriceUSD,  totalUSD,  paymentTerms,  purchaseOrderNo,  contractNo,  signOff,  remark,  materialsSent,  department,  manager,  fixedAssetNumber,  serialNumber,  location,  PIC,  NOTE )"
            //           + " VALUES (@manufacturerName,@partyManufacturerName,@materialNumber,@productName,@description,@quantity, @unitPrice, @totalPrice, @unitPriceUSD, @totalUSD, @paymentTerms, @purchaseOrderNo, @contractNo, @signOff, @remark, @materialsSent, @department, @manager, @fixedAssetNumber, @serialNumber, @location, @PIC, @NOTE )";

            //        SqlCommand cmd = new SqlCommand(insertQuery, c, tran);
            //    //    cmd.Parameters.AddWithValue("@id", item.id);

            //        cmd.Parameters.AddWithValue("@manufacturerName",        item.manufacturerName);
            //        cmd.Parameters.AddWithValue("@partyManufacturerName",   item.partyManufacturerName);
            //        cmd.Parameters.AddWithValue("@materialNumber",          item.materialNumber);
            //        cmd.Parameters.AddWithValue("@productName",             item.productName);
            //        cmd.Parameters.AddWithValue("@description",             item.description);

            //        cmd.Parameters.AddWithValue("@quantity",                item.quantity);
            //        cmd.Parameters.AddWithValue("@unitPrice",               item.unitPrice);
            //        cmd.Parameters.AddWithValue("@totalPrice",              item.totalPrice);
            //        cmd.Parameters.AddWithValue("@unitPriceUSD",            item.unitPriceUSD);
            //        cmd.Parameters.AddWithValue("@totalUSD",                item.totalUSD);

            //        cmd.Parameters.AddWithValue("@paymentTerms",            item.paymentTerms);
            //        cmd.Parameters.AddWithValue("@purchaseOrderNo",         item.purchaseOrderNo);

            //        cmd.Parameters.AddWithValue("@contractNo",              item.contractNo);
            //        cmd.Parameters.AddWithValue("@signOff",                 item.signOff);
            //        cmd.Parameters.AddWithValue("@remark",                  item.remark);
            //        cmd.Parameters.AddWithValue("@materialsSent",           item.materialsSent);
            //        cmd.Parameters.AddWithValue("@department",              item.department);
            //        cmd.Parameters.AddWithValue("@manager",                 item.manager);

            //        cmd.Parameters.AddWithValue("@fixedAssetNumber",        item.fixedAssetNumber);
            //        cmd.Parameters.AddWithValue("@serialNumber",            item.serialNumber);
            //        cmd.Parameters.AddWithValue("@location",                item.location);
            //        cmd.Parameters.AddWithValue("@PIC",                     item.PIC);
            //        cmd.Parameters.AddWithValue("@NOTE",                    item.NOTE);
            //        cmd.ExecuteNonQuery();
            //        tran.Commit();
            //    }
            //    return Json(new { msg = "Item was created" });
            //}
            //catch(Exception ex)
            //{
            //    return Json(new { msg = "Couldn't create item" });
            //}
            return Json(new { msg = "Couldn't create item" });
        } 

        [HttpPost]
        public JsonResult EditItem(FixedAssetItem item)
        {
            if( item.id == 0)
                return Json(new { msg = "Item ID is required" });
            SqlConnection c = new SqlConnection("Server=10.95.2.52; Database=FIXED_ASSET_INVENTORY;User Id=IMPIT;Password=PSG+123.;TrustServerCertificate=True");
            c.Open();
            using (SqlTransaction tran = c.BeginTransaction())
            {  
                    string insertQuery = "UPDATE [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] "
                       +" SET manufacturerName=@manufacturerName," +
                       "partyManufacturerName=@partyManufacturerName," +
                       "materialNumber=@materialNumber," +
                       "productName=@productName," +
                       "description=@description," +
                       "quantity=@quantity, " +
                       "unitPrice=@unitPrice, " +
                       "totalPrice=@totalPrice, " +
                       "unitPriceUSD=@unitPriceUSD, " +
                       "totalUSD=@totalUSD, " +
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
                       "NOTE=@NOTE " +
                       "WHERE Id=@id";


                    SqlCommand cmd = new SqlCommand(insertQuery, c, tran); 
                    cmd.Parameters.AddWithValue("@id", item.id);

                    cmd.Parameters.AddWithValue("@manufacturerName", string.IsNullOrEmpty(item.manufacturerName) ?"":item.manufacturerName);
                    cmd.Parameters.AddWithValue("@partyManufacturerName", string.IsNullOrEmpty(item.partyManufacturerName) ? "" : item.partyManufacturerName);
                    cmd.Parameters.AddWithValue("@materialNumber", string.IsNullOrEmpty(item.materialNumber) ? "" : item.materialNumber);
                    cmd.Parameters.AddWithValue("@productName", string.IsNullOrEmpty(item.productName) ? "" : item.productName);
                    cmd.Parameters.AddWithValue("@description", string.IsNullOrEmpty(item.description) ? "" : item.description);

                    cmd.Parameters.AddWithValue("@quantity", item.quantity);
                    cmd.Parameters.AddWithValue("@unitPrice", item.unitPrice);
                    cmd.Parameters.AddWithValue("@totalPrice", item.totalPrice);
                    cmd.Parameters.AddWithValue("@unitPriceUSD", item.unitPriceUSD);
                    cmd.Parameters.AddWithValue("@totalUSD", item.totalUSD);

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
                    cmd.Parameters.AddWithValue("@location", string.IsNullOrEmpty(item.location) ? "" : item.location);
                    cmd.Parameters.AddWithValue("@PIC", string.IsNullOrEmpty(item.PIC) ? "" : item.PIC);
                    cmd.Parameters.AddWithValue("@NOTE", string.IsNullOrEmpty(item.NOTE) ? "" : item.NOTE);
                    cmd.ExecuteNonQuery(); 
                tran.Commit();
            }
            return Json(new { msg = "Item was modified" });
        }

        [HttpGet]
        public JsonResult GetItem(int id)
        {
            var dr = new List<Object>();
            SqlConnection c = new SqlConnection("Server=10.95.2.52; Database=FIXED_ASSET_INVENTORY;User Id=IMPIT;Password=PSG+123.;TrustServerCertificate=True");
            c.Open();
            SqlCommand cmd = new SqlCommand("SELECT * FROM [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV] WHERE id=@id", c);
            cmd.Parameters.AddWithValue("@id",id);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var d = new
                {
                    id = reader.IsDBNull(0) ? 0.ToString() : reader["Id"].ToString(),
                    manufacturerName = reader.IsDBNull(1) ? "" : reader["manufacturerName"].ToString(),
                    partyManufacturerName = reader.IsDBNull(1) ? "" : reader["partyManufacturerName"].ToString(),
                    materialNumber = reader.IsDBNull(2) ? "" : reader["materialNumber"].ToString(),
                    productName = reader.IsDBNull(3) ? "" : reader["productName"].ToString(),
                    description = reader.IsDBNull(4) ? "" : reader["description"].ToString(),
                    quantity = reader.IsDBNull(5) ? "" : reader["quantity"].ToString(),
                    unitPrice = reader.IsDBNull(6) ? "" : reader["unitPrice"].ToString(),
                    totalPrice = reader.IsDBNull(7) ? "" : reader["totalPrice"].ToString(),
                    unitPriceUSD = reader.IsDBNull(7) ? "" : reader["unitPriceUSD"].ToString(),
                    totalUSD = reader.IsDBNull(7) ? "" : reader["totalUSD"].ToString(),
                    paymentTerms = reader.IsDBNull(7) ? "" : reader["paymentTerms"].ToString(),
                    purchaseOrderNo = reader.IsDBNull(7) ? "" : reader["purchaseOrderNo"].ToString(),
                    contractNo = reader.IsDBNull(7) ? "" : reader["contractNo"].ToString(),
                    signOff = reader.IsDBNull(7) ? "" : reader["signOff"].ToString(),
                    remark = reader.IsDBNull(7) ? "" : reader["remark"].ToString(),
                    materialsSent = reader.IsDBNull(7) ? "" : reader["materialsSent"].ToString(),
                    department = reader.IsDBNull(7) ? "" : reader["department"].ToString(),
                    manager = reader.IsDBNull(7) ? "" : reader["manager"].ToString(),
                    fixedAssetNumber = reader.IsDBNull(7) ? "" : reader["fixedAssetNumber"].ToString(),
                    serialNumber = reader.IsDBNull(7) ? "" : reader["serialNumber"].ToString(),
                    location = reader.IsDBNull(7) ? "" : reader["location"].ToString(),
                    PIC = reader.IsDBNull(7) ? "" : reader["PIC"].ToString(),
                    NOTE = reader.IsDBNull(7) ? "" : reader["NOTE"].ToString()  
                };
                dr.Add(d);
            }
            return Json(new
            {
                data = dr
            });
        }

        [HttpGet]
        public JsonResult DeleteItem(int id)
        {
            SqlConnection c = new SqlConnection("Server=10.95.2.52; Database=FIXED_ASSET_INVENTORY;User Id=IMPIT;Password=PSG+123.;TrustServerCertificate=True");
            c.Open();
            using (SqlTransaction tran = c.BeginTransaction())
            {
                string deleteQuery = "DELETE [FIXED_ASSET_INVENTORY].[dbo].[FIXED_ASSETS_INV]  " +
                   "WHERE Id=@id";
                
                SqlCommand cmd = new SqlCommand(deleteQuery, c, tran);
                cmd.Parameters.AddWithValue("@id", id); 
                cmd.ExecuteNonQuery();
                tran.Commit();
            }
            return Json(new { msg = "ITEM WAS REMOVED" });
        }
 
         
    }
}
