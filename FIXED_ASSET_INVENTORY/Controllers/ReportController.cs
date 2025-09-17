using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace FIXED_ASSET_INVENTORY.Controllers
{
    [Authorize]
    public class ReportController : Controller
    {
        private readonly string _connStr;
        public ReportController(IConfiguration configuration)
        {
            _connStr = configuration.GetConnectionString("PSGDbConnStr");
        }
         
        public IActionResult Index()
        {
            return View();
        }
    }
}
