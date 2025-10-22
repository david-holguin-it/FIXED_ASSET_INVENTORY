using FIXED_ASSET_INVENTORY.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Configuration;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Security.Claims;

namespace FIXED_ASSET_INVENTORY.Controllers
{
    public class LoginRequest
    {
        public string Username { get; set; }
        public string Password { get; set; }
    }

    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IHttpClientFactory _httpClientFactory;
        string _connStr = "";

        public HomeController(ILogger<HomeController> logger,IConfiguration configuration, IHttpClientFactory httpClientFactory)
        {
            _logger = logger;
            _httpClientFactory = httpClientFactory;
            if (configuration.GetConnectionString("PSGDbConnStr") != null)
                _connStr = configuration.GetConnectionString("PSGDbConnStr"); // Connection string from appsettings.json
        }

        [Authorize]
        public IActionResult Index()
        {
            //using (var pc = new PrincipalContext(ContextType.Domain, "iec.inventec"))
            //{
            //    bool isValid = pc.ValidateCredentials("IMX109294", "Redento2.0Redento2.0");
            //    if (isValid)
            //    {
            //        // Crear cookie de sesión, login exitoso
            //    }
            //    else
            //    {
            //        // Fallo
            //    }
            //}
            var username = User.Identity.Name; // DOMAIN\usuario
            var isAuthenticated = User.Identity.IsAuthenticated;
            return View();
        }

        [AllowAnonymous]
        [HttpGet]
        public IActionResult Login()
        { 
            return View();
        }

        [AllowAnonymous]
        [HttpPost]
        public async Task<IActionResult> Login(LoginRequest model)
        {
            using (var context = new PrincipalContext(ContextType.Domain, "iec.inventec"))
            {
                if (!context.ValidateCredentials(model.Username, model.Password))
                {
                    ModelState.AddModelError("", "Usuar or password are invalid");
                    return View(model);
                } 
                if (Users.userExists(model.Username, _connStr) != "OK")
                {
                    ModelState.AddModelError("", "You don't have access to this app.");
                    return View(model);
                }
            }

            var claims = new List<Claim> { new Claim(ClaimTypes.Name, model.Username) };
            var identity = new ClaimsIdentity(claims, "MyCookieAuth");
            var principal = new ClaimsPrincipal(identity);

            await HttpContext.SignInAsync("MyCookieAuth", principal);

            return RedirectToAction("Index", "Home");
        }

        // Logout
        [HttpPost("Logout")]
        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            return RedirectToAction("Login");
        }


        [AllowAnonymous]
        public IActionResult Privacy()
        {
            return View();
        }

        [AllowAnonymous]
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
