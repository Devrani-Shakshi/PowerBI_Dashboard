using Microsoft.AspNetCore.Mvc;
using PowerBI.Models;
using PowerBI.Data;

namespace PowerBI.Controllers
{
    public class AuthController : Controller
    {
        private readonly AppDbContext _db;

        public AuthController(AppDbContext db)
        {
            _db = db;
        }

        public IActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Login(string email, string password)
        {
            Console.WriteLine($"[AUTH] Login attempt for: {email}");
            
            var user = _db.Users
                .FirstOrDefault(u => u.Email == email && u.Password == password);

            if (user == null)
            {
                Console.WriteLine($"[AUTH] Login failed - Invalid credentials for: {email}");
                ViewBag.Error = "Invalid credentials";
                return View();
            }

            Console.WriteLine($"[AUTH] Login successful. User: {user.Name} (ID: {user.Id}, Role: {user.Role})");
            HttpContext.Session.SetInt32("UserId", user.Id);
            HttpContext.Session.SetString("UserRole", user.Role ?? "User");

            return RedirectToAction("Index", "Workspace");
        }

        public IActionResult Logout()
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            Console.WriteLine($"[AUTH] Logout for User ID: {userId}");
            HttpContext.Session.Clear();
            return RedirectToAction("Login");
        }
    }
}
