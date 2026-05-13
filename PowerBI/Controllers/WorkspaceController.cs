using Microsoft.AspNetCore.Mvc;
using PowerBI.Models;
using PowerBI.Services;
using PowerBI.Data;

namespace PowerBI.Controllers
{
    public class WorkspaceController : Controller
    {
        private readonly PowerBIService _service;
        private readonly AppDbContext _db;

        public WorkspaceController(PowerBIService service, AppDbContext db)
        {
            _service = service;
            _db = db;
        }

        public async Task<IActionResult> Index()
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            if (userId == null) return RedirectToAction("Login", "Auth");

            Console.WriteLine($"[WORKSPACE] Index requested for User ID: {userId}");

            try
            {
                // Sync with Power BI Service on every load to ensure "Complete Sync"
                await _service.SyncWorkspaces(userId.Value, _db);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[WORKSPACE] Sync Error: {ex.Message}");
                ViewBag.SyncError = "Failed to sync with Power BI: " + ex.Message;
            }

            var workspaces = _db.Workspaces
                .AsEnumerable()
                .GroupBy(w => w.PowerBIWorkspaceId)
                .Select(g => g.First())
                .ToList();

            Console.WriteLine($"[WORKSPACE] Returning {workspaces.Count} synced workspaces.");
            return View(workspaces);
        }

        [HttpPost]
        public async Task<IActionResult> Create(string name)
        {
            Console.WriteLine($"[WORKSPACE] Creating/Linking workspace: {name}");
            var userId = HttpContext.Session.GetInt32("UserId");

            if (userId == null) return RedirectToAction("Login", "Auth");

            try 
            {
                var ws = await _service.GetOrCreateWorkspace(name);
                
                // The sync in Index will handle the DB update, but we can do it here for immediate feedback
                var existing = _db.Workspaces.FirstOrDefault(w => w.PowerBIWorkspaceId == ws.Id.ToString());
                if (existing == null)
                {
                    _db.Workspaces.Add(new Workspace
                    {
                        Name = ws.Name,
                        PowerBIWorkspaceId = ws.Id.ToString(),
                        UserId = userId.Value
                    });
                    await _db.SaveChangesAsync();
                }

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                if (message.Contains("Forbidden"))
                {
                    message = "Access Forbidden. Please ensure: 1. 'Allow service principals to use Power BI APIs' is ENABLED in the Power BI Admin Portal. 2. The Service Principal has 'Workspace.ReadWrite.All' or similar API permissions in Azure AD.";
                }
                TempData["Error"] = "Power BI Error: " + message;
                return RedirectToAction("Index");
            }
        }

        [HttpPost]
        public async Task<IActionResult> Rename(int id, string newName)
        {
            var ws = _db.Workspaces.Find(id);
            if (ws == null) return NotFound();

            try
            {
                if (string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) return BadRequest("Invalid Workspace ID");
                await _service.RenameWorkspace(new Guid(ws.PowerBIWorkspaceId), newName);
                ws.Name = newName;
                await _db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Rename Error: " + ex.Message;
                return RedirectToAction("Index");
            }
        }

        [HttpPost]
        public async Task<IActionResult> Delete(int id)
        {
            var ws = _db.Workspaces.Find(id);
            if (ws == null) return NotFound();

            try
            {
                if (string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) return BadRequest("Invalid Workspace ID");
                await _service.DeleteWorkspace(new Guid(ws.PowerBIWorkspaceId));
                _db.Workspaces.Remove(ws);
                await _db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Delete Error: " + ex.Message;
                return RedirectToAction("Index");
            }
        }
    }
}
