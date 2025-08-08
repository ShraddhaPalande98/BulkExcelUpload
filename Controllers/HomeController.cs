using ExcelBulkUploadApp.Data;
using ExcelBulkUploadApp.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Diagnostics;
using System.Globalization;

namespace ExcelBulkUploadApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly ApplicationDbContext _applicationDbContext;

        public HomeController(IWebHostEnvironment webHostEnvironment, ApplicationDbContext applicationDbContext)
        {
            _webHostEnvironment = webHostEnvironment;
            _applicationDbContext = applicationDbContext;
        }

        [HttpGet]
        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var uploadPath = Path.Combine(_webHostEnvironment.WebRootPath, "Uploads");

                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                var fileName = Path.GetFileNameWithoutExtension(file.FileName);
                var fileExtension = Path.GetExtension(file.FileName);
                var uniqueFileName = $"{fileName}_{DateTime.Now.ToString("yyyyMMddHHmmssfff")}{fileExtension}";
                var filePath = Path.Combine(uploadPath, uniqueFileName);

                try
                {
                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(fileStream);
                    }

                    var products = new List<Product>();
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var rowCount = worksheet.Dimension.Rows;

                        for (int row = 3; row <= rowCount; row++)
                        {
                            var product = new Product
                            {
                                Name = worksheet.Cells[row, 1].Text,
                                Description = worksheet.Cells[row, 2].Text,
                                Price = decimal.Parse(worksheet.Cells[row, 3].Text),
                                Quantity = int.Parse(worksheet.Cells[row, 4].Text),
                            };

                            products.Add(product);
                        }
                    }
                    _applicationDbContext.Product.AddRange(products);
                    await _applicationDbContext.SaveChangesAsync();
                    TempData["SuccessMessage"] = "File uploaded and data saved successfully!";
                    return RedirectToAction("Upload");
                }
                catch (Exception ex)
                {
                    TempData["ErrorMessage"] = $"An error occurred while processing the file: {ex.Message}";
                    return RedirectToAction("Upload");
                }
            }
            TempData["ErrorMessage"] = "Please upload a valid file.";
            return RedirectToAction("Upload");
        }
    }
}
