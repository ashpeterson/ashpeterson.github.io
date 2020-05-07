using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ReadExcel.Models;

namespace ReadExcel.Controllers
{
    public class UserController : Controller
    {
        // Default GET method
        public IActionResult Index()
        {
            return View(new List<UserModel>());
        }

        [HttpPost]
        public IActionResult Index(IFormCollection form)
        {
            List<UserModel> users = new List<UserModel>();

            var fileName = "./Users.xlsx";
            // For .net core, the next line requires the NuGet package, 
            // System.Text.Encoding.CodePages
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read()) //Each row of the file
                    {
                        users.Add(new UserModel
                        {
                            Name = reader.GetValue(0).ToString(),
                            Email = reader.GetValue(1).ToString(),
                            Phone = reader.GetValue(2).ToString()
                        });
                    }
                }
            }
            return View(users);
        }
    }
}