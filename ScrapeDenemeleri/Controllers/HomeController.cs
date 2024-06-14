using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using ScrapeDenemeleri.Models;
using System.Diagnostics;
using System.Collections.ObjectModel;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using ClosedXML.Excel;


namespace ScrapeDenemeleri.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _env;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment env)
        {
            _logger = logger;
            _env = env;
        }

        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public IActionResult ClickButton()
        {
            IWebDriver driver = null;
            int currentPage = 1; // Mevcut sayfa numarasý
            int lastPage = 603; // Son sayfa numarasý, örneðin
            int fileCount = 1; // Dosya sayýsý

            try
            {
                driver = new ChromeDriver(); // Chrome WebDriver'ý baþlat

                while (currentPage <= lastPage)
                {
                    string directory = Path.Combine(_env.ContentRootPath, "App_Data");
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    string filePath = Path.Combine(directory, $"veri_{fileCount}.xlsx");

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Veriler");
                        worksheet.Cell(1, 1).Value = "Ýsim";
                        worksheet.Cell(1, 2).Value = "Adres";
                        worksheet.Cell(1, 3).Value = "Telefon";

                        int currentRow = 2;
                        int pageLimit = currentPage + 199; // Her dosyada 10 sayfa

                        while (currentPage <= lastPage && currentPage <= pageLimit)
                        {
                            try
                            {
                                driver.Navigate().GoToUrl($"https://e.dto.org.tr/Members?page={currentPage}"); // Mevcut sayfaya git

                                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//table[@class='table-striped']/tbody/tr/td[2]/a")));

                                IReadOnlyCollection<IWebElement> elements = driver.FindElements(By.XPath("//table[@class='table-striped']/tbody/tr/td[2]/a"));
                                foreach (var element in elements)
                                {
                                    try
                                    {
                                        element.Click();

                                        // Sayfanýn tamamen yüklendiðinden emin olmak için bekleme süresi ekleyin
                                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[2]/div[3]/div/p[1]")));

                                        // Elementleri yeniden bul
                                        IWebElement addressElement = driver.FindElement(By.XPath("/html/body/div[2]/div[3]/div/p[2]"));
                                        IWebElement phoneElement = null;
                                        try
                                        {
                                            phoneElement = driver.FindElement(By.XPath("/html/body/div[2]/div[3]/div/p[3]"));
                                        }
                                        catch (NoSuchElementException)
                                        {
                                            // Telefon elementini bulamazsa boþ býrak
                                        }
                                        IWebElement nameElement = driver.FindElement(By.XPath("/html/body/div[2]/div[1]/div/h3"));

                                        string address = addressElement.Text;
                                        string phone = phoneElement?.Text ?? "Telefon bulunamadý"; // Telefon bulunamazsa varsayýlan deðer
                                        string name = nameElement.Text;

                                        worksheet.Cell(currentRow, 1).Value = name;
                                        worksheet.Cell(currentRow, 2).Value = address;
                                        worksheet.Cell(currentRow, 3).Value = phone;
                                        currentRow++;

                                        driver.Navigate().Back(); // Geri dön

                                        // Sayfanýn geri dönmesini ve elementlerin yeniden yüklenmesini bekleyin
                                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//table[@class='table-striped']/tbody/tr/td[2]/a")));
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Eleman iþlenirken hata oluþtu: " + ex.Message);
                                    }
                                }

                                currentPage++;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Sayfa iþlenirken hata oluþtu: " + ex.Message);
                            }
                        }

                        workbook.SaveAs(filePath); // Excel dosyasýný kaydet
                        fileCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                // Eðer WebDriver oluþturulurken bir hata oluþursa, bu hatayý yakalayýn ve bir hata mesajý döndürün
                return Content("WebDriver oluþturulurken bir hata oluþtu: " + ex.Message);
            }
            finally
            {
                driver?.Quit(); // WebDriver'ý kapat
            }

            return RedirectToAction("Index"); // Anasayfaya dön
        }

    }
}
