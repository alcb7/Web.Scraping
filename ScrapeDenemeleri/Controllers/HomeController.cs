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
            int currentPage = 1; // Mevcut sayfa numaras�
            int lastPage = 603; // Son sayfa numaras�, �rne�in
            int fileCount = 1; // Dosya say�s�

            try
            {
                driver = new ChromeDriver(); // Chrome WebDriver'� ba�lat

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
                        worksheet.Cell(1, 1).Value = "�sim";
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

                                        // Sayfan�n tamamen y�klendi�inden emin olmak i�in bekleme s�resi ekleyin
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
                                            // Telefon elementini bulamazsa bo� b�rak
                                        }
                                        IWebElement nameElement = driver.FindElement(By.XPath("/html/body/div[2]/div[1]/div/h3"));

                                        string address = addressElement.Text;
                                        string phone = phoneElement?.Text ?? "Telefon bulunamad�"; // Telefon bulunamazsa varsay�lan de�er
                                        string name = nameElement.Text;

                                        worksheet.Cell(currentRow, 1).Value = name;
                                        worksheet.Cell(currentRow, 2).Value = address;
                                        worksheet.Cell(currentRow, 3).Value = phone;
                                        currentRow++;

                                        driver.Navigate().Back(); // Geri d�n

                                        // Sayfan�n geri d�nmesini ve elementlerin yeniden y�klenmesini bekleyin
                                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//table[@class='table-striped']/tbody/tr/td[2]/a")));
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Eleman i�lenirken hata olu�tu: " + ex.Message);
                                    }
                                }

                                currentPage++;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Sayfa i�lenirken hata olu�tu: " + ex.Message);
                            }
                        }

                        workbook.SaveAs(filePath); // Excel dosyas�n� kaydet
                        fileCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                // E�er WebDriver olu�turulurken bir hata olu�ursa, bu hatay� yakalay�n ve bir hata mesaj� d�nd�r�n
                return Content("WebDriver olu�turulurken bir hata olu�tu: " + ex.Message);
            }
            finally
            {
                driver?.Quit(); // WebDriver'� kapat
            }

            return RedirectToAction("Index"); // Anasayfaya d�n
        }

    }
}
