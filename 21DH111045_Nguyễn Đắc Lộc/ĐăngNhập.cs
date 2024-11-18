using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.CSharp;
using System.Security.Cryptography.X509Certificates;
using NUnit.Framework;
using Assert = NUnit.Framework.Assert;
using OpenQA.Selenium.Interactions;

namespace UnitTestProject1
{
    [TestClass]
    public class ĐăngNhập
    {
        IWebDriver driver;
        [TestInitialize]
        public void Init()
        {

            driver = new EdgeDriver();
            //driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
        }

        [TestMethod]
        public void Dang_nhap_20()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            String tk = xlRange.Cells[3][30].value.ToString();
            driver.FindElement(By.Name("username")).Click();
            driver.FindElement(By.Name("username")).SendKeys(tk);
            String mk = xlRange.Cells[4][30].value.ToString();
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).SendKeys(mk);
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
        }
        [TestMethod]
        public void Dang_nhap_21()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            String tk = xlRange.Cells[3][31].value.ToString();
            driver.FindElement(By.Name("username")).Click();
            driver.FindElement(By.Name("username")).SendKeys(tk);
            String mk = xlRange.Cells[4][31].value.ToString();
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).SendKeys(mk);
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
        }
        [TestMethod]
        public void Dang_nhap_22()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            //String tk = xlRange.Cells[3][31].value.ToString();
            //driver.FindElement(By.Name("username")).Click();
            //driver.FindElement(By.Name("username")).SendKeys(tk);
            String mk = xlRange.Cells[4][32].value.ToString();
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).SendKeys(mk);
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
        }
        [TestMethod]
        public void Dang_nhap_23()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            String tk = xlRange.Cells[3][33].value.ToString();
            driver.FindElement(By.Name("username")).Click();
            driver.FindElement(By.Name("username")).SendKeys(tk);
            //String mk = xlRange.Cells[4][31].value.ToString();
            //driver.FindElement(By.Name("password")).Click();
            //driver.FindElement(By.Name("password")).SendKeys(mk);
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
        }
        [TestMethod]
        public void Dang_nhap_24()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            //String tk = xlRange.Cells[3][34].value.ToString();
            //driver.FindElement(By.Name("username")).Click();
            //driver.FindElement(By.Name("username")).SendKeys(tk);
            //String mk = xlRange.Cells[4][34].value.ToString();
            //driver.FindElement(By.Name("password")).Click();
            //driver.FindElement(By.Name("password")).SendKeys(mk);
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
        }
        [TestMethod]
        public void Dang_nhap_25()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            String tk = xlRange.Cells[3][35].value.ToString();
            driver.FindElement(By.Name("username")).Click();
            driver.FindElement(By.Name("username")).SendKeys(tk);
            String mk = xlRange.Cells[4][35].value.ToString();
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).SendKeys(mk);
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
        }
        [TestMethod]
        public void Dang_nhap_26()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            String tk = xlRange.Cells[3][36].value.ToString();
            driver.FindElement(By.Name("username")).Click();
            driver.FindElement(By.Name("username")).SendKeys(tk);
            String mk = xlRange.Cells[4][36].value.ToString();
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).SendKeys(mk);
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
        }
        [TestMethod]
        public void Dang_nhap_27()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
        }
        [TestMethod]
        public void Dang_nhap_28()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/?page=homepage");
            driver.Manage().Window.Size = new System.Drawing.Size(866, 920);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            driver.FindElement(By.Name("username")).Click();
            driver.FindElement(By.Name("username")).SendKeys("lo");
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).SendKeys("lo");
            driver.FindElement(By.Id("rememberme")).Click();
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
            driver.FindElement(By.ClassName("setting__active")).Click();
            driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[1]/div[3]/ul/li[3]/div/div/div/div/div[2]/div/div/div/span[5]/a")).Click();
            Assert.That(driver.SwitchTo().Alert().Text, Is.EqualTo("Bạn có chắc chắn muốn thoát ?"));
            driver.FindElement(By.CssSelector(".setting__active")).Click();
        }
        [TestMethod]
        public void Dang_nhap_29()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase1.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel          
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/?page=homepage");
            driver.Manage().Window.Size = new System.Drawing.Size(866, 920);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            driver.FindElement(By.Name("username")).Click();
            driver.FindElement(By.Name("username")).SendKeys("lo");
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).SendKeys("lo");
            driver.FindElement(By.Id("rememberme")).Click();
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
            driver.FindElement(By.ClassName("setting__active")).Click();
            driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[1]/div[3]/ul/li[3]/div/div/div/div/div[2]/div/div/div/span[5]/a")).Click();
            Assert.That(driver.SwitchTo().Alert().Text, Is.EqualTo("Bạn có chắc chắn muốn thoát ?"));
            driver.FindElement(By.CssSelector(".setting__active")).Click();
        }
    }
}