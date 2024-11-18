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
    public class ThôngTinCáNhân
    {
        IWebDriver driver;
        [TestInitialize]
        public void Init()
        {

            driver = new EdgeDriver();
            //driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
        }

        [TestMethod]
        public void Thong_tin_ca_nhan_35()
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
            driver.FindElement(By.Name("username")).Click();
            driver.FindElement(By.Name("username")).SendKeys("lo");
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).SendKeys("lo");
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
            driver.FindElement(By.LinkText("Hello Nguyễn Lộc")).Click();
            String firstname = xlRange.Cells[4][36].value.ToString();
            driver.FindElement(By.Name("btnAddMore")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            driver.FindElement(By.Name("lastname")).SendKeys(firstname);
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).Click();
            {
                var element = driver.FindElement(By.Name("firstname"));
                Actions builder = new Actions(driver);
                builder.DoubleClick(element).Perform();
            }
            String lastname = xlRange.Cells[5][36].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(lastname);
            String email = xlRange.Cells[6][36].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            {
                var element = driver.FindElement(By.Name("email"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element).ClickAndHold().Perform();
            }
            {
                var element = driver.FindElement(By.Name("email"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element).Perform();
            }
            {
                var element = driver.FindElement(By.Name("email"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element).Release().Perform();
            }
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            driver.FindElement(By.CssSelector("#something .col-5:nth-child(1)")).Click();
            String pass = xlRange.Cells[7][36].value.ToString();
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).Click();
            {
                var element = driver.FindElement(By.Name("password"));
                Actions builder = new Actions(driver);
                builder.DoubleClick(element).Perform();
            }
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.CssSelector(".col-5:nth-child(2) > .form-group:nth-child(5) > .col-md-12")).Click();
            driver.FindElement(By.CssSelector("input:nth-child(4)")).Click();
            String repass = xlRange.Cells[8][36].value.ToString();
            driver.FindElement(By.Name("password2")).Click();
            driver.FindElement(By.Name("password2")).SendKeys(repass);
            driver.FindElement(By.CssSelector(".btn-success")).Click();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys("rioroll23032003@gmail.com");
            driver.FindElement(By.CssSelector(".btn-success")).Click();
        }

        [TestMethod]
        public void Thong_tin_ca_nhan_37()
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
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/?page=user");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            driver.FindElement(By.Name("username")).Click();
            driver.FindElement(By.Name("username")).SendKeys("lo");
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).SendKeys("lo");
            driver.FindElement(By.CssSelector(".form__btn > button")).Click();
            driver.FindElement(By.LinkText("Hello Nguyễn Lộc")).Click();
            String firstname = xlRange.Cells[4][36].value.ToString();
            driver.FindElement(By.Name("btnAddMore")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            driver.FindElement(By.Name("lastname")).SendKeys(firstname);
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).Click();
            {
                var element = driver.FindElement(By.Name("firstname"));
                Actions builder = new Actions(driver);
                builder.DoubleClick(element).Perform();
            }
            String lastname = xlRange.Cells[5][36].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(lastname);
            String email = xlRange.Cells[6][36].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            {
                var element = driver.FindElement(By.Name("email"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element).ClickAndHold().Perform();
            }
            {
                var element = driver.FindElement(By.Name("email"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element).Perform();
            }
            {
                var element = driver.FindElement(By.Name("email"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element).Release().Perform();
            }
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            driver.FindElement(By.CssSelector("#something .col-5:nth-child(1)")).Click();
            String pass = xlRange.Cells[7][36].value.ToString();
            driver.FindElement(By.Name("password")).Click();
            driver.FindElement(By.Name("password")).Click();
            {
                var element = driver.FindElement(By.Name("password"));
                Actions builder = new Actions(driver);
                builder.DoubleClick(element).Perform();
            }
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.CssSelector(".col-5:nth-child(2) > .form-group:nth-child(5) > .col-md-12")).Click();
            driver.FindElement(By.CssSelector("input:nth-child(4)")).Click();

            driver.FindElement(By.CssSelector(".btn-success")).Click();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys("rioroll23032003@gmail.com");
            driver.FindElement(By.CssSelector(".btn-success")).Click();
        }
    }
}