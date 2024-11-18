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
    public class ThêmSảnPhẩm
    {
        IWebDriver driver;
        [TestInitialize]
        public void Init()
        {

            driver = new EdgeDriver();
            //driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
        }
        [TestMethod]
        public void them()
        {
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/admin/");
            driver.Manage().Window.Size = new System.Drawing.Size(1552, 832);
            
            driver.FindElement(By.ClassName("btn btn-success mb-1")).Click();
            driver.FindElement(By.CssSelector("p span")).Click();
            driver.FindElement(By.CssSelector(".btn-success")).Click();
            driver.FindElement(By.Id("title")).Click();
            driver.FindElement(By.Id("title")).SendKeys("cà phê");
            driver.FindElement(By.Id("summary")).Click();
            driver.FindElement(By.Id("summary")).SendKeys("Cà phê đen đậm đà");
            driver.FindElement(By.Id("price")).Click();
            driver.FindElement(By.Id("price")).SendKeys("25");
            driver.FindElement(By.Id("qty")).Click();
            driver.FindElement(By.Id("qty")).SendKeys("24");
            driver.FindElement(By.Id("btnThem")).Click();
            driver.FindElement(By.Id("file")).Click();
            driver.FindElement(By.Id("file")).SendKeys("C:\\fakepath\\hat-ca-phe-va-nhung-dieu-thu-vi-4-1280x800 (1).jpg");
            driver.FindElement(By.Id("btnThem")).Click();
        }
    }
}