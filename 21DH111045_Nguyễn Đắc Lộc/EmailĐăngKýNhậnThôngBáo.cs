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
    public class EmailĐăngKýNhậnThôngBáo
    {
        IWebDriver driver;
        [TestInitialize]
        public void Init()
        {

            driver = new EdgeDriver();
            //driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
        }
        [TestMethod]
        public void Email_dang_ky_nhan_thong_bao_45()
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
    String email = xlRange.Cells[3][65].value.ToString();
    driver.FindElement(By.CssSelector(".newsletter__box > input")).Click();
    driver.FindElement(By.CssSelector(".newsletter__box > input")).SendKeys(email);
    driver.FindElement(By.CssSelector("button:nth-child(2)")).Click();

}

[TestMethod]
public void Email_dang_ky_nhan_thong_bao_46()
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
    String email = xlRange.Cells[3][66].value.ToString();
    driver.FindElement(By.CssSelector(".newsletter__box > input")).Click();
    driver.FindElement(By.CssSelector(".newsletter__box > input")).SendKeys(email);
    driver.FindElement(By.CssSelector("button:nth-child(2)")).Click();

}
[TestMethod]
public void Email_dang_ky_nhan_thong_bao_47()
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
    String email = xlRange.Cells[3][67].value.ToString();
    driver.FindElement(By.CssSelector(".newsletter__box > input")).Click();
    driver.FindElement(By.CssSelector(".newsletter__box > input")).SendKeys(email);
    driver.FindElement(By.CssSelector("button:nth-child(2)")).Click();
    Thread.Sleep(5000);
    driver.Close();
}

[TestMethod]
public void Email_dang_ky_nhan_thong_bao_48()
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
    String email = xlRange.Cells[3][68].value.ToString();
    driver.FindElement(By.CssSelector(".newsletter__box > input")).Click();
    driver.FindElement(By.CssSelector(".newsletter__box > input")).SendKeys(email);
    driver.FindElement(By.CssSelector("button:nth-child(2)")).Click();
    Thread.Sleep(5000);
    driver.Close();
}

[TestMethod]
public void Email_dang_ky_nhan_thong_bao_50()
{
    driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
    driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
    Thread.Sleep(1000);
    {
        var element = driver.FindElement(By.CssSelector(".owl-item:nth-child(7) > .product .second__img"));
        Actions builder = new Actions(driver);
        builder.MoveToElement(element).ClickAndHold().Perform();
    }
    Thread.Sleep(1000);
    {
        var element = driver.FindElement(By.CssSelector(".owl-item:nth-child(7) > .product .second__img"));
        Actions builder = new Actions(driver);
        builder.MoveToElement(element).Perform();
    }
    Thread.Sleep(1000);
    {
        var element = driver.FindElement(By.CssSelector(".owl-item:nth-child(7) > .product .second__img"));
        Actions builder = new Actions(driver);
        builder.MoveToElement(element).Release().Perform();
    }
    Thread.Sleep(1000);
    driver.FindElement(By.CssSelector(".owl-item:nth-child(7) > .product .second__img")).Click();
    {
        var element = driver.FindElement(By.CssSelector(".wn__product__area"));
        Actions builder = new Actions(driver);
        builder.MoveToElement(element).ClickAndHold().Perform();
    }
    Thread.Sleep(1000);
    {
        var element = driver.FindElement(By.CssSelector(".wn__product__area"));
        Actions builder = new Actions(driver);
        builder.MoveToElement(element).Perform();
    }
    Thread.Sleep(1000);
    {
        var element = driver.FindElement(By.CssSelector(".wn__product__area"));
        Actions builder = new Actions(driver);
        builder.MoveToElement(element).Release().Perform();
    }
    driver.FindElement(By.CssSelector(".wn__product__area")).Click();
    driver.FindElement(By.CssSelector(".wn__product__area")).Click();
}
[TestMethod]
public void Email_dang_ky_nhan_thong_bao_49()
{
    driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/?page=order");
    driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
    Thread.Sleep(1000);
    driver.FindElement(By.CssSelector(".with--one--item > a")).Click();
    Thread.Sleep(1000);
    driver.FindElement(By.CssSelector(".drop:nth-child(2) > a")).Click();
    Thread.Sleep(1000);
    driver.FindElement(By.CssSelector(".meninmenu > li:nth-child(4) > a")).Click();
    Thread.Sleep(1000);
    driver.FindElement(By.LinkText("Hello Nguyễn Lộc")).Click();
}
    }



 }
