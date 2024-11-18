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
    public class ĐăngKý
    {
        IWebDriver driver;
        [TestInitialize]
        public void Init()
        {

            driver = new EdgeDriver();
            //driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
        }
        [TestMethod]
        public void Dang_ky_01()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][2].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][2].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][2].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][2].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][2].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][2].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][2].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][2].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys("2024-03-08");
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_02()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][8].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][8].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][8].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][8].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][8].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][8].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][8].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][8].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys("2024-03-08");
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }

        [TestMethod]
        public void Dang_ky_03()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][4].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][4].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][4].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][4].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][4].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][4].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][4].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][4].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys("2024-03-08");
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_04()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][5].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][5].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][5].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][5].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][5].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][5].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][5].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][5].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys("2024-03-08");
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }

        [TestMethod]
        public void Dang_ky_05()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][6].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][6].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][6].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][6].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][6].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][6].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][6].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][6].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys("2024-03-08");
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_06()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][7].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][7].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][7].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][7].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][7].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][7].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][7].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][7].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][7].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_07()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][8].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][8].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][8].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][8].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][8].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][8].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][8].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][8].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][8].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_08()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][9].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][9].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][9].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][9].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][9].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][9].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][9].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][9].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][9].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_09()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][10].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][10].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][10].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][10].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][10].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][10].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][10].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][10].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][10].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_10()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][11].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][11].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][11].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][11].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][11].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][11].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][11].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][11].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][11].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_11()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][12].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][12].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][12].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][12].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][12].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][12].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][12].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][12].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][12].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_12()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][13].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys("");
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][13].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][13].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][13].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][13].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][13].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][13].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][13].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][13].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_13_1()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            //driver.FindElement(By.Name("lastname")).Click();
            //String lastname = xlRange.Cells[3][14].value.ToString();
            //driver.FindElement(By.Name("lastname")).SendKeys("");
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][14].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][14].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][14].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][14].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][14].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][14].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][14].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][14].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_13_2()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][15].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            ////String firstname = xlRange.Cells[4][15].value.ToString();
            ////driver.FindElement(By.Name("firstname")).Click();
            ////driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][15].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][15].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][15].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][15].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][15].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][15].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][15].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_13_3()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][16].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][16].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][16].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][16].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            //String email = xlRange.Cells[5][16].value.ToString();
            //driver.FindElement(By.Name("email")).Click();
            //driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][16].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][16].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][16].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][16].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_13_4()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][17].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][17].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            //String phone = xlRange.Cells[6][16].value.ToString();
            //driver.FindElement(By.Name("phone_number")).Click();
            //driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][17].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][17].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][17].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][17].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][17].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][17].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_13_5()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][18].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][18].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][18].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            //String address = xlRange.Cells[7][18].value.ToString();
            //driver.FindElement(By.Name("address")).Click();
            //driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][18].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][18].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][18].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][18].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][18].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_13_6()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][19].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][19].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][19].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][19].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][19].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            //String tk = xlRange.Cells[8][19].value.ToString();
            //driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            //driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][19].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][19].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][19].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_14()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][23].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][23].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][23].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][23].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][23].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][23].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][23].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][23].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][23].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_15()
        {
            Excel.Application dataApp;
            Excel.Workbook dataWorkbook;
            Excel.Worksheet dataWorksheet;
            Excel.Range xlRange;

            //Tìm file Excel và mở nó lên
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"C:\Users\PC\Desktop\Testcase.xlsx");
            dataWorksheet = dataWorkbook.Sheets[2];
            xlRange = dataWorksheet.UsedRange;
            //Lấy data từ Excel         
            driver.Navigate().GoToUrl("http://localhost:8080/Website/PHP/");
            driver.Manage().Window.Size = new System.Drawing.Size(1722, 936);
            driver.FindElement(By.CssSelector(".setting__active")).Click();
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][24].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][24].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][24].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][24].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][24].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][24].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][24].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][24].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][24].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_16()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][25].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][25].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][25].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][25].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][25].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][25].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][25].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][25].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][25].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_17()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][26].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][26].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][26].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][26].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][26].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][26].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][26].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][26].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][26].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_18()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.Name("lastname")).Click();
            String lastname = xlRange.Cells[3][27].value.ToString();
            driver.FindElement(By.Name("lastname")).SendKeys(lastname);
            driver.FindElement(By.CssSelector(".my_account_area .row")).Click();
            String firstname = xlRange.Cells[4][27].value.ToString();
            driver.FindElement(By.Name("firstname")).Click();
            driver.FindElement(By.Name("firstname")).SendKeys(firstname);
            String phone = xlRange.Cells[6][27].value.ToString();
            driver.FindElement(By.Name("phone_number")).Click();
            driver.FindElement(By.Name("phone_number")).SendKeys(phone);
            String address = xlRange.Cells[7][27].value.ToString();
            driver.FindElement(By.Name("address")).Click();
            driver.FindElement(By.Name("address")).SendKeys(address);
            String email = xlRange.Cells[5][27].value.ToString();
            driver.FindElement(By.Name("email")).Click();
            driver.FindElement(By.Name("email")).SendKeys(email);
            String tk = xlRange.Cells[8][27].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).Click();
            driver.FindElement(By.CssSelector(".input__box:nth-child(6) > input")).SendKeys(tk);
            String pass = xlRange.Cells[9][27].value.ToString();
            driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).Click();
            String repass = xlRange.Cells[10][27].value.ToString(); driver.FindElement(By.CssSelector(".input__box:nth-child(7) > input")).SendKeys(pass);
            driver.FindElement(By.Name("pre_password")).Click();
            driver.FindElement(By.Name("pre_password")).SendKeys(repass);
            string year = xlRange.Cells[12][27].value.ToString();
            driver.FindElement(By.Name("date")).Click();
            driver.FindElement(By.Name("date")).SendKeys(year);
            driver.FindElement(By.CssSelector("input:nth-child(3)")).Click();
            driver.FindElement(By.Name("date")).SendKeys("0002-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0020-03-23");
            driver.FindElement(By.Name("date")).SendKeys("0200-03-23");
            driver.FindElement(By.Name("date")).SendKeys("2003-03-23");
            driver.FindElement(By.CssSelector(".btn")).Click();
            driver.FindElement(By.CssSelector(".btn")).Click();


            //dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();

        }
        [TestMethod]
        public void Dang_ky_19()
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
            driver.FindElement(By.LinkText("Đăng kí tài khoản")).Click();
            driver.FindElement(By.ClassName("breadcrumb_item")).Click();
        }
    }
}