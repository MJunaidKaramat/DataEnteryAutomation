using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Drawing;
using System.Security.Policy;
using System.Timers;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Net;

namespace VectorSurveillanceDataEntry
{
    [TestClass]
    public class UnitTest1
    {   
        private TestContext testContext;
        public TestContext TestContext
        {
            get { return testContext; }
            set { testContext = value; }
        }

        string url = "https://vectorsurveillance.punjab.gov.pk/";
        string browser = "chrome";
        string uName = "ddho.muridke";
        string uPassword = "03334688537";
        //Indoor Form
        public static IEnumerable<object[]> indoorExcelData()
        {
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo("IndoorUpdated.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                var rowCount = worksheet.Dimension.End.Row;
                for (int i = 1; i <= rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                        worksheet.Cells[i,3].Value?.ToString().Trim(),
                        worksheet.Cells[i,4].Value?.ToString().Trim(),
                        worksheet.Cells[i,5].Value?.ToString().Trim(),
                        worksheet.Cells[i,6].Value?.ToString().Trim(),
                        worksheet.Cells[i,7].Value?.ToString().Trim(),
                        worksheet.Cells[i,8].Value?.ToString().Trim(),
                        worksheet.Cells[i,9].Value?.ToString().Trim(),
                        worksheet.Cells[i,10].Value?.ToString().Trim(),
                        worksheet.Cells[i,11].Value?.ToString().Trim(),
                        worksheet.Cells[i,12].Value?.ToString().Trim(),
                        worksheet.Cells[i,13].Value?.ToString().Trim(),
                        worksheet.Cells[i,14].Value?.ToString().Trim(),
                        worksheet.Cells[i,15].Value?.ToString().Trim(),
                        worksheet.Cells[i,16].Value?.ToString().Trim(),
                        worksheet.Cells[i,17].Value?.ToString().Trim(),
                        worksheet.Cells[i,18].Value?.ToString().Trim(),
                        worksheet.Cells[i,19].Value?.ToString().Trim(),
                        worksheet.Cells[i,20].Value?.ToString().Trim(),
                        worksheet.Cells[i,21].Value?.ToString().Trim(),
                        worksheet.Cells[i,22].Value?.ToString().Trim(),
                        worksheet.Cells[i,23].Value?.ToString().Trim(),
                        worksheet.Cells[i,24].Value?.ToString().Trim(),
                        worksheet.Cells[i,25].Value?.ToString().Trim(),
                        worksheet.Cells[i,26].Value?.ToString().Trim(),
                        worksheet.Cells[i,27].Value?.ToString().Trim(),
                        worksheet.Cells[i,28].Value?.ToString().Trim(),
                        worksheet.Cells[i,29].Value?.ToString().Trim(),
                        worksheet.Cells[i,30].Value?.ToString().Trim(),
                        worksheet.Cells[i,31].Value?.ToString().Trim(),
                        worksheet.Cells[i,32].Value?.ToString().Trim(),
                        worksheet.Cells[i,33].Value?.ToString().Trim(),
                        worksheet.Cells[i,34].Value?.ToString().Trim(),
                        worksheet.Cells[i,35].Value?.ToString().Trim(),
                        worksheet.Cells[i,36].Value?.ToString().Trim(),
                        worksheet.Cells[i,37].Value?.ToString().Trim(),
                        worksheet.Cells[i,38].Value?.ToString().Trim()
                    };
                }
            }
        }
        [DynamicData(nameof(indoorExcelData), DynamicDataSourceType.Method)]
        [TestMethod]
        public void IndoorForm(string uc, string locality, string houseCheckNo, string housePositive,
            string airCondition, string airConditionP, string tap, string tapP, string tire, string tireP,
            string flower, string flowerP, string drink, string drinkP, string wash, string washP, string debris,
            string debrisP, string bird, string birdP, string animal, string animalP, string garbage, string garbageP,
            string waterTank, string waterTankP, string refrigrate, string refrigrateP, string mainHole, string mainHoleP,
            string houseWater, string houseWaterP, string cooler, string coolerP, string houseJunk, string houseJunkP,
            string roofJunk, string roofJunkP)
        {

            bool linkOpen = false;
            IndoorEntry entry = new IndoorEntry();
            if (!linkOpen)
            {
                entry.browserSelection(browser);
                entry.landingPage(url);
                entry.LoginForm(uName, uPassword);
                linkOpen = true;
            }

            if (linkOpen)
            {
             entry.IndoorDataEntryForm(uc, locality, houseCheckNo, housePositive, airCondition, airConditionP, tap, tapP,
             tire, tireP, flower, flowerP, drink, drinkP, wash, washP, debris, debrisP, bird, birdP, animal, animalP,
             garbage, garbageP, waterTank, waterTankP, refrigrate, refrigrateP, mainHole, mainHoleP, houseWater,
             houseWaterP, cooler, coolerP, houseJunk, houseJunkP, roofJunk, roofJunkP);
            }
            entry.closeBrowser();
        }


        public static IEnumerable<object[]> outdoorExcelData()
        {
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo("OutdoorUpdated.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet2"];
                var rowCount = worksheet.Dimension.End.Row;
                for (int i = 1; i <= rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                        worksheet.Cells[i,3].Value?.ToString().Trim(),
                        worksheet.Cells[i,4].Value?.ToString().Trim(),
                        worksheet.Cells[i,5].Value?.ToString().Trim(),
                        worksheet.Cells[i,6].Value?.ToString().Trim(),
                        worksheet.Cells[i,7].Value?.ToString().Trim(),
                        worksheet.Cells[i,8].Value?.ToString().Trim(),
                        worksheet.Cells[i,9].Value?.ToString().Trim(),
                        worksheet.Cells[i,10].Value?.ToString().Trim(),
                        worksheet.Cells[i,11].Value?.ToString().Trim(),
                        worksheet.Cells[i,12].Value?.ToString().Trim(),
                        worksheet.Cells[i,13].Value?.ToString().Trim(),
                        worksheet.Cells[i,14].Value?.ToString().Trim(),
                        worksheet.Cells[i,15].Value?.ToString().Trim(),
                        worksheet.Cells[i,16].Value?.ToString().Trim(),
                        worksheet.Cells[i,17].Value?.ToString().Trim(),
                        worksheet.Cells[i,18].Value?.ToString().Trim(),
                        worksheet.Cells[i,19].Value?.ToString().Trim(),
                        worksheet.Cells[i,20].Value?.ToString().Trim(),
                        worksheet.Cells[i,21].Value?.ToString().Trim(),
                        worksheet.Cells[i,22].Value?.ToString().Trim(),
                        worksheet.Cells[i,23].Value?.ToString().Trim(),
                        worksheet.Cells[i,24].Value?.ToString().Trim(),
                        worksheet.Cells[i,25].Value?.ToString().Trim(),
                        worksheet.Cells[i,26].Value?.ToString().Trim(),
                        worksheet.Cells[i,27].Value?.ToString().Trim(),
                        worksheet.Cells[i,28].Value?.ToString().Trim(),
                        worksheet.Cells[i,29].Value?.ToString().Trim(),
                        worksheet.Cells[i,30].Value?.ToString().Trim(),
                        worksheet.Cells[i,31].Value?.ToString().Trim(),
                        worksheet.Cells[i,32].Value?.ToString().Trim(),
                        worksheet.Cells[i,33].Value?.ToString().Trim(),
                        worksheet.Cells[i,34].Value?.ToString().Trim(),
                        worksheet.Cells[i,35].Value?.ToString().Trim(),
                        worksheet.Cells[i,36].Value?.ToString().Trim(),
                        worksheet.Cells[i,37].Value?.ToString().Trim(),
                        worksheet.Cells[i,38].Value?.ToString().Trim()
                    };
                }
            }
        }
        bool linkOpen = false;
        [DynamicData(nameof(outdoorExcelData), DynamicDataSourceType.Method)]
        [TestMethod]
        public void OutdoorForm(string uc, string locality, string houseCheckNo, string housePositive, string rdSideBushes, string rdSideBushesP,
            string parkBushes, string parkBushesP, string tire, string tireP, string treeHole, string treeHoleP, string garbage, string garbageP,
            string constructionWaste, string constructionWasteP, string flowerPot, string flowerPotP, string bird, string birdP, string cooler,
            string coolerP, string airConditioner, string airConditionerP, string tap, string tapP, string waterTank, string waterTankP, string openGutter,
            string openGutterP, string stagnantWater, string stagnantWaterP, string waterPot, string waterPotP, string fountain, string fountainP,
            string ceramic, string ceramicP)
        {
            OutdoorEntry entry = new OutdoorEntry();
            if (!linkOpen)
            {
                entry.browserSelection(browser);
                entry.landingPage(url);
                entry.LoginForm(uName, uPassword);
                linkOpen = true;
            }
            if (linkOpen)
            {


                entry.OutdoorDataEntryForm(uc, locality, houseCheckNo, housePositive, rdSideBushes, rdSideBushesP, parkBushes, parkBushesP, tire, tireP,
                 treeHole, treeHoleP, garbage, garbageP, constructionWaste, constructionWasteP, flowerPot, flowerPotP, bird, birdP, cooler, coolerP,
                 airConditioner, airConditionerP, tap, tapP, waterTank, waterTankP, openGutter, openGutterP, stagnantWater, stagnantWaterP, waterPot,
                 waterPotP, fountain, fountainP, ceramic, ceramicP);
            }
            entry.closeBrowser();
        }

        public static IEnumerable<object[]> covidExcelData()
        {
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo("COVIDFILE.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                var rowCount = worksheet.Dimension.End.Row;
                for (int i = 1; i <= rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                        worksheet.Cells[i,3].Value?.ToString().Trim(),
                        worksheet.Cells[i,4].Value?.ToString().Trim(),
                        worksheet.Cells[i,5].Value?.ToString().Trim(),
                        worksheet.Cells[i,6].Value?.ToString().Trim(),
                        worksheet.Cells[i,7].Value?.ToString().Trim()
                    };
                }
            }
        }
        [DynamicData(nameof(covidExcelData), DynamicDataSourceType.Method)]
        //[TestMethod]
        public void CovidDataEntryForm(string CNIC, string name, string fName, string gender, string age, 
            string contact, string address)
        {
            CovidDataEntry obj = new CovidDataEntry();
            obj.browserSelection(browser);
            obj.landingPage("https://covid-19.pshealthpunjab.gov.pk/patients");
            string userName = "ceo.sheikhupura@hisdu.com";
            string userPassword = "ceodhaskp";
            obj.loginActivity(userName, userPassword);
            obj.DataEntry(CNIC, name, fName, gender, age, contact, address);
        }
        public void csvData()
        {
            #region InputData
            string uc = TestContext.DataRow["uc"].ToString();
            string locality = TestContext.DataRow["uc"].ToString();
            string houseCheckNo = TestContext.DataRow["uc"].ToString();
            string housePositive = TestContext.DataRow["uc"].ToString();
            string airCondition = TestContext.DataRow["uc"].ToString();
            string airConditionP = TestContext.DataRow["uc"].ToString();
            string tap = TestContext.DataRow["uc"].ToString();
            string tapP = TestContext.DataRow["uc"].ToString();
            string tire = TestContext.DataRow["uc"].ToString();
            string tireP = TestContext.DataRow["uc"].ToString();
            string flower = TestContext.DataRow["uc"].ToString();
            string flowerP = TestContext.DataRow["uc"].ToString();
            string drink = TestContext.DataRow["uc"].ToString();
            string drinkP = TestContext.DataRow["uc"].ToString();
            string wash = TestContext.DataRow["uc"].ToString();
            string washP = TestContext.DataRow["uc"].ToString();
            string debris = TestContext.DataRow["uc"].ToString();
            string debrisP = TestContext.DataRow["uc"].ToString();
            string bird = TestContext.DataRow["uc"].ToString();
            string birdP = TestContext.DataRow["uc"].ToString();
            string animal = TestContext.DataRow["uc"].ToString();
            string animalP = TestContext.DataRow["uc"].ToString();
            string garbage = TestContext.DataRow["uc"].ToString();
            string garbageP = TestContext.DataRow["uc"].ToString();
            string waterTank = TestContext.DataRow["uc"].ToString();
            string waterTankP = TestContext.DataRow["uc"].ToString();
            string refrigrate = TestContext.DataRow["uc"].ToString();
            string refrigrateP = TestContext.DataRow["uc"].ToString();
            string mainHole = TestContext.DataRow["uc"].ToString();
            string mainHoleP = TestContext.DataRow["uc"].ToString();
            string houseWater = TestContext.DataRow["uc"].ToString();
            string houseWaterP = TestContext.DataRow["uc"].ToString();
            string cooler = TestContext.DataRow["uc"].ToString();
            string coolerP = TestContext.DataRow["uc"].ToString();
            string houseJunk = TestContext.DataRow["uc"].ToString();
            string houseJunkP = TestContext.DataRow["uc"].ToString();
            string roofJunk = TestContext.DataRow["uc"].ToString();
            string roofJunkP = TestContext.DataRow["uc"].ToString();
            #endregion
        }

        public void EPIMonthlyReportDataEntry()
        {

        }
    }

}
