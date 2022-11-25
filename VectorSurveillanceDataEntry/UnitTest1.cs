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
using static System.Net.WebRequestMethods;

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

        public static IEnumerable<object[]> EPIMonthlyReport()
        {
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo("EPIReportUpload.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                var rowCount = worksheet.Dimension.End.Row;
                for (int i = 1; i <= rowCount; i++)
                {
                    yield return new object[]
                    {
                        #region readData
                        worksheet.Cells[i, 1].Value?.ToString().Trim(),
                        worksheet.Cells[i, 2].Value?.ToString().Trim(),
                        worksheet.Cells[i, 3].Value?.ToString().Trim(),
                        worksheet.Cells[i, 4].Value?.ToString().Trim(),
                        worksheet.Cells[i, 5].Value?.ToString().Trim(),
                        worksheet.Cells[i, 6].Value?.ToString().Trim(),
                        worksheet.Cells[i, 7].Value?.ToString().Trim(),
                        worksheet.Cells[i, 8].Value?.ToString().Trim(),
                        worksheet.Cells[i, 9].Value?.ToString().Trim(),
                        worksheet.Cells[i, 10].Value?.ToString().Trim(),
                        worksheet.Cells[i, 11].Value?.ToString().Trim(),
                        worksheet.Cells[i, 12].Value?.ToString().Trim(),
                        worksheet.Cells[i, 13].Value?.ToString().Trim(),
                        worksheet.Cells[i, 14].Value?.ToString().Trim(),
                        worksheet.Cells[i, 15].Value?.ToString().Trim(),
                        worksheet.Cells[i, 16].Value?.ToString().Trim(),
                        worksheet.Cells[i, 17].Value?.ToString().Trim(),
                        worksheet.Cells[i, 18].Value?.ToString().Trim(),
                        worksheet.Cells[i, 19].Value?.ToString().Trim(),
                        worksheet.Cells[i, 20].Value?.ToString().Trim(),
                        worksheet.Cells[i, 21].Value?.ToString().Trim(),
                        worksheet.Cells[i, 22].Value?.ToString().Trim(),
                        worksheet.Cells[i, 23].Value?.ToString().Trim(),
                        worksheet.Cells[i, 24].Value?.ToString().Trim(),
                        worksheet.Cells[i, 25].Value?.ToString().Trim(),
                        worksheet.Cells[i, 26].Value?.ToString().Trim(),
                        worksheet.Cells[i, 27].Value?.ToString().Trim(),
                        worksheet.Cells[i, 28].Value?.ToString().Trim(),
                        worksheet.Cells[i, 29].Value?.ToString().Trim(),
                        worksheet.Cells[i, 30].Value?.ToString().Trim(),
                        worksheet.Cells[i, 31].Value?.ToString().Trim(),
                        worksheet.Cells[i, 32].Value?.ToString().Trim(),
                        worksheet.Cells[i, 33].Value?.ToString().Trim(),
                        worksheet.Cells[i, 34].Value?.ToString().Trim(),
                        worksheet.Cells[i, 35].Value?.ToString().Trim(),
                        worksheet.Cells[i, 36].Value?.ToString().Trim(),
                        worksheet.Cells[i, 37].Value?.ToString().Trim(),
                        worksheet.Cells[i, 38].Value?.ToString().Trim(),
                        worksheet.Cells[i, 39].Value?.ToString().Trim(),
                        worksheet.Cells[i, 40].Value?.ToString().Trim(),
                        worksheet.Cells[i, 41].Value?.ToString().Trim(),
                        worksheet.Cells[i, 42].Value?.ToString().Trim(),
                        worksheet.Cells[i, 43].Value?.ToString().Trim(),
                        worksheet.Cells[i, 44].Value?.ToString().Trim(),
                        worksheet.Cells[i, 45].Value?.ToString().Trim(),
                        worksheet.Cells[i, 46].Value?.ToString().Trim(),
                        worksheet.Cells[i, 47].Value?.ToString().Trim(),
                        worksheet.Cells[i, 48].Value?.ToString().Trim(),
                        worksheet.Cells[i, 49].Value?.ToString().Trim(),
                        worksheet.Cells[i, 50].Value?.ToString().Trim(),
                        worksheet.Cells[i, 51].Value?.ToString().Trim(),
                        worksheet.Cells[i, 52].Value?.ToString().Trim(),
                        worksheet.Cells[i, 53].Value?.ToString().Trim(),
                        worksheet.Cells[i, 54].Value?.ToString().Trim(),
                        worksheet.Cells[i, 55].Value?.ToString().Trim(),
                        worksheet.Cells[i, 56].Value?.ToString().Trim(),
                        worksheet.Cells[i, 57].Value?.ToString().Trim(),
                        worksheet.Cells[i, 58].Value?.ToString().Trim(),
                        worksheet.Cells[i, 59].Value?.ToString().Trim(),
                        worksheet.Cells[i, 60].Value?.ToString().Trim(),
                        worksheet.Cells[i, 61].Value?.ToString().Trim(),
                        worksheet.Cells[i, 62].Value?.ToString().Trim(),
                        worksheet.Cells[i, 63].Value?.ToString().Trim(),
                        worksheet.Cells[i, 64].Value?.ToString().Trim(),
                        worksheet.Cells[i, 65].Value?.ToString().Trim(),
                        worksheet.Cells[i, 66].Value?.ToString().Trim(),
                        worksheet.Cells[i, 67].Value?.ToString().Trim(),
                        worksheet.Cells[i, 68].Value?.ToString().Trim(),
                        worksheet.Cells[i, 69].Value?.ToString().Trim(),
                        worksheet.Cells[i, 70].Value?.ToString().Trim(),
                        worksheet.Cells[i, 71].Value?.ToString().Trim(),
                        worksheet.Cells[i, 72].Value?.ToString().Trim(),
                        worksheet.Cells[i, 73].Value?.ToString().Trim(),
                        worksheet.Cells[i, 74].Value?.ToString().Trim(),
                        worksheet.Cells[i, 75].Value?.ToString().Trim(),
                        worksheet.Cells[i, 76].Value?.ToString().Trim(),
                        worksheet.Cells[i, 77].Value?.ToString().Trim(),
                        worksheet.Cells[i, 78].Value?.ToString().Trim(),
                        worksheet.Cells[i, 79].Value?.ToString().Trim(),
                        worksheet.Cells[i, 80].Value?.ToString().Trim(),
                        worksheet.Cells[i, 81].Value?.ToString().Trim(),
                        worksheet.Cells[i, 82].Value?.ToString().Trim(),
                        worksheet.Cells[i, 83].Value?.ToString().Trim(),
                        worksheet.Cells[i, 84].Value?.ToString().Trim(),
                        worksheet.Cells[i, 85].Value?.ToString().Trim(),
                        worksheet.Cells[i, 86].Value?.ToString().Trim(),
                        worksheet.Cells[i, 87].Value?.ToString().Trim(),
                        worksheet.Cells[i, 88].Value?.ToString().Trim(),
                        worksheet.Cells[i, 89].Value?.ToString().Trim(),
                        worksheet.Cells[i, 90].Value?.ToString().Trim(),
                        worksheet.Cells[i, 91].Value?.ToString().Trim(),
                        worksheet.Cells[i, 92].Value?.ToString().Trim(),
                        worksheet.Cells[i, 93].Value?.ToString().Trim(),
                        worksheet.Cells[i, 94].Value?.ToString().Trim(),
                        worksheet.Cells[i, 95].Value?.ToString().Trim(),
                        worksheet.Cells[i, 96].Value?.ToString().Trim(),
                        worksheet.Cells[i, 97].Value?.ToString().Trim(),
                        worksheet.Cells[i, 98].Value?.ToString().Trim(),
                        worksheet.Cells[i, 99].Value?.ToString().Trim(),
                        worksheet.Cells[i, 100].Value?.ToString().Trim(),
                        worksheet.Cells[i, 101].Value?.ToString().Trim(),
                        worksheet.Cells[i, 102].Value?.ToString().Trim(),
                        worksheet.Cells[i, 103].Value?.ToString().Trim(),
                        worksheet.Cells[i, 104].Value?.ToString().Trim(),
                        worksheet.Cells[i, 105].Value?.ToString().Trim(),
                        worksheet.Cells[i, 106].Value?.ToString().Trim(),
                        worksheet.Cells[i, 107].Value?.ToString().Trim(),
                        worksheet.Cells[i, 108].Value?.ToString().Trim(),
                        worksheet.Cells[i, 109].Value?.ToString().Trim(),
                        worksheet.Cells[i, 110].Value?.ToString().Trim(),
                        worksheet.Cells[i, 111].Value?.ToString().Trim(),
                        worksheet.Cells[i, 112].Value?.ToString().Trim(),
                        worksheet.Cells[i, 113].Value?.ToString().Trim(),
                        worksheet.Cells[i, 114].Value?.ToString().Trim(),
                        worksheet.Cells[i, 115].Value?.ToString().Trim(),
                        worksheet.Cells[i, 116].Value?.ToString().Trim(),
                        worksheet.Cells[i, 117].Value?.ToString().Trim(),
                        worksheet.Cells[i, 118].Value?.ToString().Trim(),
                        worksheet.Cells[i, 119].Value?.ToString().Trim(),
                        worksheet.Cells[i, 120].Value?.ToString().Trim(),
                        worksheet.Cells[i, 121].Value?.ToString().Trim(),
                        worksheet.Cells[i, 122].Value?.ToString().Trim(),
                        worksheet.Cells[i, 123].Value?.ToString().Trim(),
                        worksheet.Cells[i, 124].Value?.ToString().Trim(),
                        worksheet.Cells[i, 125].Value?.ToString().Trim(),
                        worksheet.Cells[i, 126].Value?.ToString().Trim(),
                        worksheet.Cells[i, 127].Value?.ToString().Trim(),
                        worksheet.Cells[i, 128].Value?.ToString().Trim(),
                        worksheet.Cells[i, 129].Value?.ToString().Trim(),
                        worksheet.Cells[i, 130].Value?.ToString().Trim(),
                        worksheet.Cells[i, 131].Value?.ToString().Trim(),
                        worksheet.Cells[i, 132].Value?.ToString().Trim(),
                        worksheet.Cells[i, 133].Value?.ToString().Trim(),
                        worksheet.Cells[i, 134].Value?.ToString().Trim(),
                        worksheet.Cells[i, 135].Value?.ToString().Trim(),
                        worksheet.Cells[i, 136].Value?.ToString().Trim(),
                        worksheet.Cells[i, 137].Value?.ToString().Trim(),
                        worksheet.Cells[i, 138].Value?.ToString().Trim(),
                        worksheet.Cells[i, 139].Value?.ToString().Trim(),
                        worksheet.Cells[i, 140].Value?.ToString().Trim(),
                        worksheet.Cells[i, 141].Value?.ToString().Trim(),
                        worksheet.Cells[i, 142].Value?.ToString().Trim(),
                        worksheet.Cells[i, 143].Value?.ToString().Trim(),
                        worksheet.Cells[i, 144].Value?.ToString().Trim(),
                        worksheet.Cells[i, 145].Value?.ToString().Trim(),
                        worksheet.Cells[i, 146].Value?.ToString().Trim(),
                        worksheet.Cells[i, 147].Value?.ToString().Trim(),
                        worksheet.Cells[i, 148].Value?.ToString().Trim(),
                        worksheet.Cells[i, 149].Value?.ToString().Trim(),
                        worksheet.Cells[i, 150].Value?.ToString().Trim(),
                        worksheet.Cells[i, 151].Value?.ToString().Trim(),
                        worksheet.Cells[i, 152].Value?.ToString().Trim(),
                        worksheet.Cells[i, 153].Value?.ToString().Trim(),
                        worksheet.Cells[i, 154].Value?.ToString().Trim(),
                        worksheet.Cells[i, 155].Value?.ToString().Trim(),
                        worksheet.Cells[i, 156].Value?.ToString().Trim(),
                        worksheet.Cells[i, 157].Value?.ToString().Trim(),
                        worksheet.Cells[i, 158].Value?.ToString().Trim(),
                        worksheet.Cells[i, 159].Value?.ToString().Trim(),
                        worksheet.Cells[i, 160].Value?.ToString().Trim(),
                        worksheet.Cells[i, 161].Value?.ToString().Trim(),
                        worksheet.Cells[i, 162].Value?.ToString().Trim(),
                        worksheet.Cells[i, 163].Value?.ToString().Trim(),
                        worksheet.Cells[i, 164].Value?.ToString().Trim(),
                        worksheet.Cells[i, 165].Value?.ToString().Trim(),
                        worksheet.Cells[i, 166].Value?.ToString().Trim(),
                        worksheet.Cells[i, 167].Value?.ToString().Trim(),
                        worksheet.Cells[i, 168].Value?.ToString().Trim(),
                        worksheet.Cells[i, 169].Value?.ToString().Trim(),
                        worksheet.Cells[i, 170].Value?.ToString().Trim(),
                        worksheet.Cells[i, 171].Value?.ToString().Trim(),
                        worksheet.Cells[i, 172].Value?.ToString().Trim(),
                        worksheet.Cells[i, 173].Value?.ToString().Trim(),
                        worksheet.Cells[i, 174].Value?.ToString().Trim(),
                        worksheet.Cells[i, 175].Value?.ToString().Trim(),
                        worksheet.Cells[i, 176].Value?.ToString().Trim(),
                        worksheet.Cells[i, 177].Value?.ToString().Trim(),
                        worksheet.Cells[i, 178].Value?.ToString().Trim(),
                        worksheet.Cells[i, 179].Value?.ToString().Trim(),
                        worksheet.Cells[i, 180].Value?.ToString().Trim(),
                        worksheet.Cells[i, 181].Value?.ToString().Trim(),
                        worksheet.Cells[i, 182].Value?.ToString().Trim(),
                        worksheet.Cells[i, 183].Value?.ToString().Trim(),
                        worksheet.Cells[i, 184].Value?.ToString().Trim(),
                        worksheet.Cells[i, 185].Value?.ToString().Trim(),
                        worksheet.Cells[i, 186].Value?.ToString().Trim(),
                        worksheet.Cells[i, 187].Value?.ToString().Trim(),
                        worksheet.Cells[i, 188].Value?.ToString().Trim(),
                        worksheet.Cells[i, 189].Value?.ToString().Trim(),
                        worksheet.Cells[i, 190].Value?.ToString().Trim(),
                        worksheet.Cells[i, 191].Value?.ToString().Trim(),
                        worksheet.Cells[i, 192].Value?.ToString().Trim(),
                        worksheet.Cells[i, 193].Value?.ToString().Trim(),
                        worksheet.Cells[i, 194].Value?.ToString().Trim(),
                        worksheet.Cells[i, 195].Value?.ToString().Trim(),
                        worksheet.Cells[i, 196].Value?.ToString().Trim(),
                        worksheet.Cells[i, 197].Value?.ToString().Trim(),
                        worksheet.Cells[i, 198].Value?.ToString().Trim(),
                        worksheet.Cells[i, 199].Value?.ToString().Trim(),
                        worksheet.Cells[i, 200].Value?.ToString().Trim(),
                        worksheet.Cells[i, 201].Value?.ToString().Trim(),
                        worksheet.Cells[i, 202].Value?.ToString().Trim(),
                        worksheet.Cells[i, 203].Value?.ToString().Trim(),
                        worksheet.Cells[i, 204].Value?.ToString().Trim(),
                        worksheet.Cells[i, 205].Value?.ToString().Trim(),
                        worksheet.Cells[i, 206].Value?.ToString().Trim(),
                        worksheet.Cells[i, 207].Value?.ToString().Trim(),
                        worksheet.Cells[i, 208].Value?.ToString().Trim(),
                        worksheet.Cells[i, 209].Value?.ToString().Trim(),
                        worksheet.Cells[i, 210].Value?.ToString().Trim(),
                        worksheet.Cells[i, 211].Value?.ToString().Trim(),
                        worksheet.Cells[i, 212].Value?.ToString().Trim(),
                        worksheet.Cells[i, 213].Value?.ToString().Trim(),
                        worksheet.Cells[i, 214].Value?.ToString().Trim(),
                        worksheet.Cells[i, 215].Value?.ToString().Trim(),
                        worksheet.Cells[i, 216].Value?.ToString().Trim(),
                        worksheet.Cells[i, 217].Value?.ToString().Trim(),
                        worksheet.Cells[i, 218].Value?.ToString().Trim(),
                        worksheet.Cells[i, 219].Value?.ToString().Trim(),
                        worksheet.Cells[i, 220].Value?.ToString().Trim(),
                        worksheet.Cells[i, 221].Value?.ToString().Trim(),
                        worksheet.Cells[i, 222].Value?.ToString().Trim(),
                        worksheet.Cells[i, 223].Value?.ToString().Trim(),
                        worksheet.Cells[i, 224].Value?.ToString().Trim(),
                        worksheet.Cells[i, 225].Value?.ToString().Trim(),
                        worksheet.Cells[i, 226].Value?.ToString().Trim(),
                        worksheet.Cells[i, 227].Value?.ToString().Trim(),
                        worksheet.Cells[i, 228].Value?.ToString().Trim(),
                        worksheet.Cells[i, 229].Value?.ToString().Trim(),
                        worksheet.Cells[i, 230].Value?.ToString().Trim(),
                        worksheet.Cells[i, 231].Value?.ToString().Trim(),
                        worksheet.Cells[i, 232].Value?.ToString().Trim(),
                        worksheet.Cells[i, 233].Value?.ToString().Trim(),
                        worksheet.Cells[i, 234].Value?.ToString().Trim(),
                        worksheet.Cells[i, 235].Value?.ToString().Trim()
                        #endregion
                    };
                }
            }
        }
        [DynamicData(nameof(EPIMonthlyReport), DynamicDataSourceType.Method)]
        //[TestMethod]
        public void EPIMonthlyReportDataEntry(string UC, string tFixedCenters, string tFunctioningCenters, string reportingCenters, string planned, string sessionsHeld,
            string outSessionsPlaned, string outSessionsHeld, string hepF1, string hepF2, string hepO1, string hepO2, string bcg01, string bcg02, string bcg03, string bcg04,
            string opv00, string opv01, string opv02, string opv03, string opv10, string opv11, string opv12, string opv13, string opv20, string opv21, string opv22, string opv23,
            string opv30, string opv31, string opv32, string opv33, string FPenta11, string FPenta12, string OPenta13, string OPenta14, string FPenta21, string FPenta22,
            string OPenta23, string OPenta24, string FPenta31, string FPenta32, string OPenta33, string OPenta34, string FPCV11, string FPCV12, string OPCV13, string OPCV14,
            string FPCV21, string FPCV22, string OPCV23, string OPCV24, string FPCV31, string FPCV32, string OPCV33, string OPCV34, string FRota11, string FRota12, string ORota13,
            string ORota14, string FRota21, string FRota22, string ORota23, string ORota24, string FIPV11, string FIPV12, string OIPV13, string OIPV14, string FIPV21, string FIPV22,
            string OIPV23, string OIPV24, string FTCV1, string FTCV2, string OTCV3, string OTCV4, string FMR11, string FMR12, string OMR13, string OMR14, string FMR21, string FMR22,
            string OMR23, string OMR24, string FDTP1, string FDTP2, string ODTP3, string ODTP4, string HEBSegDue1, string HEBSegDef2, string HEBSegUCOut3, string HEBSegDisOut4,
            string BCGSegDue1, string BCGSegDef2, string BCGSegUCOut3, string BCGSegDisOut4, string OPV0SegDue1, string OPV0SegDef2, string OPV0SegUCOut3, string OPV0SegDisOut4,
            string OPV1SegDue1, string OPV1SegDef2, string OPV1SegUCOut3, string OPV1SegDisOut4, string OPV2SegDue1, string OPV2SegDef2, string OPV2SegUCOut3, string OPV2egDisOut4,
            string OPV3SegDue1, string OPV3SegDef2, string OPV3SegUCOut3, string OPV3SegDisOut4, string Penta1SegDue11, string Penta1SegDef12, string Penta1SegUCOut13,
            string Penta1SegDisOut14, string Penta2SegDue21, string Penta2SegDef22, string Penta2SegUCOut23, string Penta2SegDisOut24, string Penta3SegDue31, string Penta3SegDef32,
            string Penta3SegUCOut33, string Penta3SegDisOut34, string Rota1SegDue11, string Rota1SegDef12, string Rota1SegUCOut13, string Rota1SegDisOut14, string Rota2SegDue21,
            string Rota2SegDef22, string Rota2SegUCOut23, string Rota2SegDisOut24, string IPV1SegDue11, string IPV1SegDef12, string IPV1SegUCOut13, string IPV1SegDisOut14,
            string IPV2SegDue21, string IPV2SegDef22, string IPV2SegUCOut22, string IPV2SegDisOut24, string TCVSegDue1, string TCVSegDef2, string TCVSegUCOut3, string TCVSegDisOut4,
            string MR1SegDue11, string MR1SegDef12, string MR1SegUCOut13, string MR1SegDisOut14, string MR2SegDue21, string MR2SegDef22, string MR2SegUCOut23, string MR2SegDisOut24,
            string DTPSegDue1, string DTPSegDef2, string DTPSegUCOut3, string DTPSegDisOut4, string FTT11, string FTT22, string FTT33, string FTT44, string FTT55, string OTT11,
            string OTT22, string OTT33, string OTT44, string OTT55, string HEPBOpenBalance1, string HEPBReceived2, string HEPBCloseBalance3, string BCGOpenBalance1, string BCGReceived2,
            string BCGCloseBalance3, string OPVOpenBalance1, string OPVReceived2, string OPVCloseBalance3, string PentaOpenBalance1, string PentaReceived2, string PentaCloseBalance3,
            string PCVOpenBalance1, string PCVReceived2, string PCVCloseBalance3, string ROTAOpenBalacne1, string ROTAReceived2, string ROTACloseBalance3, string IPVOpenBalacne1,
            string IPVReceived2, string IPVCloseBalance3, string TCVOpenBalacne1, string TCVReceived2, string TCVCloseBalance3, string MROpenBalacne1, string MRReceived2, string MRCloseBalance3,
            string DTPOpenBalacne1, string DTPReceived2, string DTPCloseBalance3, string TTOpenBalacne1, string TTReceived2, string TTCloseBalance3, string AD005OpenBalacne1,
            string AD005Received2, string AD005CloseBalance3, string AD05OpenBalacne1, string AD05Received2, string AD05CloseBalance3, string Syr2OpenBalacne1, string Syr2Received2,
            string Syr2CloseBalance3, string Syr5OpenBalacne1, string Syr5Received2, string Syr5CloseBalance3, string CardChildOpenBalacne1, string CardChildReceived2,
            string CardChildCloseBalance3, string CardWomenOpenBalacne1, string CardWomenReceived2, string CardWomenCloseBalance3, string UsedSafteyBox1, string DisposedSafteyBox2,
            string methodDisposeDropDown11, string UsedVials1, string DisposedVials2, string methodDisposeDropDown21, string UnusedVials1, string UnusedDisposedVials2, string methodDisposeDropDown31,
            string MissedNA1, string Refusal2, string Sick1, string newMaleChild1, string newFemaleChild2, string recordedZeroDue1, string recordedZeroDefaulter2, string coverageZeroDue1,
            string coverageZeroDefaulter2)
        {
            EPIMonthlyReport obj = new EPIMonthlyReport();

            obj.browserSelection(browser);
            obj.landingPage("https://punjab.epimis.pk/");
            Thread.Sleep(10000);
            obj.Login();
           // Thread.Sleep(10000);
            obj.EntryPage(UC, tFixedCenters, tFunctioningCenters, reportingCenters, planned, sessionsHeld, outSessionsPlaned, outSessionsHeld, hepF1, hepF2, hepO1, hepO2, bcg01,
                bcg02, bcg03, bcg04, opv00, opv01, opv02, opv03, opv10, opv11, opv12, opv13, opv20, opv21, opv22, opv23, opv30, opv31, opv32, opv33, FPenta11, FPenta12, OPenta13,
                OPenta14, FPenta21, FPenta22, OPenta23, OPenta24, FPenta31, FPenta32, OPenta33, OPenta34, FPCV11, FPCV12, OPCV13, OPCV14, FPCV21, FPCV22, OPCV23, OPCV24, FPCV31, 
                FPCV32, OPCV33, OPCV34, FRota11, FRota12, ORota13, ORota14, FRota21, FRota22, ORota23, ORota24, FIPV11, FIPV12, OIPV13, OIPV14, FIPV21, FIPV22, OIPV23, OIPV24, FTCV1,
                FTCV2, OTCV3, OTCV4, FMR11, FMR12, OMR13, OMR14, FMR21, FMR22, OMR23, OMR24, FDTP1, FDTP2, ODTP3, ODTP4, HEBSegDue1, HEBSegDef2, HEBSegUCOut3, HEBSegDisOut4, BCGSegDue1, 
                BCGSegDef2, BCGSegUCOut3, BCGSegDisOut4, OPV0SegDue1, OPV0SegDef2, OPV0SegUCOut3, OPV0SegDisOut4, OPV1SegDue1, OPV1SegDef2, OPV1SegUCOut3, OPV1SegDisOut4, OPV2SegDue1, 
                OPV2SegDef2, OPV2SegUCOut3, OPV2egDisOut4, OPV3SegDue1, OPV3SegDef2, OPV3SegUCOut3, OPV3SegDisOut4, Penta1SegDue11, Penta1SegDef12, Penta1SegUCOut13, Penta1SegDisOut14, 
                Penta2SegDue21, Penta2SegDef22, Penta2SegUCOut23, Penta2SegDisOut24, Penta3SegDue31, Penta3SegDef32, Penta3SegUCOut33, Penta3SegDisOut34, Rota1SegDue11, Rota1SegDef12, 
                Rota1SegUCOut13, Rota1SegDisOut14, Rota2SegDue21, Rota2SegDef22, Rota2SegUCOut23, Rota2SegDisOut24, IPV1SegDue11, IPV1SegDef12, IPV1SegUCOut13, IPV1SegDisOut14,
                IPV2SegDue21, IPV2SegDef22, IPV2SegUCOut22, IPV2SegDisOut24, TCVSegDue1, TCVSegDef2, TCVSegUCOut3, TCVSegDisOut4, MR1SegDue11, MR1SegDef12, MR1SegUCOut13, 
                MR1SegDisOut14, MR2SegDue21, MR2SegDef22, MR2SegUCOut23, MR2SegDisOut24, DTPSegDue1, DTPSegDef2, DTPSegUCOut3, DTPSegDisOut4, FTT11, FTT22, FTT33, FTT44, FTT55, 
                OTT11, OTT22, OTT33, OTT44, OTT55, HEPBOpenBalance1, HEPBReceived2, HEPBCloseBalance3, BCGOpenBalance1, BCGReceived2, BCGCloseBalance3, OPVOpenBalance1, OPVReceived2,
                OPVCloseBalance3, PentaOpenBalance1, PentaReceived2, PentaCloseBalance3, PCVOpenBalance1, PCVReceived2, PCVCloseBalance3, ROTAOpenBalacne1, ROTAReceived2, 
                ROTACloseBalance3, IPVOpenBalacne1,IPVReceived2, IPVCloseBalance3, TCVOpenBalacne1, TCVReceived2, TCVCloseBalance3, MROpenBalacne1, MRReceived2, MRCloseBalance3,
                DTPOpenBalacne1, DTPReceived2, DTPCloseBalance3, TTOpenBalacne1, TTReceived2, TTCloseBalance3, AD005OpenBalacne1, AD005Received2, AD005CloseBalance3, AD05OpenBalacne1, 
                AD05Received2, AD05CloseBalance3, Syr2OpenBalacne1, Syr2Received2, Syr2CloseBalance3, Syr5OpenBalacne1, Syr5Received2, Syr5CloseBalance3, CardChildOpenBalacne1, 
                CardChildReceived2,CardChildCloseBalance3, CardWomenOpenBalacne1, CardWomenReceived2, CardWomenCloseBalance3, UsedSafteyBox1, DisposedSafteyBox2, methodDisposeDropDown11, 
                UsedVials1, DisposedVials2, methodDisposeDropDown21, UnusedVials1, UnusedDisposedVials2, methodDisposeDropDown31, MissedNA1, Refusal2, Sick1, newMaleChild1, 
                newFemaleChild2, recordedZeroDue1, recordedZeroDefaulter2, coverageZeroDue1, coverageZeroDefaulter2);
        }
    }

}
