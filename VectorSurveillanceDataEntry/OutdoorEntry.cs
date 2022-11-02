using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using FP_InterWood;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

namespace VectorSurveillanceDataEntry
{
    public class OutdoorEntry : GeneralMethodsClass
    {
        #region loginPage
        By userNameField = By.XPath("//input[@name='username']");
        By passwordField = By.XPath("//input[@name='password']");
        By loginButton = By.XPath("//button[@id='login-button']");
        By outdoorFormButton = By.XPath("//a[@href='https://vectorsurveillance.punjab.gov.pk/outdoor/add']");
        By selectUC = By.XPath("//select[@name='patientUC']");
        By localityField = By.XPath("//input[@id='locality']");
        By housesCheckField = By.XPath("//input[@name='houses_check']");
        By housesPositiveField = By.XPath("//input[@name='houses_positive']");
        By rdSideBushesField = By.XPath("//input[@name='rd_sd_bushes_check']");//1
        By rdSideBushesPositiveField = By.XPath("//input[@name='rd_sd_bushes_positive']");//2
        By parkBushesField = By.XPath("//input[@name='pk_bushes_check']"); //3
        By parkBushesPositiveField = By.XPath("//input[@name='pk_bushes_positive']");//4
        By tireField = By.XPath("//input[@name='tire_check']");//5
        By tirePositiveField = By.XPath("//input[@name='tire_positive']");//6
        By treeHoleField = By.XPath("//input[@name='tree_check']");//7
        By treeHolePositiveField = By.XPath("//input[@name='tree_positive']");//8
        By garbageField = By.XPath("//input[@name='garbage_check']");//9
        By garbagePositiveField = By.XPath("//input[@name='garbage_positive']");//10
        By constructionWasteField = By.XPath("//input[@name='contruction_check']");//11
        By constructionWastePositiveField = By.XPath("//input[@name='contruction_positive']");//12
        By flowerPotField = By.XPath("//input[@name='pot_check']");//13
        By flowerPotPositiveField = By.XPath("//input[@name='pot_positive']");//14
        By birdField = By.XPath("//input[@name='tray_check']");//15
        By birdPositiveField = By.XPath("//input[@name='tray_positive']");//16
        By coolerField = By.XPath("//input[@name='cooler_check']");//17
        By coolerPositiveField = By.XPath("//input[@name='cooler_positive']");//18
        By airConditionerField = By.XPath("//input[@name='ac_water_check']"); //19
        By airConditionerPositiveField = By.XPath("//input[@name='ac_water_positive']");//20
        By waterTapField = By.XPath("//input[@name='water_check']");//21
        By waterTapPositiveField = By.XPath("//input[@name='water_positive']");//22
        By waterTankField = By.XPath("//input[@name='tank_check']");//23
        By waterTankPositiveField = By.XPath("//input[@name='tank_positive']");//24
        By openGutterField = By.XPath("//input[@name='open_gutter_drain_check']");//25
        By openGutterPositiveField = By.XPath("//input[@name='open_gutter_drain_positive']");//26
        By stagnantWaterField = By.XPath("//input[@name='stagnant_water_check']");//27
        By stagnantWaterPositiveField = By.XPath("//input[@name='stagnant_water_positive']");//28
        By waterPotShopField = By.XPath("//input[@name='water_pots_puncture_shop_check']");//29
        By waterPotShopPositiveField = By.XPath("//input[@name='water_pots_puncture_shop_positive']");//30
        By fountainField = By.XPath("//input[@name='roadside_fountian_check']");//31
        By fountainPositiveField = By.XPath("//input[@name='roadside_fountian_positive']");//32
        By ceramicJunkField = By.XPath("//input[@name='ceramicwaste_check']");//33
        By ceramicJunkPositiveField = By.XPath("//input[@name='ceramicwaste_positive']");//34
        By submitButton = By.XPath("//input[@name='submit']");
        #endregion


        public void LoginForm(string uName, string uPass)
        {
            inputText(userNameField, uName);
            inputText(passwordField, uPass);
            ClickableItem(loginButton);
            ClickableItem(outdoorFormButton);
        }
        public void OutdoorDataEntryForm(string uc, string locality, string houseCheckNo, string housePositive, string rdSideBushes, string rdSideBushesP, 
            string parkBushes, string parkBushesP, string tire, string tireP, string treeHole, string treeHoleP, string garbage, string garbageP, 
            string constructionWaste, string constructionWasteP, string flowerPot, string flowerPotP, string bird, string birdP, string cooler, 
            string coolerP, string airConditioner, string airConditionerP, string tap, string tapP, string waterTank, string waterTankP, string openGutter, 
            string openGutterP, string stagnantWater, string stagnantWaterP, string waterPot, string waterPotP, string fountain, string fountainP,
            string ceramic, string ceramicP)
        {
            dropDownItemSelect(selectUC, uc);
            inputText(localityField, locality);
            inputText(housesCheckField, houseCheckNo);
            inputText(housesPositiveField, housePositive);
            inputText(rdSideBushesField, rdSideBushes);
            inputText(rdSideBushesPositiveField, rdSideBushesP);
            inputText(parkBushesField, parkBushes);
            inputText(parkBushesPositiveField, parkBushesP);
            inputText(tireField, tire);
            inputText(tirePositiveField, tireP);
            inputText(treeHoleField, treeHole);
            inputText(treeHolePositiveField, treeHoleP);
            inputText(garbageField, garbage);
            inputText(garbagePositiveField, garbageP);
            inputText(constructionWasteField, constructionWaste);
            inputText(constructionWastePositiveField, constructionWasteP);
            inputText(flowerPotField, flowerPot);
            inputText(flowerPotPositiveField, flowerPotP);
            inputText(birdField, bird);
            inputText(birdPositiveField, birdP);
            inputText(coolerField, cooler);
            inputText(coolerPositiveField, coolerP);
            inputText(airConditionerField, airConditioner);
            inputText(airConditionerPositiveField, airConditionerP);
            inputText(waterTapField, tap);
            inputText(waterTapPositiveField, tapP);
            inputText(waterTankField, waterTank);
            inputText(waterTankPositiveField, waterTankP);
            inputText(openGutterField, openGutter);
            inputText(openGutterPositiveField, openGutterP);
            inputText(stagnantWaterField, stagnantWater);
            inputText(stagnantWaterPositiveField, stagnantWaterP);
            inputText(waterPotShopField, waterPot);
            inputText(waterPotShopPositiveField, waterPotP);
            inputText(fountainField, fountain);
            inputText(fountainPositiveField, fountainP);
            inputText(ceramicJunkField, ceramic);
            inputText(ceramicJunkPositiveField, ceramicP);
            ClickableItem(submitButton);
            Thread.Sleep(1000);
            acceptAlert();
            Thread.Sleep(5000);
        }
        void temp()
        {

        }
    }
}
