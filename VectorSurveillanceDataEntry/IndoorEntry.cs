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
    public class IndoorEntry: GeneralMethodsClass
    {
        #region loginPage
        By userNameField = By.XPath("//input[@name='username']");
        By passwordField = By.XPath("//input[@name='password']");
        By loginButton = By.XPath("//button[@id='login-button']");
        By indoorFormButton = By.XPath("//a[@href='https://vectorsurveillance.punjab.gov.pk/indoor/add']");
        By selectUC = By.XPath("//select[@name='patientUC']");
        By localityField = By.XPath("//input[@id='locality']");
        By housesCheckField = By.XPath("//input[@name='houses_check']");
        By housesPositiveField = By.XPath("//input[@name='houses_positive']");
        By airConditionerField = By.XPath("//input[@name='ac_water_check']");
        By airConditionerPositiveField = By.XPath("//input[@name='ac_water_positive']");
        By waterTapField = By.XPath("//input[@name='leak_water_check']");
        By waterTapPositiveField = By.XPath("//input[@name='leak_water_positive']");
        By tireField = By.XPath("//input[@name='tire_check']");
        By tirePositiveField = By.XPath("//input[@name='tire_positive']");
        By flowerPotField = By.XPath("//input[@name='vases_check']");
        By flowerPotPositiveField = By.XPath("//input[@name='vases_positive']");
        By drinkPotField = By.XPath("//input[@name='drinking_check']");
        By drinkPotPositiveField = By.XPath("//input[@name='drinking_positive']");
        By washingPotField = By.XPath("//input[@name='washing_check']");
        By washingPotPositiveField = By.XPath("//input[@name='washing_positive']");
        By debrisField = By.XPath("//input[@name='construction_debris_check']");
        By debrisPositiveField = By.XPath("//input[@name='construction_debris_positive']");
        By birdField = By.XPath("//input[@name='pot_check']");
        By birdPositiveField = By.XPath("//input[@name='pot_positive']");
        By animalPotField = By.XPath("//input[@name='animal_pot_check']");
        By animalPotPositiveField = By.XPath("//input[@name='animal_pot_positive']");
        By garbageField = By.XPath("//input[@name='garbage_check']");
        By garbagePositiveField = By.XPath("//input[@name='garbage_positive']");
        By waterTankField = By.XPath("//input[@name='water_check']");
        By waterTankPositiveField = By.XPath("//input[@name='water_positive']");
        By refrigrateField = By.XPath("//input[@name='tray_check']");
        By refrigratePositiveField = By.XPath("//input[@name='tray_positive']");
        By mainHoleField = By.XPath("//input[@name='cover_check']");
        By mainHolePositiveField = By.XPath("//input[@name='cover_positive']");
        By houseWaterField = By.XPath("//input[@name='roof_check']");
        By houseWaterPositiveField = By.XPath("//input[@name='roof_positive']");
        By coolerField = By.XPath("//input[@name='cooler_check']");
        By coolerPositiveField = By.XPath("//input[@name='cooler_positive']");
        By houseJunkField = By.XPath("//input[@name='household_check']");
        By houseJunkPositiveField = By.XPath("//input[@name='household_positive']");
        By roofJunkField = By.XPath("//input[@name='roof_top_junk_check']");
        By roofJunkPositiveField = By.XPath("//input[@name='roof_top_junk_positive']");
        By submitButton = By.XPath("//input[@name='submit']");

        #endregion

        public void LoginForm(string uName, string uPass)
        {
            inputText(userNameField, uName);
            inputText(passwordField, uPass);   
            ClickableItem(loginButton);
            ClickableItem(indoorFormButton);
        }
        public void IndoorDataEntryForm(string uc, string locality, string houseCheckNo, string housePositive,
            string airCondition, string airConditionP, string tap, string tapP, string tire, string tireP,
            string flower, string flowerP, string drink, string drinkP, string wash, string washP, string debris, 
            string debrisP, string bird, string birdP, string animal, string animalP, string garbage, string garbageP,
            string waterTank, string waterTankP, string refrigrate, string refrigrateP, string mainHole, string mainHoleP,
            string houseWater, string houseWaterP, string cooler, string coolerP, string houseJunk, string houseJunkP, 
            string roofJunk, string roofJunkP)
        {
            dropDownItemSelect(selectUC, uc);
            inputText(localityField, locality);
            inputText(housesCheckField, houseCheckNo);
            inputText(housesPositiveField, housePositive);
            inputText(airConditionerField, airCondition);
            inputText(airConditionerPositiveField, airConditionP);
            inputText(waterTapField, tap);
            inputText(waterTapPositiveField, tapP);
            inputText(tireField, tire);
            inputText(tirePositiveField, tireP);
            inputText(flowerPotField, flower);
            inputText(flowerPotPositiveField, flowerP);
            inputText(drinkPotField, drink);
            inputText(drinkPotPositiveField, drinkP);
            inputText(washingPotField, wash);
            inputText(washingPotPositiveField, washP);
            inputText(debrisField, debris);
            inputText(debrisPositiveField, debrisP);
            inputText(birdField, bird);
            inputText(birdPositiveField, birdP);
            inputText(animalPotField, animal);
            inputText(animalPotPositiveField, animalP);
            inputText(garbageField, garbage);
            inputText(garbagePositiveField, garbageP);
            inputText(waterTankField, waterTank);
            inputText(waterTankPositiveField, waterTankP);
            inputText(refrigrateField, refrigrate);
            inputText(refrigratePositiveField, refrigrateP);
            inputText(mainHoleField, mainHole);
            inputText(mainHolePositiveField, mainHoleP);
            inputText(houseWaterField, houseWater);
            inputText(houseWaterPositiveField, houseWaterP);
            inputText(coolerField, cooler);
            inputText(coolerPositiveField, coolerP);
            inputText(houseJunkField, houseJunk);
            inputText(houseJunkPositiveField, houseJunkP);
            inputText(roofJunkField, roofJunk);
            inputText(roofJunkPositiveField, roofJunkP);
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
