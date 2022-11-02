using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
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
    public class CovidDataEntry:GeneralMethodsClass
    {
        By enterUserName = By.XPath("//input[@name = 'Email']");
        By enterPassword = By.XPath("//input[@name = 'Password']");
        By loginButton = By.XPath("//button[text()= 'Log In']");
        string customLink = "https://covid-19.pshealthpunjab.gov.pk/patients";

        By clusterField = By.XPath("//select[@name='ClusterType']");
        By mediaTypeField = By.XPath("//select[@name='MediaType']");
        By relationField = By.XPath("//select[@name='Relation']");
        By relationGuardianField = By.XPath("//select[@name='RelationWithGuradian']");
        By CNICField = By.XPath("//select[@name='CnicorPassport']");
        By CNICinputField = By.XPath("//input[@name ='CNIC']");
        By nameInputField = By.XPath("//input[@id='exampleInputName1']");
        By fNameInputField = By.XPath("//input[@id='exampleInputName3']");
        By selectGenderField = By.XPath("//select[@name='Gender']");
        By ageInputField = By.XPath("//input[@id='exampleInputName4']");
        By phone1InputField = By.XPath("//input[@id='exampleInputMoible6']");
        By phone2InputField = By.XPath("//input[@id='RelativeMobileNo']");
        By nationalityField = By.XPath("//select[@name='Nationality']");
        By divisionField = By.XPath("//select[@name='DivCode']");
        By districtField = By.XPath("//select[@name='DisCode']");
        By tehsilField = By.XPath("//select[@name='TehCode']");
        By addressInputField = By.XPath("//input[@id='exampleInputName10']");
        By patientConditionField = By.XPath("//select[@name='Condition']");
        By selectVaccineField = By.XPath("//select[@name='VaccinationStatus']");
        By selectVaccineTypeField = By.XPath("//select[@name='VaccinationType']");
        By submitButton = By.XPath("//*[@id=\"wrapper\"]/*/button[1]");

        public void loginActivity(string uName, string pass)
        {
            inputText(enterUserName, uName);
            inputText(enterPassword, pass);
            ClickableItem(loginButton);
            driver.Navigate().GoToUrl(customLink);
        }

        public void DataEntry(string CNIC, string name, string fName, string gender, string age, string contact,
            string address)
        {
            dropDownItemSelect(clusterField, "Smart Sampling");
            dropDownItemSelect(mediaTypeField, "Market");
            dropDownItemSelect(relationField, "Guardian");
            dropDownItemSelect(relationGuardianField, "Brother");
            dropDownItemSelect(CNICField, "CNIC");
            inputText(CNICinputField, CNIC);
            inputText(nameInputField, name);
            inputText(fNameInputField, fName);
            dropDownItemSelect(selectGenderField, gender);
            inputText(ageInputField, age);
            inputText(phone1InputField, contact);
            inputText(phone2InputField, contact);
            dropDownItemSelect(nationalityField, "Pakistani");
            dropDownItemSelect(divisionField, "string:035");
            Thread.Sleep(2000);
            dropDownItemSelect(districtField, "string:035004");
            Thread.Sleep(2000);
            dropDownItemSelect(tehsilField, "string:035004004");
            //Thread.Sleep(2000);
            inputText(addressInputField, address);
            dropDownItemSelect(patientConditionField, "Asymptomatic");
            dropDownItemSelect(selectVaccineField, "Vaccinated");
            dropDownItemSelect(selectVaccineTypeField, "Pfizer");

            ClickableItem(submitButton);
            Thread.Sleep(1000);
            acceptAlert();
            Thread.Sleep(1000);
        }
    }
}
