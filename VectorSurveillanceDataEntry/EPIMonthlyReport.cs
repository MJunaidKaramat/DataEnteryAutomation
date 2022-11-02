using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using FP_InterWood;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

namespace VectorSurveillanceDataEntry
{
    public class EPIMonthlyReport:GeneralMethodsClass
    {
        #region objectsForLogin
        By userName = By.XPath("//input[@id='username']");
        By userPassword = By.XPath("//input[@id='password']");
        By loginButton = By.XPath("//button[contains(text(),'EPI MIS Login')]");
        #endregion
        #region RedirectEntryPage
        By firstDropDownItemClick = By.XPath("//a[@href=\"https://punjab.epimis.pk/#\"]/span[contains(text(),'Service')]");
        By secondDropDownItemClick = By.XPath("//a[@href=\"https://punjab.epimis.pk/#\"]/span[contains(text(),'Mont')]");
        By thirdDropDownItemClick = By.XPath("//a[@href=\"https://punjab.epimis.pk/vaccination\"]/span[contains(text(),'Mont')]");
        By addNewReportButton = By.XPath("//*[@id=\"skin-blue\"]/div[2]/div[2]/section/div/div/div/div[2]/table/thead/tr/th[7]/a");

        #endregion
        #region DataEntryFormElements
        //UC Selection
        By selectTehsilDropDown = By.XPath("//select[@id='tcode']");
        By selectUCDropDown = By.XPath("//select[@id='uncode']");
        By totalFixedCentersField = By.Name("tot_fixed_centers");
        By totalFunctioningCentersField = By.Name("functioning_centers");
        By totalReportingCentersField = By.Name("reporting_centers");
        By totalPlannedSessionsField = By.Name("fixed_vacc_planned");
        By totalHeldSessionsField = By.Name("fixed_vacc_held");
        By totalOutPlannedSessionsField = By.Name("or_vacc_planned");
        By totalOutHeldSessionsField = By.Name("or_vacc_held");
        //Childhood Routine Immunization Coverage
        By FHepVacM = By.Name("vaccinated[hepb_r1_f2]");
        By FHepVacF = By.Name("vaccinated[hepb_r2_f2]");
        By OHepVacM = By.Name("vaccinated[hepb_r7_f2]");
        By OHepVacF = By.Name("vaccinated[hepb_r8_f2]");

        By FBCGVacM = By.Name("vaccinated[bcg_r1_f1]");
        By FBCGVacF = By.Name("vaccinated[bcg_r2_f1]");
        By OBCGVacM = By.Name("vaccinated[bcg_r7_f1]");
        By OBCGVacF = By.Name("vaccinated[bcg_r8_f1]");

        By FOPV0M = By.Name("vaccinated[opv0_r1_f3]");
        By FOPV0F = By.Name("vaccinated[opv0_r2_f3]");
        By OOPV0M = By.Name("vaccinated[opv0_r7_f3]");
        By OOPV0F = By.Name("vaccinated[opv0_r8_f3]");
        By FOPV1M = By.Name("vaccinated[opv1_r1_f4]");
        By FOPV1F = By.Name("vaccinated[opv1_r2_f4]");
        By OOPV1M = By.Name("vaccinated[opv1_r7_f4]");
        By OOPV1F = By.Name("vaccinated[opv1_r8_f4]");
        By FOPV2M = By.Name("vaccinated[opv2_r1_f5]");
        By FOPV2F = By.Name("vaccinated[opv2_r2_f5]");
        By OOPV2M = By.Name("vaccinated[opv2_r7_f5]");
        By OOPV2F = By.Name("vaccinated[opv2_r8_f5]");
        By FOPV3M = By.Name("vaccinated[opv3_r1_f6]");
        By FOPV3F = By.Name("vaccinated[opv3_r2_f6]");
        By OOPV3M = By.Name("vaccinated[opv3_r7_f6]");
        By OOPV3F = By.Name("vaccinated[opv3_r8_f6]");

        By FPenta1M = By.Name("vaccinated[penta1_r1_f7]");
        By FPenta1F = By.Name("vaccinated[penta1_r2_f7]");
        By OPenta1M = By.Name("vaccinated[penta1_r7_f7]");
        By OPenta1F = By.Name("vaccinated[penta1_r8_f7]");
        By FPenta2M = By.Name("vaccinated[penta2_r1_f8]");
        By FPenta2F = By.Name("vaccinated[penta2_r2_f8]");
        By OPenta2M = By.Name("vaccinated[penta2_r7_f8]");
        By OPenta2F = By.Name("vaccinated[penta2_r8_f8]");
        By FPenta3M = By.Name("vaccinated[penta3_r1_f9]");
        By FPenta3F = By.Name("vaccinated[penta3_r2_f9]");
        By OPenta3M = By.Name("vaccinated[penta3_r7_f9]");
        By OPenta3F = By.Name("vaccinated[penta3_r8_f9]");

        By FPCV1M = By.Name("vaccinated[pcv1_r1_f10]");
        By FPCV1F = By.Name("vaccinated[pcv1_r2_f10]");
        By OPCV1M = By.Name("vaccinated[pcv1_r7_f10]");
        By OPCV1F = By.Name("vaccinated[pcv1_r8_f10]");
        By FPCV2M = By.Name("vaccinated[pcv2_r1_f11]");
        By FPCV2F = By.Name("vaccinated[pcv2_r2_f11]");
        By OPCV2M = By.Name("vaccinated[pcv2_r7_f11]");
        By OPCV2F = By.Name("vaccinated[pcv2_r8_f11]");
        By FPCV3M = By.Name("vaccinated[pcv3_r1_f12]");
        By FPCV3F = By.Name("vaccinated[pcv3_r2_f12]");
        By OPCV3M = By.Name("vaccinated[pcv3_r7_f12]");
        By OPCV3F = By.Name("vaccinated[pcv3_r8_f12]");

        By FRota1M = By.Name("vaccinated[rota1_r1_f14]");
        By FRota1F = By.Name("vaccinated[rota1_r2_f14]");
        By ORota1M = By.Name("vaccinated[rota1_r7_f14]");
        By ORota1F = By.Name("vaccinated[rota1_r8_f14]");
        By FRota2M = By.Name("vaccinated[rota2_r1_f15]");
        By FRota2F = By.Name("vaccinated[rota2_r2_f15]");
        By ORota2M = By.Name("vaccinated[rota2_r7_f15]");
        By ORota2F = By.Name("vaccinated[rota2_r8_f15]");

        By FIPV1M = By.Name("vaccinated[ipv1_r1_f13]");
        By FIPV1F = By.Name("vaccinated[ipv1_r2_f13]");
        By OIPV1M = By.Name("vaccinated[ipv1_r7_f13]");
        By OIPV1F = By.Name("vaccinated[ipv1_r8_f13]");
        By FIPV2M = By.Name("vaccinated[ipv2_r1_f19]");
        By FIPV2F = By.Name("vaccinated[ipv2_r2_f19]");
        By OIPV2M = By.Name("vaccinated[ipv2_r7_f19]");
        By OIPV2F = By.Name("vaccinated[ipv2_r8_f19]");

        By FTCVM = By.Name("vaccinated[tcv_r1_f20]");
        By FTCVF = By.Name("vaccinated[tcv_r2_f20]");
        By OTCVM = By.Name("vaccinated[tcv_r7_f20]");
        By OTCVF = By.Name("vaccinated[tcv_r8_f20]");
        
        By FMR1M = By.Name("vaccinated[mr1_r1_f21]");
        By FMR1F = By.Name("vaccinated[mr1_r2_f21]");
        By OMR1M = By.Name("vaccinated[mr1_r7_f21]");
        By OMR1F = By.Name("vaccinated[mr1_r8_f21]");
        By FMR2M = By.Name("vaccinated[mr2_r3_f22]");
        By FMR2F = By.Name("vaccinated[mr2_r4_f21]");
        By OMR2M = By.Name("vaccinated[mr2_r9_f21]");
        By OMR2F = By.Name("vaccinated[mr2_r10_f21]");

        By FDTPM = By.Name("vaccinated[dtp_r3_f23]");
        By FDTPF = By.Name("vaccinated[dtp_r4_f23]");
        By ODTPM = By.Name("vaccinated[dtp_r9_f23]");
        By ODTPF = By.Name("vaccinated[dtp_r10_f23]");
        //B: Immunization Coverage Data Segregation
        By HEBSegDue = By.Name("coverage_segregation[hepb_seg_r1_f2]");
        By HEBSegDef = By.Name("coverage_segregation[hepb_seg_r2_f2]");
        By HEBSegUCOut = By.Name("coverage_segregation[hepb_oui_r3_f2]");
        By HEBSegDisOut = By.Name("coverage_segregation[hepb_od_r4_f2]");

        By BCGSegDue = By.Name("coverage_segregation[bcg_seg_r1_f1]");
        By BCGSegDef = By.Name("coverage_segregation[bcg_seg_r2_f1]");
        By BCGSegUCOut = By.Name("coverage_segregation[bcg_oui_r3_f1]");
        By BCGSegDisOut = By.Name("coverage_segregation[bcg_od_r4_f1]");

        By OPV0SegDue = By.Name("coverage_segregation[opv0_seg_r1_f3]");
        By OPV0SegDef = By.Name("coverage_segregation[opv0_seg_r2_f3]");
        By OPV0SegUCOut = By.Name("coverage_segregation[opv0_oui_r3_f3]");
        By OPV0SegDisOut = By.Name("coverage_segregation[opv0_od_r4_f3]");
        By OPV1SegDue = By.Name("coverage_segregation[opv1_seg_r1_f4]");
        By OPV1SegDef = By.Name("coverage_segregation[opv1_seg_r2_f4]");
        By OPV1SegUCOut = By.Name("coverage_segregation[opv1_oui_r3_f4]");
        By OPV1SegDisOut = By.Name("coverage_segregation[opv1_od_r4_f4]");
        By OPV2SegDue = By.Name("coverage_segregation[opv2_seg_r1_f5]");
        By OPV2SegDef = By.Name("coverage_segregation[opv2_seg_r2_f5]");
        By OPV2SegUCOut = By.Name("coverage_segregation[opv2_oui_r3_f5]");
        By OPV2egDisOut = By.Name("coverage_segregation[opv2_od_r4_f5]");
        By OPV3SegDue = By.Name("coverage_segregation[opv3_seg_r1_f6]");
        By OPV3SegDef = By.Name("coverage_segregation[opv3_seg_r2_f6]");
        By OPV3SegUCOut = By.Name("coverage_segregation[opv3_oui_r3_f6]");
        By OPV3SegDisOut = By.Name("coverage_segregation[opv3_od_r4_f6]");

        By Penta1SegDue = By.Name("coverage_segregation[penta1_seg_r1_f7]");
        By Penta1SegDef = By.Name("coverage_segregation[penta1_seg_r2_f7]");
        By Penta1SegUCOut = By.Name("coverage_segregation[penta1_oui_r3_f7]");
        By Penta1SegDisOut = By.Name("coverage_segregation[penta1_od_r4_f7]");
        By Penta2SegDue = By.Name("coverage_segregation[penta2_seg_r1_f8]");
        By Penta2SegDef = By.Name("coverage_segregation[penta2_seg_r2_f8]");
        By Penta2SegUCOut = By.Name("coverage_segregation[penta2_oui_r3_f8]");
        By Penta2SegDisOut = By.Name("coverage_segregation[penta2_od_r4_f8]");
        By Penta3SegDue = By.Name("coverage_segregation[penta3_seg_r1_f9]");
        By Penta3SegDef = By.Name("coverage_segregation[penta3_seg_r2_f9]");
        By Penta3SegUCOut = By.Name("coverage_segregation[penta3_oui_r3_f9]");
        By Penta3SegDisOut = By.Name("coverage_segregation[penta3_od_r4_f9]");

        By Rota1SegDue = By.Name("coverage_segregation[rota1_seg_r1_f14]");
        By Rota1SegDef = By.Name("coverage_segregation[rota1_seg_r2_f14]");
        By Rota1SegUCOut = By.Name("coverage_segregation[rota1_oui_r3_f14]");
        By Rota1SegDisOut = By.Name("coverage_segregation[rota1_od_r4_f14]");
        By Rota2SegDue = By.Name("coverage_segregation[rota2_seg_r1_f15]");
        By Rota2SegDef = By.Name("coverage_segregation[rota2_seg_r2_f15]");
        By Rota2SegUCOut = By.Name("coverage_segregation[rota2_oui_r3_f15]");
        By Rota2SegDisOut = By.Name("coverage_segregation[rota2_od_r4_f15]");

        By IPV1SegDue = By.Name("coverage_segregation[ipv1_seg_r1_f13]");
        By IPV1SegDef = By.Name("coverage_segregation[ipv1_seg_r2_f13]");
        By IPV1SegUCOut = By.Name("coverage_segregation[ipv1_oui_r3_f13]");
        By IPV1SegDisOut = By.Name("coverage_segregation[ipv1_od_r4_f13]");
        By IPV2SegDue = By.Name("coverage_segregation[ipv2_seg_r1_f19]");
        By IPV2SegDef = By.Name("coverage_segregation[ipv2_seg_r2_f19]");
        By IPV2SegUCOut = By.Name("coverage_segregation[ipv2_oui_r3_f19]");
        By IPV2SegDisOut = By.Name("coverage_segregation[ipv2_od_r4_f19]");
        
        By TCVSegDue = By.Name("coverage_segregation[tcv_seg_r1_f20]");
        By TCVSegDef = By.Name("coverage_segregation[tcv_seg_r2_f20]");
        By TCVSegUCOut = By.Name("coverage_segregation[tcv_oui_r3_f20]");
        By TCVSegDisOut = By.Name("coverage_segregation[tcv_od_r4_f20]");

        By MR1SegDue = By.Name("coverage_segregation[mr1_seg_r1_f21]");
        By MR1SegDef = By.Name("coverage_segregation[mr1_seg_r2_f21]");
        By MR1SegUCOut = By.Name("coverage_segregation[mr1_oui_r3_f21]");
        By MR1SegDisOut = By.Name("coverage_segregation[mr1_od_r4_f21]");
        By MR2SegDue = By.Name("coverage_segregation[mr2_seg_r1_f22]");
        By MR2SegDef = By.Name("coverage_segregation[mr2_seg_r2_f22]");
        By MR2SegUCOut = By.Name("coverage_segregation[mr2_oui_r3_f22]");
        By MR2SegDisOut = By.Name("coverage_segregation[mr2_od_r4_f22]");

        By DTPSegDue = By.Name("coverage_segregation[dtp_seg_r1_f23]");
        By DTPSegDef = By.Name("coverage_segregation[dtp_seg_r2_f23]");
        By DTPSegUCOut = By.Name("coverage_segregation[dtp_oui_r3_f23]");
        By DTPSegDisOut = By.Name("coverage_segregation[dtp_od_r4_f23]");
        //C: TD routine immunization Coverage
        By FTT1 = By.Name("td_routine[td1_r1_f1]");
        By FTT2 = By.Name("td_routine[td2_r1_f2]");
        By FTT3 = By.Name("td_routine[td3_r1_f3]");
        By FTT4 = By.Name("td_routine[td4_r1_f4]");
        By FTT5 = By.Name("td_routine[td5_r1_f5]");
        By OTT1 = By.Name("td_routine[td1_r3_f1]");
        By OTT2 = By.Name("td_routine[td2_r3_f2]");
        By OTT3 = By.Name("td_routine[td3_r3_f3]");
        By OTT4 = By.Name("td_routine[td4_r3_f4]");
        By OTT5 = By.Name("td_routine[td5_r3_f5]");
        //D: Vaccine & Logistics Management
        By HEPBOpenBalance = By.Name("vm[vm1_r1_f2]");
        By HEPBReceived = By.Name("vm[vm1_r2_f2]");
        By HEPBCloseBalance = By.Name("vm[vm1_r3_f2]");
        By BCGOpenBalance = By.Name("vm[vm2_r1_f1]");
        By BCGReceived = By.Name("vm[vm2_r2_f1]");
        By BCGCloseBalance = By.Name("vm[vm2_r3_f1]");
        By OPVOpenBalance = By.Name("vm[vm3_r1_f3]");
        By OPVReceived = By.Name("vm[vm3_r2_f3]");
        By OPVCloseBalance = By.Name("vm[vm3_r3_f3]");
        By PentaOpenBalance = By.Name("vm[vm4_r1_f7]");
        By PentaReceived = By.Name("vm[vm4_r2_f7]");
        By PentaCloseBalance = By.Name("vm[vm4_r3_f7]");
        By PCVOpenBalance = By.Name("vm[vm5_r1_f10]");
        By PCVReceived = By.Name("vm[vm5_r2_f10]");
        By PCVCloseBalance = By.Name("vm[vm5_r3_f10]");
        By ROTAOpenBalacne = By.Name("vm[vm6_r1_f14]");
        By ROTAReceived = By.Name("vm[vm6_r2_f14]");
        By ROTACloseBalance = By.Name("vm[vm6_r3_f14]");
        By IPVOpenBalacne = By.Name("vm[vm7_r1_f13]");
        By IPVReceived = By.Name("vm[vm7_r2_f13]");
        By IPVCloseBalance = By.Name("vm[vm7_r3_f13]");
        By TCVOpenBalacne = By.Name("vm[vm8_r1_f20]");
        By TCVReceived = By.Name("vm[vm8_r2_f20]");
        By TCVCloseBalance = By.Name("vm[vm8_r3_f20]");
        By MROpenBalacne = By.Name("vm[vm9_r1_f21]");
        By MRReceived = By.Name("vm[vm9_r2_f21]");
        By MRCloseBalance = By.Name("vm[vm9_r3_f21]");
        By DTPOpenBalacne = By.Name("vm[vm10_r1_f23]");
        By DTPReceived = By.Name("vm[vm10_r2_f23]");
        By DTPCloseBalance = By.Name("vm[vm10_r3_f23]");
        By TTOpenBalacne = By.Name("vm[vm11_r1_f24]");
        By TTReceived = By.Name("vm[vm11_r2_f24]");
        By TTCloseBalance = By.Name("vm[vm11_r3_f24]");
        By AD005OpenBalacne = By.Name("vm[vm12_r1_f25]");
        By AD005Received = By.Name("vm[vm12_r2_f25]");
        By AD005CloseBalance = By.Name("vm[vm12_r3_f25]");
        By AD05OpenBalacne = By.Name("vm[vm13_r1_f26]");
        By AD05Received = By.Name("vm[vm13_r2_f26]");
        By AD05CloseBalance = By.Name("vm[vm13_r3_f26]");
        By Syr2OpenBalacne = By.Name("vm[vm14_r1_f27]");
        By Syr2Received = By.Name("vm[vm14_r2_f27]");
        By Syr2CloseBalance = By.Name("vm[vm14_r3_f27]");
        By Syr5OpenBalacne = By.Name("vm[vm15_r1_f28]");
        By Syr5Received = By.Name("vm[vm15_r2_f28]");
        By Syr5CloseBalance = By.Name("vm[vm15_r3_f28]");
        By CardChildOpenBalacne = By.Name("vm[vm16_r1_f30]");
        By CardChildReceived = By.Name("vm[vm16_r2_f30]");
        By CardChildCloseBalance = By.Name("vm[vm16_r3_f30]");
        By CardWomenOpenBalacne = By.Name("vm[vm17_r1_f31]");
        By CardWomenReceived = By.Name("vm[vm17_r2_f31]");
        By CardWomenCloseBalance = By.Name("vm[vm17_r3_f31]");
        //E: EPI Waste Manangement
        By UsedSafteyBox = By.Name("lm[lvm1_r1_f1]");
        By DisposedSafteyBox = By.Name("lm[lvm1_r2_f1]");
        By methodDisposeDropDown1 = By.Name("lm[lvm1_r3_f1]");//Value 1
        By UsedVials = By.Name("lm[lvm2_r1_f2]");
        By DisposedVials = By.Name("lm[lvm2_r2_f2]");
        By methodDisposeDropDown2 = By.Name("lm[lvm2_r3_f2]");//Value 1
        By UnusedVials = By.Name("lm[lvm3_r1_f3]");
        By UnusedDisposedVials = By.Name("lm[lvm3_r2_f3]");
        By methodDisposeDropDown3 = By.Name("lm[lvm3_r3_f3]");//Value 1
        //F: EPI Missed Children 
        By MissedNA = By.XPath("//input[@name='missed_r1']");
        By Refusal = By.XPath("//input[@name='missed_r2']");
        By Sick = By.XPath("//input[@name='missed_r3']");
        //G: New Child Registeration
        By newMaleChild = By.XPath("//input[@name='new_born_m']");
        By newFemaleChild = By.XPath("//input[@name='new_born_f']");
        //H: Zero Dose Childern
        By recordedZeroDue = By.XPath("//input[@name='recorded_zero_due']");
        By recordedZeroDefaulter = By.XPath("//input[@name='recorded_zero_defaulter']");
        By coverageZeroDue = By.XPath("//input[@name='covered_zero_due']");
        By coverageZeroDefaulter = By.XPath("//input[@name='covered_zero_defaulter']");

        #endregion
        public void Login()
        {
            inputText(userName, "sheikhupura154_deo");
            inputText(userPassword, "pnj5842");
            ClickableItem(loginButton);
        }
        public void EntryPage(string UC, string tFixedCenters, string tFunctioningCenters, string reportingCenters, string planned, string sessionsHeld, 
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
            string DTPSegDue1, string DTPSegDef2, string DTPSegUCOut3, string DTPSegDisOut4



)
        {
            #region done
            ClickableItem(firstDropDownItemClick);
            ClickableItem(secondDropDownItemClick);
            ClickableItem(thirdDropDownItemClick);
            ClickableItem(addNewReportButton);

            dropDownItemSelect(selectUCDropDown, UC);
            inputText(totalFixedCentersField, tFixedCenters);
            inputText(totalFunctioningCentersField, tFunctioningCenters);
            inputText(totalReportingCentersField, reportingCenters);
            inputText(totalPlannedSessionsField, planned);
            inputText(totalHeldSessionsField, sessionsHeld);
            inputText(totalOutPlannedSessionsField, outSessionsPlaned);
            inputText(totalOutHeldSessionsField, outSessionsHeld);

            inputText(FHepVacM, hepF1);
            inputText(FHepVacF, hepF2);
            inputText(OHepVacM, hepO1);
            inputText(OHepVacF, hepO2);
            inputText(FBCGVacM, bcg01);
            inputText(FBCGVacF, bcg02);
            inputText(OBCGVacM, bcg03);
            inputText(OBCGVacF, bcg04);
            inputText(FOPV0M, opv00);
            inputText(FOPV0F, opv01);
            inputText(OOPV0M, opv02);
            inputText(OOPV0F, opv03);
            inputText(FOPV1M, opv10);
            inputText(FOPV1F, opv11);
            inputText(OOPV1M, opv12);
            inputText(OOPV1F, opv13);
            inputText(FOPV2M, opv20);
            inputText(FOPV2F, opv21);
            inputText(OOPV2M, opv22);
            inputText(OOPV2F, opv23);
            inputText(FOPV3M, opv30);
            inputText(FOPV3F, opv31);
            inputText(OOPV3M, opv32);
            inputText(OOPV3F, opv33);

            inputText(FPenta1M, FPenta11);
            inputText(FPenta1F, FPenta12);
            inputText(OPenta1M, OPenta13);
            inputText(OPenta1F, OPenta14);
            inputText(FPenta2M, FPenta21);
            inputText(FPenta2F, FPenta22);
            inputText(OPenta2M, OPenta23);
            inputText(OPenta2F, OPenta24);
            inputText(FPenta3M, FPenta31);
            inputText(FPenta3F, FPenta32);
            inputText(OPenta3M, OPenta33);
            inputText(OPenta3F, OPenta34);

            inputText(FPCV1M, FPCV11);
            inputText(FPCV1F, FPCV12);
            inputText(OPCV1M, OPCV13);
            inputText(OPCV1F, OPCV14);
            inputText(FPCV2M, FPCV21);
            inputText(OPCV2F, FPCV22);
            inputText(OPCV2M, OPCV23);
            inputText(FPCV2F, OPCV24);
            inputText(FPCV3M, FPCV31);
            inputText(FPCV3F, FPCV32);
            inputText(OPCV3M, OPCV33);
            inputText(OPCV3F, OPCV34);

            inputText(FRota1M, FRota11);
            inputText(FRota1F, FRota12);
            inputText(ORota1M, ORota13);
            inputText(ORota1F, ORota14);
            inputText(FRota2M, FRota21);
            inputText(FRota2F, FRota22);
            inputText(ORota2M, ORota23);
            inputText(ORota2F, ORota24);
            inputText(FIPV1M, FIPV11);
            inputText(FIPV1F, FIPV12);
            inputText(OIPV1M, OIPV13);
            inputText(OIPV1F, OIPV14);
            inputText(FIPV2M, FIPV21);
            inputText(FIPV2F, FIPV22);
            inputText(OIPV2M, OIPV23);
            inputText(OIPV2F, OIPV24);
            inputText(FTCVM, FTCV1);
            inputText(FTCVF, FTCV2);
            inputText(OTCVM, OTCV3);
            inputText(OTCVF, OTCV4);
            inputText(FMR1M, FMR11);
            inputText(FMR1F, FMR12);
            inputText(OMR1M, OMR13);
            inputText(OMR1F, OMR14);
            inputText(FMR2M, FMR21);
            inputText(FMR2F, FMR22);
            inputText(OMR2M, OMR23);
            inputText(OMR2F, OMR24);
            inputText(FDTPM, FDTP1);
            inputText(FDTPF, FDTP2);
            inputText(ODTPM, ODTP3);
            inputText(ODTPF, ODTP4);
            inputText(HEBSegDue, HEBSegDue1);
            inputText(HEBSegDef, HEBSegDef2);
            inputText(HEBSegUCOut, HEBSegUCOut3);
            inputText(HEBSegDisOut, HEBSegDisOut4);
            inputText(BCGSegDue, BCGSegDue1);
            inputText(BCGSegDef, BCGSegDef2);
            inputText(BCGSegUCOut, BCGSegUCOut3);
            inputText(BCGSegDisOut, BCGSegDisOut4);
            inputText(OPV0SegDue, OPV0SegDue1);
            inputText(OPV0SegDef, OPV0SegDef2);
            inputText(OPV0SegUCOut, OPV0SegUCOut3);
            inputText(OPV0SegDisOut, OPV0SegDisOut4);
            inputText(OPV1SegDue, OPV1SegDue1);
            inputText(OPV1SegDef, OPV1SegDef2);
            inputText(OPV1SegUCOut, OPV1SegUCOut3);
            inputText(OPV1SegDisOut, OPV1SegDisOut4);
            inputText(OPV2SegDue, OPV2SegDue1);
            inputText(OPV2SegDef, OPV2SegDef2);
            inputText(OPV2SegUCOut, OPV2SegUCOut3);
            inputText(OPV2egDisOut, OPV2egDisOut4);
            inputText(OPV3SegDue, OPV3SegDue1);
            inputText(OPV3SegDef, OPV3SegDef2);
            inputText(OPV3SegUCOut, OPV3SegUCOut3);
            inputText(OPV3SegDisOut, OPV3SegDisOut4);
            inputText(Penta1SegDue, Penta1SegDue11);
            inputText(Penta1SegDef, Penta1SegDef12);
            inputText(Penta1SegUCOut, Penta1SegUCOut13);
            inputText(Penta1SegDisOut, Penta1SegDisOut14);
            inputText(Penta2SegDue, Penta2SegDue21);
            inputText(Penta2SegDef, Penta2SegDef22);
            inputText(Penta2SegUCOut, Penta2SegUCOut23);
            inputText(Penta2SegDisOut, Penta2SegDisOut24);
            inputText(Penta3SegDue, Penta3SegDue31);
            inputText(Penta3SegDef, Penta3SegDef32);
            inputText(Penta3SegUCOut, Penta3SegUCOut33);
            inputText(Penta3SegDisOut, Penta3SegDisOut34);
            inputText(Rota1SegDue, Rota1SegDue11);
            inputText(Rota1SegDef, Rota1SegDef12);
            inputText(Rota1SegUCOut, Rota1SegUCOut13);
            inputText(Rota1SegDisOut, Rota1SegDisOut14);
            inputText(Rota2SegDue, Rota2SegDue21);
            inputText(Rota2SegDef, Rota2SegDef22);
            inputText(Rota2SegUCOut, Rota2SegUCOut23);
            inputText(Rota2SegDisOut, Rota2SegDisOut24);
            inputText(IPV1SegDue, IPV1SegDue11);
            inputText(IPV1SegDef, IPV1SegDef12);
            inputText(IPV1SegUCOut, IPV1SegUCOut13);
            inputText(IPV1SegDisOut, IPV1SegDisOut14);
            inputText(IPV2SegDue, IPV2SegDue21);
            inputText(IPV2SegDef, IPV2SegDef22);
            inputText(IPV2SegUCOut, IPV2SegUCOut22);
            inputText(IPV2SegDisOut, IPV2SegDisOut24);
            inputText(TCVSegDue, TCVSegDue1);
            inputText(TCVSegDef, TCVSegDef2);
            inputText(TCVSegUCOut, TCVSegUCOut3);
            inputText(TCVSegDisOut, TCVSegDisOut4);
            inputText(MR1SegDue, MR1SegDue11);
            inputText(MR1SegDef, MR1SegDef12);
            inputText(MR1SegUCOut, MR1SegUCOut13);
            inputText(MR1SegDisOut, MR1SegDisOut14);
            inputText(MR2SegDue, MR2SegDue21);
            inputText(MR2SegDef, MR2SegDef22);
            inputText(MR2SegUCOut, MR2SegUCOut23);
            inputText(MR2SegDisOut, MR2SegDisOut24);
            inputText(DTPSegDue, DTPSegDue1);
            inputText(DTPSegDef, DTPSegDef2);
            inputText(DTPSegUCOut, DTPSegUCOut3);
            inputText(DTPSegDisOut, DTPSegDisOut4);

            #endregion





        }
    }
}
