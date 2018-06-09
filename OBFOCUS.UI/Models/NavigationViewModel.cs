using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OBFOCUS.UI.Utils;

namespace OBFOCUS.UI.Models
{
    public class NavigationViewModel
    {
        public string Header { get; set; }
        public string Image { get; set; }
        public bool NewItem { get; set; }
        public List<NavigationViewModel> Items { get; set; }

        public NavigationViewModel()
        { }

        public NavigationViewModel(string HeaderText, string Image)
        {
            this.Header = HeaderText;
            this.Image = Image;
        }

        //public static NavigationViewModel[] GetData(string val)
        //{
        //    return new NavigationViewModel[]
        //    {
        //        new NavigationViewModel
        //        {
        //            Header = "Electronics",
        //            Image = "/Content/images/electronics.png",
        //            Items = new NavigationViewModel[]
        //            {
        //                new NavigationViewModel { Header="Trimmers/Shavers" },
        //                new NavigationViewModel { Header="Tablets" },
        //                new NavigationViewModel { Header="Phones",
        //                    Image ="/Content/images/phones.png",
        //                    Items = new NavigationViewModel[] {
        //                        new NavigationViewModel { Header="Apple" },
        //                        new NavigationViewModel { Header="Motorola", NewItem=true },
        //                        new NavigationViewModel { Header="Nokia" },
        //                        new NavigationViewModel { Header="Samsung" }}
        //                },
        //                new NavigationViewModel { Header="Speakers", NewItem=true },
        //                new NavigationViewModel { Header="Monitors" }
        //            }
        //        },
        //        new NavigationViewModel{
        //            Header = "Toys",
        //            Image = "/Content/images/toys.png",
        //            Items = new NavigationViewModel[]{
        //                new NavigationViewModel{ Header = "Shopkins" },
        //                new NavigationViewModel{ Header = "Train Sets" },
        //                new NavigationViewModel{ Header = "Science Kit", NewItem = true },
        //                new NavigationViewModel{ Header = "Play-Doh" },
        //                new NavigationViewModel{ Header = "Crayola" }
        //            }
        //        },
        //        new NavigationViewModel{
        //            Header = "Home",
        //            Image = "/Content/images/home.png",
        //            Items = new NavigationViewModel[] {
        //                new NavigationViewModel{ Header = "Coffeee Maker" },
        //                new NavigationViewModel{ Header = "Breadmaker", NewItem = true },
        //                new NavigationViewModel{ Header = "Solar Panel", NewItem = true },
        //                new NavigationViewModel{ Header = "Work Table" },
        //                new NavigationViewModel{ Header = "Propane Grill" }
        //            }
        //        }
        //    };
        //}

        public List<NavigationViewModel> LoadTreeView(string role)
        {
            List<NavigationViewModel> navigation = new List<NavigationViewModel>();
            NavigationViewModel pNode = null;
            NavigationViewModel cNode = null;
            

            role = role.ToUpper();
            if (role == "GENERAL" || role == "FULL" || role == "LIMITED" || role == "GENETICIST")
            {
                pNode = new NavigationViewModel("Maintain Data", CommonUtil.GetImagePath("SwitchBoard/MaintainData.ico"));
                navigation.Add(pNode);
            }

            pNode.Items = new List<NavigationViewModel>();

            if (role == "GENERAL" || role == "FULL" || role == "LIMITED")
            {
                cNode = new NavigationViewModel("Charts", CommonUtil.GetImagePath("Classes/Charts.ico"));
                pNode.Items.Add(cNode);
            }

            if (role == "GENERAL" || role == "FULL" || role == "GENETICIST")
            {
                cNode = new NavigationViewModel("Charts", CommonUtil.GetImagePath("Classes/Examinations.ico"));
                pNode.Items.Add(cNode);
            }

            if (role == "GENERAL" || role == "FULL")
            {
                cNode = new NavigationViewModel("Outcome", CommonUtil.GetImagePath("Classes/Outcome.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Diagnosis", CommonUtil.GetImagePath("Classes/PatientsByDiagnosis.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Antenatal", CommonUtil.GetImagePath("Classes/antenatal.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Flagged Charts", CommonUtil.GetImagePath("Classes/PendingLaboratories.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Chart Documents", CommonUtil.GetImagePath("Classes/WorkBooks.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Defaults", CommonUtil.GetImagePath("Classes/Defaults.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Examiners", CommonUtil.GetImagePath("Classes/Physician.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Site", CommonUtil.GetImagePath("Classes/Site.ico"));
                pNode.Items.Add(cNode);

                pNode = new NavigationViewModel("Views and Reports", CommonUtil.GetImagePath("SwitchBoard/ViewsAndReports.ico"));
                navigation.Add(pNode);
                cNode = new NavigationViewModel("Amniocentesis Log", CommonUtil.GetImagePath("Classes/AmniocentesisLog.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Statistics By Diagnosis", CommonUtil.GetImagePath("Classes/StatisticsByDiagnosis.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Outcome by Diagnosis", CommonUtil.GetImagePath("Classes/OutcomeByDiagnosis.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Referring Practitioners", CommonUtil.GetImagePath("Classes/ReferringPractioners.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Undelivered Tracked Anomalies", CommonUtil.GetImagePath("Classes/UndeliveredTrackedAnomalies.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Word Templates", CommonUtil.GetImagePath("Classes/Word.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Patients By Del Hospital", CommonUtil.GetImagePath("Classes/workbooks.ico"));
                pNode.Items.Add(cNode);

                pNode = new NavigationViewModel("Reference", CommonUtil.GetImagePath("SwitchBoard/Reference.ico"));
                navigation.Add(pNode);
                cNode = new NavigationViewModel("Formulary", CommonUtil.GetImagePath("Classes/formulary.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Knowledge Base", CommonUtil.GetImagePath("Classes/KnowledgeBase.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Image Catalogue", CommonUtil.GetImagePath("Classes/ImageCatalog.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Medline", CommonUtil.GetImagePath("Classes/medline.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Patient Instructions", CommonUtil.GetImagePath("Classes/PatientInstruction.ico"));
                pNode.Items.Add(cNode);

                pNode = new NavigationViewModel("Tools", CommonUtil.GetImagePath("SwitchBoard/Tools.ico"));
                navigation.Add(pNode);
                cNode = new NavigationViewModel("AFI Percentile", CommonUtil.GetImagePath("Classes/AFIPercentile.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("ANA ImmunoFluorescence Patterns", CommonUtil.GetImagePath("Classes/ANAImmunofluorescence.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Appendicitis Algorithms", CommonUtil.GetImagePath("Classes/AppendicitisAlgorithm.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Diabetic Calculator", CommonUtil.GetImagePath("Classes/DiabetesCalculator.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Fetal Length", CommonUtil.GetImagePath("Classes/FetalLength.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Gestational Age", CommonUtil.GetImagePath("Classes/GestationalAge.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Hepatitis Serology", CommonUtil.GetImagePath("Classes/HepatitisSerology.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Hypokalemia", CommonUtil.GetImagePath("Classes/HypokalemiaSerology.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Peak Flow", CommonUtil.GetImagePath("Classes/PeakFlow.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Risk from Teratogen Exposure", CommonUtil.GetImagePath("Classes/Risk.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Risk of AD Disease", CommonUtil.GetImagePath("Classes/Risk.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Risk of AR Disease", CommonUtil.GetImagePath("Classes/Risk.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Risk of Preterm Delivery", CommonUtil.GetImagePath("Classes/Risk.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Syncope Vs. Vertigo", CommonUtil.GetImagePath("Classes/syncopeVsVertigo.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Thyroid Function Tests", CommonUtil.GetImagePath("Classes/ThyroidFunctionTest.ico"));
                pNode.Items.Add(cNode);

                pNode = new NavigationViewModel("User Administration", CommonUtil.GetImagePath("SwitchBoard/MaintainData.ico"));
                navigation.Add(pNode);
                cNode = new NavigationViewModel("Merge Records", CommonUtil.GetImagePath("Classes/Mergerecords.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Transpose Records", CommonUtil.GetImagePath("Classes/transposerecords.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Unlock Chart Records", CommonUtil.GetImagePath("Project/key.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Physicians", CommonUtil.GetImagePath("Classes/Physician.ico"));
                pNode.Items.Add(cNode);
                cNode = new NavigationViewModel("Program Settings", CommonUtil.GetImagePath("Classes/security.ico"));
                pNode.Items.Add(cNode);
            }

            return navigation;
        }
    }
}
