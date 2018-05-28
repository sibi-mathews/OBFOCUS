using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OBFOCUS.UI.Models
{
    public class NavigationViewModel
    {
        public string Header { get; set; }
        public string Image { get; set; }
        public bool NewItem { get; set; }
        public NavigationViewModel[] Items { get; set; }

        public static NavigationViewModel[] GetData(string val)
        {
            return new NavigationViewModel[]
            {
                new NavigationViewModel
                {
                    Header = "Electronics",
                    Image = "/Content/images/electronics.png",
                    Items = new NavigationViewModel[]
                    {
                        new NavigationViewModel { Header="Trimmers/Shavers" },
                        new NavigationViewModel { Header="Tablets" },
                        new NavigationViewModel { Header="Phones",
                            Image ="/Content/images/phones.png",
                            Items = new NavigationViewModel[] {
                                new NavigationViewModel { Header="Apple" },
                                new NavigationViewModel { Header="Motorola", NewItem=true },
                                new NavigationViewModel { Header="Nokia" },
                                new NavigationViewModel { Header="Samsung" }}
                        },
                        new NavigationViewModel { Header="Speakers", NewItem=true },
                        new NavigationViewModel { Header="Monitors" }
                    }
                },
                new NavigationViewModel{
                    Header = "Toys",
                    Image = "/Content/images/toys.png",
                    Items = new NavigationViewModel[]{
                        new NavigationViewModel{ Header = "Shopkins" },
                        new NavigationViewModel{ Header = "Train Sets" },
                        new NavigationViewModel{ Header = "Science Kit", NewItem = true },
                        new NavigationViewModel{ Header = "Play-Doh" },
                        new NavigationViewModel{ Header = "Crayola" }
                    }
                },
                new NavigationViewModel{
                    Header = "Home",
                    Image = "/Content/images/home.png",
                    Items = new NavigationViewModel[] {
                        new NavigationViewModel{ Header = "Coffeee Maker" },
                        new NavigationViewModel{ Header = "Breadmaker", NewItem = true },
                        new NavigationViewModel{ Header = "Solar Panel", NewItem = true },
                        new NavigationViewModel{ Header = "Work Table" },
                        new NavigationViewModel{ Header = "Propane Grill" }
                    }
                }
            };
        }
    }
}
