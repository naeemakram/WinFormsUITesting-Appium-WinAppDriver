using System;
using System.Diagnostics;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;

namespace WinFormsUITesting
{
    [TestClass]
    public class WinFormTests
    {
        static WindowsDriver<WindowsElement> sessionWinForm;

        [ClassInitialize]
        public static void FirstThingsFirst(TestContext testContext)
        {
            AppiumOptions dcWinForms = new AppiumOptions();
            dcWinForms.AddAdditionalCapability("app",
                WinFormsUITesting.Properties.Settings.Default.ApplicationPath);
            //@"C:\Users\Naeem\source\repos\DoNotDistrurbMortgageCalculatorFrom1999\DoNotDistrurbMortgageCalculatorFrom1999\bin\Debug\DoNotDistrurbMortgageCalculatorFrom1999.exe");

            sessionWinForm = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"),
                dcWinForms);
        }

        [TestMethod]
        public void PopupTest()
        {
            sessionWinForm.FindElementByName("OpenMessageStrip").Click();

            Thread.Sleep(1000);

            var appiumOptionsDesktop = new AppiumOptions();
            appiumOptionsDesktop.AddAdditionalCapability("app", "Root");

            var sessionDesktop = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appiumOptionsDesktop);

            sessionDesktop.FindElementByXPath
            ("/Pane[@Name =\"Desktop 1\"][@ClassName=\"#32769\"]/Window[starts-with(@AutomationId,\"Form\")]" +
            "[@Name=\"Do Not Distrurb Mortgage Calculator From 1999 v20.8.2.1b\"]/Window[@Name=\"Alert\"]" +
            "/Button[@Name=\"OK\"][@ClassName=\"Button\"]")            
            .Click();

            sessionWinForm.FindElementByName("OpenMessageStrip").Click();

            Thread.Sleep(1000);

            sessionWinForm.FindElementByName("Alert").FindElementByName("OK").Click();

        }

        [TestMethod]
        public void TreeTest()
        {
            var tvControl = sessionWinForm.FindElementByAccessibilityId("treeView1");

            foreach(var tn in tvControl.FindElementsByTagName("TreeItem"))
            {
                Debug.WriteLine(
                    $"*** BEFORE: {tn.Text} - Displayed: {tn.Displayed} - Enabled: {tn.Enabled} - Selected: {tn.Selected}");
            }

            var nodeWorld = tvControl.FindElementByName("World");
            DoubleClickElement(nodeWorld);

            Thread.Sleep(1000);

            foreach (var tn in tvControl.FindElementsByTagName("TreeItem"))
            {
                Debug.WriteLine(
                    $"*** AFTER: {tn.Text} - Displayed: {tn.Displayed} - Enabled: {tn.Enabled} - Selected: {tn.Selected}");
            }

            var nodeAsia = tvControl.FindElementByName("Europe");
            DoubleClickElement(nodeAsia);

            var nodePakistan = tvControl.FindElementByName("Austria");

            WebDriverWait wdvPakistan = new WebDriverWait(sessionWinForm, TimeSpan.FromSeconds(2));
            wdvPakistan.Until(x => nodePakistan.Displayed);

            nodePakistan.Click();

        }

        private static void DoubleClickElement(OpenQA.Selenium.Appium.AppiumWebElement nodeWorld)
        {
            Actions actsTree = new Actions(sessionWinForm);
            actsTree.MoveToElement(nodeWorld);
            actsTree.DoubleClick();
            actsTree.Perform();
			// nothing
        }

        [TestMethod]
        public void GridTest()
        {
            var ratesGrid = sessionWinForm.FindElementByName("Rates Grid");
            var allHeaders = ratesGrid.FindElementsByTagName("Header");

            Debug.WriteLine($"*** Headers count: {allHeaders.Count}");
            
            foreach(var h in allHeaders)
            {
                Debug.WriteLine($"*** {h.Text} - {h.Displayed} - {h.Enabled}");
            }

            var allCells = ratesGrid.FindElementsByTagName("DataItem");
            Debug.WriteLine($"Grid cells count: {allCells.Count}");

            foreach(var gridCell in allCells)
            {
                Debug.WriteLine($"*** Cell Name: {gridCell.GetAttribute("Name")} - Text: {gridCell.Text}");
                if(gridCell.GetAttribute("Name").StartsWith("State Row") && gridCell.Text.Equals("FL"))
                {
                    gridCell.Click();
                    Actions actsGrid = new Actions(sessionWinForm);
                    actsGrid.MoveToElement(gridCell);
                    actsGrid.MoveToElement(gridCell, (gridCell.Size.Width / 4) - 10, gridCell.Size.Height / 2);
                    actsGrid.Click();
                    actsGrid.Perform();
                }
            }
        }



        [TestMethod]
        public void CheckboxTest()
        {
            var check = sessionWinForm.FindElementByName("checkBox1");
            check.Click();

            System.Threading.Thread.Sleep(1000);

            System.Diagnostics.Debug.WriteLine($"Value of checkbox is: {check.Selected}");
        }

        [TestMethod]
        public void RadioTest()
        {
            var radioFirst = sessionWinForm.FindElementByName("First");

            System.Diagnostics.Debug.WriteLine($"***** Value of radio First: {radioFirst.Selected}");

            radioFirst.Click();

            System.Threading.Thread.Sleep(1000);

            System.Diagnostics.Debug.WriteLine($"***** Value of radio First: {radioFirst.Selected}");
        }

        [TestMethod]
        public void ComboTest()
        {
            var combo = sessionWinForm.FindElementByAccessibilityId("comboBox1");
            var open = combo.FindElementByName("Open");

            var listItems = combo.FindElementsByTagName("ListItem");
            Debug.WriteLine($"Before: Number of list items found: {listItems.Count}");


            combo.SendKeys(Keys.Down);
            open.Click();

            listItems = combo.FindElementsByTagName("ListItem");
            Debug.WriteLine($"After: Number of list items found: {listItems.Count}");

            // maybe check number of elements in combo 
            //Assert.AreEqual(6, listItems.Count, "Combo box doesn't contain expected number of elements.");

            foreach(var comboKid in listItems)
            {
                if(comboKid.Text == "NJ")
                {
                    WebDriverWait wdv = new WebDriverWait(sessionWinForm, TimeSpan.FromSeconds(10));
                    wdv.Until(x => comboKid.Displayed);

                    comboKid.Click();
                }
            }
        }

        [TestMethod]
        public void MenuTest()
        {
            var allMenus = sessionWinForm.FindElementsByTagName("MenuItem");

            Debug.WriteLine($"All menu items found by search: {allMenus.Count}");

            foreach (var m in allMenus)
            {
                Debug.WriteLine($"+++++ Menu: {m.GetAttribute("Name")} - Displayed: {m.Displayed}");
            }

            WebDriverWait wdv = new WebDriverWait(sessionWinForm, TimeSpan.FromSeconds(10));

            foreach (var mainMenuItem in allMenus)
            {
                if(mainMenuItem.GetAttribute("Name").Equals("File"))
                {
                    mainMenuItem.Click();

                    var newMenu = mainMenuItem.FindElementByName("New");

                    wdv.Until(x => newMenu.Displayed);

                    newMenu.Click();

                    System.Threading.Thread.Sleep(400);

                    var lenderMenuThirdLevel = newMenu.FindElementByName("Third");

                    Actions actionForRightClick = new Actions(sessionWinForm);
                    actionForRightClick.MoveToElement(lenderMenuThirdLevel);
                    actionForRightClick.Click();
                    actionForRightClick.Perform();
                }
            }

        }

        [TestMethod]
        public void ListBoxBlindClickFailTest()
        {
            //lbStates
            var lbStates = sessionWinForm.FindElementByAccessibilityId("lbStates");

            var valueToClick = "NC";// this value can be read from a data source like an Excel file. 

            var listItemToClick = lbStates.FindElementByName(valueToClick);

            Debug.WriteLine($"{listItemToClick.Displayed} - {listItemToClick.Text}");

            listItemToClick.Click();
            // update
        }


        [TestMethod]
        public void ListBoxTest()
        {
            //lbStates
            var lbStates = sessionWinForm.FindElementByAccessibilityId("lbStates");

            var allListItems = lbStates.FindElementsByTagName("ListItem");

            var valueToClick = "NC";// this value can be read from a data source like an Excel file. 

            var maxClicks = 10;

            lbStates.SendKeys(Keys.Home);

            foreach(var ali in allListItems)
            {
                Debug.WriteLine($"{ali.Displayed} - {ali.Text}");

                if (ali.Text.Equals(valueToClick) && !ali.Displayed)
                {
                    var downButton = lbStates.FindElementByAccessibilityId("DownButton");
                    var listItemToClick = lbStates.FindElementByName(valueToClick);

                    while(!listItemToClick.Displayed && (maxClicks-- > 0))
                    {
                        downButton.Click();
                        listItemToClick = lbStates.FindElementByName(valueToClick);

                        if (listItemToClick.Displayed)
                        {
                            listItemToClick.Click();
                        }
                    }
                }
                else if (ali.Text.Equals(valueToClick) && ali.Displayed)
                {
                    ali.Click();
                }

            }

        }

        [TestMethod]
        public void ListBoxWithKeyboardTest()
        {
            //lbStates
            var lbStates = sessionWinForm.FindElementByAccessibilityId("lbStates");

            var allListItems = lbStates.FindElementsByTagName("ListItem");

            var valueToClick = "WA";// this value can be read from a data source like an Excel file. 

            var maxClicks = 10;

            lbStates.SendKeys(Keys.Home);

            foreach (var ali in allListItems)
            {
                Debug.WriteLine($"{ali.Displayed} - {ali.Text}");
                if (ali.Text.Equals(valueToClick) && !ali.Displayed)
                {
                    var listItemToClick = lbStates.FindElementByName(valueToClick);

                    while (!listItemToClick.Displayed && (maxClicks-- > 0))
                    {
                        lbStates.SendKeys(Keys.Down);
                        listItemToClick = lbStates.FindElementByName(valueToClick);

                        if (listItemToClick.Displayed)
                        {
                            listItemToClick.Click();
                        }
                    }
                }
                else if (ali.Text.Equals(valueToClick) && ali.Displayed)
                {
                    ali.Click();
                }

            }

        }

        [TestMethod]
        public void SliderTest()
        {
            var slider = sessionWinForm.FindElementByAccessibilityId("trackSlider");
            var thumb = slider.FindElementByName("Position");
            var btnCalculate = sessionWinForm.FindElementByName("Calculate ");

            Actions actDrag = new Actions(sessionWinForm);            
            //actDrag.DragAndDrop(thumb, btnCalculate);// to/from example, move way away from start                        
            //actDrag.DragAndDropToOffset(thumb, 0, 0);// move back to zero
            actDrag.DragAndDropToOffset(thumb, offsetX: slider.Rect.Width/2, 0);// slide half            
            //actDrag.Release();
            actDrag.Perform();


        }

    }
}
