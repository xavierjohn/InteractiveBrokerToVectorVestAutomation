using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using System.IO;
using System.Globalization;


namespace VV7
{
    class Program
    {
        const string VectorVestTitle = "VectorVest 7 US - Intraday";

        static public List<Trade> LoadTrades()
        {
           // IDataObject iData = Clipboard.GetDataObject();
            char[] delimiterChars = { '\t', '\t', '\t', '\t', '\t', '\t', '\t' };

            //// Determines whether the data is in a format you can use.
            //if (iData.GetDataPresent(DataFormats.Text))
            //{
            //    // Yes it is, so display it in a text box.
            //    strReader = new StringReader( (String)iData.GetData(DataFormats.Text));
            //}

            StreamReader strReader = new StreamReader("C:\\VV\\Trade.txt");
            string aLine;
            var lt = new List<Trade>();

            while (true)
            {
                aLine = strReader.ReadLine();
                if (aLine != null)
                {
                    string[] words = aLine.Split(delimiterChars);
                    var t = new Trade()
                    {
                        TradeType = null,
                        Symbol = words[2],
                        TradeDate = words[3],
                        Quantity = Int32.Parse(words[4], NumberStyles.AllowThousands, new CultureInfo("en-US")),
                        Price = Double.Parse(words[5]),
                        Commision = Double.Parse(words[6])
                    };

                    if (words[1] == "BOT" && words[7] != "Short")
                    {
                        t.TradeType = "Buy Long";
                    }
                    else if (words[1] == "BOT" && words[7] == "Short")
                    {
                        t.TradeType = "Cover Short";
                    }
                    else if (words[1] == "SLD" && words[7] != "Short")
                    {
                        t.TradeType = "Sell Long";
                    }
                    else if (words[1] == "SLD" && words[7] == "Short")
                    {
                        t.TradeType = "Sell Short";
                    }
                    else
                    {
                        throw new ArgumentException();
                    }

                    lt.Add(t);
                    Console.WriteLine(words[1] + " | " + words[2] + " | " + words[3] + " | " + words[4] + " | " + words[5] + " | " + words[6] + " | " + words[7]);
                }
                else
                {
                    break;
                }
            }
            return lt;
        }

        static bool openTradeWnd = true;
        static void Main(string[] args)
        {
            int count = 0;
            var lt = LoadTrades();
            var tradeTypes = new List<string>();
            tradeTypes.Add("Buy Long");
            tradeTypes.Add("Sell Short");
            tradeTypes.Add("Sell Long");
            tradeTypes.Add("Cover Short");

            if (args.Length > 0 && args[0] == "-D")
            {
                openTradeWnd = false;
            }
            AutomationElement clientAppRootInstance;
            do
            {
                ++count;
                //Get the Root Element of client application from the Desktop tree
                clientAppRootInstance = AutomationElement.RootElement.FindFirst(TreeScope.Children,
                                                                                new PropertyCondition(
                                                                                    AutomationElement.NameProperty,
                                                                                    VectorVestTitle));
                Thread.Sleep(100);
            } while (clientAppRootInstance == null && count < 10);

            if (clientAppRootInstance == null)
            {
                Console.WriteLine( VectorVestTitle + " not found!");
                return;
            }

            if (openTradeWnd)
            {
                AutomationElement buttonInstance = clientAppRootInstance.FindFirst(TreeScope.Descendants,
                                                                                   new PropertyCondition(
                                                                                       AutomationElement.
                                                                                           AutomationIdProperty,
                                                                                       "btnAddTrade"));
                PressButton(buttonInstance);
            }

            count = 0;
            AutomationElement placeTradeInstance = null;
            do
            {
                ++count;
                // Get place trade window
                placeTradeInstance = clientAppRootInstance.FindFirst(TreeScope.Descendants,
                                                                     new PropertyCondition(
                                                                         AutomationElement.NameProperty,
                                                                         "Place Trade"));
                Thread.Sleep(100);
            } while (placeTradeInstance == null && count < 10);

            if (placeTradeInstance == null)
            {
                Console.WriteLine("placeTradeInstance not found.");
                return;
            }

            AutomationElementCollection comboCollection = placeTradeInstance.FindAll(TreeScope.Children,
                                                                                     new PropertyCondition(
                                                                                         AutomationElement.
                                                                                             ControlTypeProperty,
                                                                                         ControlType.ComboBox));
            AutomationElementCollection editCollection = placeTradeInstance.FindAll(TreeScope.Descendants,
                                                                                     new PropertyCondition(
                                                                                         AutomationElement.
                                                                                             ControlTypeProperty,
                                                                                         ControlType.Edit));
            var symbolEdit = editCollection[0];
            var quantityEdit = editCollection[1];
            var priceEdit = editCollection[2]; 

            AutomationElementCollection buttonCollection = placeTradeInstance.FindAll(TreeScope.Children, new PropertyCondition(
                                                                                  AutomationElement.
                                                                                      ControlTypeProperty,
                                                                                  ControlType.Button));

            var applyButton = buttonCollection[3];
            var commissionBotton = buttonCollection[1];

            AutomationElement tradeTypeComboInstance = comboCollection[0];
            AutomationElement symbolComboInstance = comboCollection[1];
            if (tradeTypeComboInstance == null)
            {
                Console.WriteLine("tradeTypeComboInstance not found.");
                return;
            }

            foreach (var trade in lt)
            {
                tradeTypeComboInstance.SetFocus();
                ExpandCollapsePattern expandPattern = (ExpandCollapsePattern)tradeTypeComboInstance.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                expandPattern.Expand();
                Thread.Sleep(100);
                var cbItems = tradeTypeComboInstance.FindAll(TreeScope.Descendants, new PropertyCondition(
                                                                 AutomationElement.ControlTypeProperty,
                                                                 ControlType.ListItem));
                if (cbItems == null)
                {
                    Console.WriteLine("Trade Type combo not found");
                    return;
                }
                // Sell Long and Cover Short
                if (String.Equals(trade.TradeType, tradeTypes[2]) || String.Equals(trade.TradeType, tradeTypes[3]))
                {
                    if (trade.TradeType == tradeTypes[2])
                    {
                        SelectComboBox(cbItems[2]);
                    }
                    else
                    {
                        SelectComboBox(cbItems[3]);
                    }
                    symbolComboInstance.SetFocus();
                    expandPattern = (ExpandCollapsePattern)symbolComboInstance.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                    expandPattern.Expand();
                    Thread.Sleep(500);
                    cbItems = symbolComboInstance.FindAll(TreeScope.Descendants, new PropertyCondition(
                                                                                         AutomationElement.ControlTypeProperty,
                                                                                         ControlType.ListItem));

                    if (cbItems != null)
                    {
                        AutomationElement itemToSelect = null;
                        foreach (AutomationElement item in cbItems)
                        {
                            var str = GetListItemValue(item);
                            if (String.Equals(str, trade.Symbol))
                            {
                                itemToSelect = item;
                                break;
                            }
                        }
                        if (itemToSelect == null)
                        {
                            Console.WriteLine(trade.Symbol + " not found");
                            return;
                        }
                        SelectComboBox(itemToSelect);

                    }
                    expandPattern.Collapse();
                }

                // Buy Long and  Sell Short
                if ( trade.TradeType == tradeTypes[1] || trade.TradeType == tradeTypes[0])
                {
                    if (trade.TradeType == tradeTypes[1])
                    {
                        SelectComboBox(cbItems[1]);
                    }
                    else
                    {
                        SelectComboBox(cbItems[0]);
                    }
                    SetText(symbolEdit, trade.Symbol);
                }

                SetText(quantityEdit, trade.Quantity.ToString());  //quantity Edit control
                SetText(priceEdit, trade.Price.ToString());     // Price
                PressButton(commissionBotton);

                {  // Process the commission window
                    AutomationElement editCommissionInstance = null;
                    count = 0;
                    do
                    {
                        ++count;
                        // Get place trade window
                        editCommissionInstance = clientAppRootInstance.FindFirst(TreeScope.Descendants,
                                                                                 new PropertyCondition(
                                                                                     AutomationElement.NameProperty,
                                                                                     "Edit Commission"));
                        Thread.Sleep(100);
                    } while (editCommissionInstance == null && count < 10);
                    var controlCollection = editCommissionInstance.FindAll(TreeScope.Descendants, new PropertyCondition(
                                                                                                  AutomationElement.
                                                                                                      ControlTypeProperty,
                                                                                                  ControlType.Edit));
                    SetText(controlCollection[0], trade.Commision.ToString());
                    controlCollection = editCommissionInstance.FindAll(TreeScope.Descendants, new PropertyCondition(
                                                                                                  AutomationElement.
                                                                                                      AutomationIdProperty,
                                                                                                  "btnOK"));
                    PressButton(controlCollection[0]);
                }
                Thread.Sleep(100);
                PressButton(applyButton);
            }
        }

        static private String GetListItemValue(AutomationElement elementList)
        {
            // Set up the request.
            CacheRequest cacheRequest = new CacheRequest();

            // Get a full reference to the cached objects.
            cacheRequest.AutomationElementMode = AutomationElementMode.Full;

            // Cache all elements, regardless of whether they are control or content elements.
            cacheRequest.TreeFilter = Automation.RawViewCondition;

            // Cache the name property.
            cacheRequest.Add(AutomationElement.NameProperty);

            // Activate the request.
            cacheRequest.Push();

            // Obtain an element and cache the requested items.
            Condition cond = new AndCondition(Condition.TrueCondition, Condition.TrueCondition);
            AutomationElement elementListItem = elementList.FindFirst(TreeScope.Descendants, cond);

            // Deactivate the request.
            cacheRequest.Pop();

            // Retrieve the cached property and pattern.
            return elementListItem.Current.Name;
        }

        static private void SelectComboBox(AutomationElement elementList)
        {
            object pattern = null;
            if (elementList.TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern))
            {
                ((SelectionItemPattern)pattern).Select();
            }
            else
            {
                Console.WriteLine("Can't automate selection of combo box");
            }
        }

        static private void PressButton (AutomationElement elementList)
        {
            if (elementList == null)
            {
                Console.WriteLine("Button instance not found.");
                return;
            }
            InvokePattern buttonPattern = (InvokePattern)elementList.GetCurrentPattern(InvokePattern.Pattern);
            //Once get the pattern then calling Invoke method on that
            buttonPattern.Invoke();
        }

        static private void SetText(AutomationElement elementList, string text)
        {
            elementList.SetFocus();

            object valuePattern = null;

            if (elementList.TryGetCurrentPattern(
                ValuePattern.Pattern, out valuePattern))
            {
                ((ValuePattern)valuePattern).SetValue(text);
            }
        }
    }
}
