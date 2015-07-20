using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.DataFacades.XChangeRates.Metadata;

namespace FinancialAnalytics.DataFacades
{
    public class ImportExport
    {
        public static QuotesData[] XmlToQuoteData(string xmlString, System.Globalization.CultureInfo culture = null)
        {
            var quotesDataResult = new List<QuotesData>();
            try
            {
                Type type = typeof (QuotesData);
                var properties = type.GetProperties();
                var stringReader = new StringReader(xmlString);
                var document = XDocument.Load(stringReader);
                var root = document.Root;
                if (root != null)
                {
                    var nodes = root.Descendants("quote").ToList();
                    for (int i = 0; i < nodes.Count; i++)
                    {
                        var quote = new QuotesData();
                        for (int j = 0; j < properties.Length; j++)
                        {
                            type.InvokeMember(properties[j].Name, BindingFlags.SetProperty, null, quote,
                                new[] { FindChildElement(nodes[i], properties[j].Name).Value });
                        }
                        quotesDataResult.Add(quote);
                    }

                }
            }
            catch (Exception ex)
            { }
            
            return quotesDataResult.ToArray();
        }

        public static HistoricalData.HistoricalData[] XmlToHistoricalData(string xmlString, System.Globalization.CultureInfo culture = null)
        {
            var historicalDataResult = new List<HistoricalData.HistoricalData>();
            try
            {
                Type type = typeof(HistoricalData.HistoricalData);
                var properties = typeof(HistoricalData.HistoricalData).GetProperties();
                var stringReader = new StringReader(xmlString);
                var document = XDocument.Load(stringReader);
                var root = document.Root;
                if (root != null)
                {
                    var nodes = root.Descendants("quote").ToList();
                    for (int i = 0; i < nodes.Count; i++)
                    {
                        var historicalData = new HistoricalData.HistoricalData();
                        for (int j = 0; j < properties.Length; j++)
                        {
                            var propertyName = properties[j].Name;
                            object parameter;
                            if (propertyName.Equals("Symbol"))
                            {
                                parameter = GetAttributeValue(nodes[i], "Symbol");
                            }
                            else
                            {
                                var element = FindChildElement(nodes[i], properties[j].Name);
                                parameter = element != null ? element.Value : string.Empty;
                            }
                            type.InvokeMember(propertyName, BindingFlags.SetProperty, null, historicalData,
                                new[] { parameter });
                        }
                        historicalDataResult.Add(historicalData);
                    }

                }
            }
            catch (Exception ex)
            { }

            return historicalDataResult.ToArray();
        }

        public static XChangeRate[] XmlToXChangeRate(string xmlString, System.Globalization.CultureInfo culture = null)
        {
            var xChangeRatesResult = new List<XChangeRate>();
            try
            {
                Type type = typeof(XChangeRate);
                var properties = type.GetProperties();
                var stringReader = new StringReader(xmlString);
                var document = XDocument.Load(stringReader);
                var root = document.Root;
                if (root != null)
                {
                    var nodes = root.Descendants("rate").ToList();
                    for (int i = 0; i < nodes.Count; i++)
                    {
                       var xChangeRate = new XChangeRate();
                        for (int j = 0; j < properties.Length; j++)
                        {
                            var propertyName = properties[j].Name;
                            object parameter;
                            if (propertyName.Equals("Id"))
                            {
                                parameter = GetAttributeValue(nodes[i], "id");
                            }
                            else
                            {
                                var element = FindChildElement(nodes[i], properties[j].Name);
                                parameter = element != null ? element.Value : string.Empty;
                            }
                            type.InvokeMember(propertyName, BindingFlags.SetProperty, null, xChangeRate,
                                new[] {parameter});
                        }
                        xChangeRatesResult.Add(xChangeRate);
                    }

                }
            }
            catch (Exception ex)
            { }

            return xChangeRatesResult.ToArray();
        }

        #region Tools
        private static XElement FindChildElement(XElement parentElement, string nodeName)
        {
            XElement result = null;
            try
            {
                if (parentElement != null)
                {
                    result = parentElement.Element(nodeName);
                }
            }
            catch (Exception ex)
            { }
            return result;
        }

        private static string GetAttributeValue(XElement element, string attributeName)
        {
            var attribute = element.Attribute(attributeName);
            return attribute != null ? attribute.Value : string.Empty;
        }
        #endregion

    }
}
