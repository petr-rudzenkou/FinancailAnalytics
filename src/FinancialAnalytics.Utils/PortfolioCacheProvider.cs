using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace FinancialAnalytics.Utils
{
    public static class PortfolioCacheProvider
    {
        private const string CacheFolder = @"%LOCALAPPDATA%\Financial Analytics\User Data\";
        private const string PortfolioFileName = "User_Portfolio.xml";

        private readonly static List<string> _portfolioSymbols = new List<string>();

        static PortfolioCacheProvider()
        {
            Initialize();
        }

        public static IEnumerable<string> PortfolioSymbols
        {
            get { return _portfolioSymbols; }
        }

        private static void Initialize()
        {
            var portfolioXmlFile = GetPortfolioFilePath();
            try
            {
                if (File.Exists(portfolioXmlFile))
                {
                    XDocument document = XDocument.Load(portfolioXmlFile);
                    var nodes = document.Descendants("Symbol");

                    foreach (var node in nodes)
                    {
                        if (!_portfolioSymbols.Contains(node.Value))
                        {
                            _portfolioSymbols.Add(node.Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var document = new XDocument(new XElement("Portfolio"));
                document.Save(portfolioXmlFile);
            }
        }

        public static void Add(string symbol)
        {
            var xmlFilePath = GetPortfolioFilePath();
            if (string.IsNullOrEmpty(xmlFilePath))
                return;

            var xmlFileDirectory = Path.GetDirectoryName(xmlFilePath);
            if (!Directory.Exists(xmlFileDirectory))
            {
                Directory.CreateDirectory(xmlFileDirectory);
            }

            XDocument document;
            //Create file if it doesn't exist
            if (!File.Exists(xmlFilePath))
            {
                document = new XDocument(new XElement("Portfolio"));
                document.Save(xmlFilePath);
            }

            try
            {
                document = XDocument.Load(xmlFilePath);
            }
            catch (Exception ex)
            {
                document = new XDocument(new XElement("Portfolio"));
            }

            if (document.Root == null)
                document.Add(new XElement("Portfolio"));

            var root = document.Root;
            var nodes = root.Descendants("Symbol").ToList();

            var items = nodes.Where(x => x.Value == symbol).ToList();
            if (items.Count > 0)
                return;

            var portfolioNode = new XElement("Symbol", symbol);
            root.Add(portfolioNode);
            document.Save(xmlFilePath);

            if (!_portfolioSymbols.Contains(symbol))
            {
                _portfolioSymbols.Add(symbol);
            }
        }

        public static void Remove(string symbol)
        {
            var xmlFilePath = GetPortfolioFilePath();
            if (File.Exists(xmlFilePath))
            {
                XDocument document;
                try
                {
                    document = XDocument.Load(xmlFilePath);
                }
                catch (Exception ex)
                {
                    document = new XDocument(new XElement("Portfolio"));
                }

                var root = document.Root;
                if (root != null)
                {
                    var items = root.Descendants("Symbol").Where(x => x.Value == symbol).ToList();
                    if (items.Count > 0)
                    {
                        items.Remove();
                    }
                    document.Save(xmlFilePath);
                }
            }
            if (_portfolioSymbols.Contains(symbol))
            {
                _portfolioSymbols.Remove(symbol);
            }
        }

        private static string GetPortfolioFilePath()
        {
            string path = string.Empty;
            try
            {
                string localCacheFolder = Environment.ExpandEnvironmentVariables(CacheFolder);
                path = Path.Combine(localCacheFolder, PortfolioFileName);
            }
            catch (Exception ex)
            {
            }
            return path;
        }
    }
} 
