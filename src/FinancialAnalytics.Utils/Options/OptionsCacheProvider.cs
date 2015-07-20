using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace FinancialAnalytics.Utils.Options
{
    public static class OptionsCacheProvider
    {
        private const string CacheFolder = @"%LOCALAPPDATA%\Financial Analytics\User Data\";
        private const string OptionsFileName = "User_Options.xml";

        private static readonly List<OptionBase> _options = new List<OptionBase>();

        static OptionsCacheProvider()
        {
            Initialize();
        }

        public static IEnumerable<OptionBase> Options
        {
            get { return _options; }
        }

        public static void Initialize()
        {
            var xmlFilePath = GetOptionsFilePath();
            try
            {
                if (File.Exists(xmlFilePath))
                {
                    XDocument document = XDocument.Load(xmlFilePath);
                    var nodes = document.Descendants("Option");

                    _options.Clear();
                    foreach (var node in nodes)
                    {
                        var nameAttr = node.Attribute("Name");
                        if (!_options.Any(x => x.Name.Equals(nameAttr.Value)))
                        {
                            switch (nameAttr.Value)
                            {
                                case OptionsConstants.DailyRefreshTime:
                                    {
                                        _options.Add(new DailyRefreshTimeOption()
                                        {
                                            Name = nameAttr.Value,
                                            DisplayName = node.Attribute("DisplayName").Value,
                                            DailyRefreshTime = node.Attribute("DailyRefreshTime").Value,
                                            IsSelected = Boolean.Parse(node.Attribute("IsSelected").Value)
                                        });
                                        break;
                                    }

                                case OptionsConstants.RefreshFrequency:
                                    {
                                        var refreshFrequencyMeasure = node.Attribute("RefreshFrequencyMeasure").Value;
                                        var refreshFrequency = int.Parse(node.Attribute("RefreshFrequency").Value);
                                        //int value;
                                        //switch (refreshFrequencyMeasure)
                                        //{
                                        //    case RefreshFrequencyMeasure.Sec:
                                        //        value = refreshFrequency * 1000;
                                        //        break;
                                        //    case RefreshFrequencyMeasure.Min:
                                        //        value = refreshFrequency * 1000 * 60;
                                        //        break;
                                        //    case RefreshFrequencyMeasure.Hours:
                                        //        value = refreshFrequency * 1000 * 60 * 60;
                                        //        break;
                                        //    case RefreshFrequencyMeasure.Days:
                                        //        value = refreshFrequency * 1000 * 60 * 60 * 24;
                                        //        break;
                                        //    default:
                                        //        value = 60000;
                                        //        break;
                                        //}
                                        _options.Add(new RefreshFrequencyOption()
                                        {
                                            Name = nameAttr.Value,
                                            DisplayName = node.Attribute("DisplayName").Value,
                                            RefreshFrequency = refreshFrequency,
                                            RefreshFrequencyMeasure = refreshFrequencyMeasure,
                                            IsSelected = Boolean.Parse(node.Attribute("IsSelected").Value)
                                        });
                                        break;
                                    }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var document = new XDocument(new XElement("Options"));
                document.Save(xmlFilePath);
            }
        }

        public static void Set(OptionBase option)
        {
            var name = option.Name;
            var displayName = option.DisplayName;

            var xmlFilePath = GetOptionsFilePath();
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
                document = new XDocument(new XElement("Options"));
                document.Save(xmlFilePath);
            }

            try
            {
                document = XDocument.Load(xmlFilePath);
            }
            catch (Exception ex)
            {
                document = new XDocument(new XElement("Options"));
            }

            if (document.Root == null)
                document.Add(new XElement("Options"));

            var root = document.Root;
            var nodes = root.Descendants("Option").ToList();

            var options = nodes.Where(x => x.Attribute("Name").Value.Equals(name)).ToList();
            if (options.Any())
            {
                options.Remove();
            }

            var optionNode = new XElement("Option");
            optionNode.Add(new XAttribute("Name", name));
            optionNode.Add(new XAttribute("DisplayName", displayName));
            optionNode.Add(new XAttribute("IsSelected", option.IsSelected));

            var dailyRefreshTimeOption = option as DailyRefreshTimeOption;
            if (dailyRefreshTimeOption != null)
            {
                optionNode.Add(new XAttribute("DailyRefreshTime", dailyRefreshTimeOption.DailyRefreshTime));
            }

            var refreshFrequencyOption = option as RefreshFrequencyOption;
            if (refreshFrequencyOption != null)
            {
                optionNode.Add(new XAttribute("RefreshFrequency", refreshFrequencyOption.RefreshFrequency));
                optionNode.Add(new XAttribute("RefreshFrequencyMeasure", refreshFrequencyOption.RefreshFrequencyMeasure));
            }

            root.Add(optionNode);

            document.Save(xmlFilePath);

        }

        public static void UnSet(OptionBase option)
        {
            var name = option.Name;

            var xmlFilePath = GetOptionsFilePath();
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
                document = new XDocument(new XElement("Options"));
                document.Save(xmlFilePath);
            }

            try
            {
                document = XDocument.Load(xmlFilePath);
            }
            catch (Exception ex)
            {
                document = new XDocument(new XElement("Options"));
            }

            if (document.Root == null)
                document.Add(new XElement("Options"));

            var root = document.Root;
            var nodes = root.Descendants("Option").ToList();

            var op = nodes.FirstOrDefault(x => x.Attribute("Name").Value.Equals(name));
            if (op != null)
            {
               op.Attribute("IsSelected").SetValue("false");
               document.Save(xmlFilePath);
            }
        }


        private static string GetOptionsFilePath()
        {
            string path = string.Empty;
            try
            {
                string localCacheFolder = Environment.ExpandEnvironmentVariables(CacheFolder);
                path = Path.Combine(localCacheFolder, OptionsFileName);
            }
            catch (Exception ex)
            { }
            return path;
        }
    }
}
