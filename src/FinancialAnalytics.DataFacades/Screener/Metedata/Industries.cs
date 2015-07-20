using System.Collections.Generic;
using System.Linq;

namespace FinancialAnalytics.DataFacades.Screener.Metedata
{
    public static class Industries
    {
        public static readonly Industry Any = new Industry("0", "Any", "Any");
        public static readonly Industry AccidentAndHelathInsurance = new Industry("431", "Accident & Helath Insurance (Financial)", "Accident & Helath Insurance (Financial)");
        public static readonly Industry AdvertisingAgencies = new Industry("720", "Advertising Agencies", "Advertising Agencies");
        public static readonly Industry AutoManufacturersMagor = new Industry("330", "Auto Manufacturers - Major", "Auto Manufacturers - Major");
        public static readonly Industry Biotechnology = new Industry("515", "Biotechnology", "Biotechnology");
        public static readonly Industry BroadcastingTv = new Industry("723", "Broadcasting - TV", "Broadcasting - TV");
        public static readonly Industry BusinessSoftwareAndServices = new Industry("826", "Business Software & Services", "Business Software & Services");
        //public static readonly Industry Cigarettes = new Industry("350", "Cigarettes", "Cigarettes");
        public static readonly Industry ChemicalsMajorDiversified = new Industry("110", "Chemicals - Major Diversified", "Chemicals - Major Diversified");
        public static readonly Industry ComputerBasedSystems = new Industry("812", "Computer Based Systems", "Computer Based Systems");
        public static readonly Industry CreditServices = new Industry("424", "Credit Services", "Credit Services");
        //public static readonly Industry DrugManufacturersMajor = new Industry("510", "Drug Manufacturers - Major", "Drug Manufacturers - Major");
        public static readonly Industry EducationAndTrainingServices = new Industry("766", "Education & Training Services", "Education & Training Services");
        public static readonly Industry ElectronicEquipment = new Industry("314", "Electronic Equipment", "Electronic Equipment");
        public static readonly Industry Hospitals = new Industry("524", "Hospitals", "Hospitals");
        public static readonly Industry InternetInformationProviders = new Industry("851", "Internet Information Providers", "Internet Information Providers");
        public static readonly Industry TechnicalAndSystemSoftware = new Industry("822", "Technical & System Software", "Technical & System Software");
        public static readonly Industry MusicAndVideoStores = new Industry("743", "Music & Video Stores", "Music & Video Stores");
        public static readonly Industry MoneyCenterBanks = new Industry("410", "Money Center Banks", "Money Center Banks");
        public static readonly Industry BusinessServices = new Industry("760", "Business Services", "Business Services");
        public static readonly Industry OilAndGasRefiningAndMarketing = new Industry("122", "Oil & Gas Refining & Marketing", "Oil & Gas Refining & Marketing");
        public static readonly Industry Restaurants = new Industry("712", "Restaurants", "Restaurants");
        public static readonly Industry TobaccoProductsAndOther = new Industry("351", "Tobacco Products, Other", "Tobacco Products, Other");
        public static readonly Industry WirelessCommunications = new Industry("840", "Wireless Communications", "Wireless Communications");
        //public static readonly Industry WaterUtilities = new Industry("914", "Water Utilities", "Water Utilities");

        public static readonly List<Industry> All = new List<Industry>
                                {
                                    Any,
                                    AccidentAndHelathInsurance,
                                    AdvertisingAgencies,
                                    AutoManufacturersMagor,
                                    Biotechnology,
                                    BroadcastingTv,
                                    BusinessSoftwareAndServices,
                                    //Cigarettes,
                                    ChemicalsMajorDiversified,
                                    ComputerBasedSystems,
                                    CreditServices,
                                    //DrugManufacturersMajor,
                                    EducationAndTrainingServices,
                                    ElectronicEquipment,
                                    Hospitals,
                                    InternetInformationProviders,
                                    TechnicalAndSystemSoftware,
                                    MusicAndVideoStores,
                                    MoneyCenterBanks,
                                    BusinessServices,
                                    OilAndGasRefiningAndMarketing,
                                    Restaurants,
                                    TobaccoProductsAndOther,
                                    WirelessCommunications
                                    //WaterUtilities
                                };

        public static Industry ById(string id)
        {
            return All.FirstOrDefault(x => x.Id == id);
        }
    }
}
