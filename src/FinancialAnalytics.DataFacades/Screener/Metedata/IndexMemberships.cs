using System.Collections.Generic;
using System.Linq;

namespace FinancialAnalytics.DataFacades.Screener.Metedata
{
    public class IndexMemberships
    {
        public static IndexMembership Any = new IndexMembership("", "Any", "Any");
        public static IndexMembership DowJonesIndustirals = new IndexMembership("%5EDJI", "Dow Jones Industirals", "Dow Jones Industirals");
        public static IndexMembership DowJonesTransportation = new IndexMembership("%5EDJT", "Dow Jones Transportation", "Dow Jones Transportation");
        public static IndexMembership DowJonesUtilities = new IndexMembership("%5EDJU", "Dow Jones Utilities", "Dow Jones Utilities");
        public static IndexMembership SP500 = new IndexMembership("%5ESPC", "S&P 500", "S&P 500");
        public static IndexMembership SP400MidCap = new IndexMembership("%5EMID", "S&P 400 MidCap", "S&P 400 MidCap");
        public static IndexMembership SP600SmallCap = new IndexMembership("%5ESML", "S&P 600 SmallCap", "S&P 600 SmallCap");

        public static readonly List<IndexMembership> All = new List<IndexMembership>
								{
									Any,
									DowJonesIndustirals,
									DowJonesTransportation,
									DowJonesUtilities,
									SP500,
									SP400MidCap,
									SP600SmallCap,
								};

        public static IndexMembership ById(string id)
        {
            return All.FirstOrDefault(x => x.Id == id);
        }
    }
}
