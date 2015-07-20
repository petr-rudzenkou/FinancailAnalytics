using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Caliburn.Micro;
using FinancialAnalytics.Views.LeagueTable.Interfaces;

namespace FinancialAnalytics.Views.LeagueTable
{
    public class LeagueTableViewModel : Screen, ILeagueTableViewModel
    {
        public LeagueTableViewModel()
        {
            DisplayName = Resources.ViewsResources.LeagueTable_WindowTitle;
        }
    }
}
