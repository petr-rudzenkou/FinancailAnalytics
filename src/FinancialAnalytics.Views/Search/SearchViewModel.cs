using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Caliburn.Micro;
using FinancialAnalytics.Views.Search.Interfaces;

namespace FinancialAnalytics.Views.Search
{
    public class SearchViewModel : Screen, ISearchViewModel
    {
        public SearchViewModel()
        {
            DisplayName = Resources.ViewsResources.Search_WindowTitle;
        }
    }
}
