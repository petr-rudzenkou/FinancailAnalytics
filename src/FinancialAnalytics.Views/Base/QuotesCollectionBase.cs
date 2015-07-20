using System.Collections.Generic;
using System.Linq;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Quotes;

namespace FinancialAnalytics.Views.Base
{
    public class QuotesCollectionBase : IQuotesCollectionBase
    {
        private readonly BindableCollection<QuotesData> _quotes = new BindableCollection<QuotesData>();
        public BindableCollection<QuotesData> Quotes
        {
            get { return _quotes; }
        }

        public virtual void Add(QuotesData quotesData)
        {
            if (!_quotes.Any(x => x.Symbol == quotesData.Symbol))
            {
                _quotes.Add(quotesData);
            }
        }

        public virtual void Add(IEnumerable<QuotesData> quotesData)
        {
            foreach (var q in quotesData)
            {
                if (!_quotes.Any(x => x.Symbol == q.Symbol))
                {
                    _quotes.Add(q);
                }
            }
        }

        public virtual void Remove(QuotesData quotesData)
        {
            var quotes = _quotes.FirstOrDefault(x => x.Symbol == quotesData.Symbol);
            if (quotes != null)
            {
                _quotes.Remove(quotes);
            }
        }

        public virtual void Remove(IEnumerable<QuotesData> quotesData)
        {
            foreach (var q in quotesData)
            {
                var quotes = _quotes.FirstOrDefault(x => x.Symbol == q.Symbol);
                if (quotes != null)
                {
                    _quotes.Remove(quotes);
                }
            }
        }

        public void Clear()
        {
            _quotes.Clear();
        }

        public virtual void Remove(string symbol)
        {
            var quotes = _quotes.FirstOrDefault(x => x.Symbol == symbol);
            if (quotes != null)
            {
                _quotes.Remove(quotes);
            }
        }

        public virtual void Remove(IEnumerable<string> symbols)
        {
            foreach (var s in symbols)
            {
                var quotes = _quotes.FirstOrDefault(x => x.Symbol == s);
                if (quotes != null)
                {
                    _quotes.Remove(quotes);
                }
            }
        }
    }
}
