using System;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Vbe.Interfaces
{
    public interface IVBProject : IEntityWrapper<IVBProject>
    {
		IVBComponents VBComponents { get; }
    }
}
