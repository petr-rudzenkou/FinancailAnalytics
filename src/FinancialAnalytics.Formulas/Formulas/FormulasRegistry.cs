using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.RightsManagement;
using System.Text;

namespace FinancialAnalytics.Formulas.Formulas
{
    public class FormulasRegistry
    {
        private static readonly Dictionary<string, FormulaItem> _fomrulas = new Dictionary<string, FormulaItem>();

        public static bool Contains(string address)
        {
            return _fomrulas.ContainsKey(address);
        }

        public static void Register(string address, FormulaItem fomrulaItem)
        {
            _fomrulas.Add(address, fomrulaItem);
        }

        public static void UnRegister(string address)
        {
            if (Contains(address))
            {
                _fomrulas.Remove(address);
            }
        }

        public static FormulaItem GetFormulaItem(string address)
        {
            return _fomrulas[address];
        }
    }
}
