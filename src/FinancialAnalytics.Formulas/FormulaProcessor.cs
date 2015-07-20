using System;
using System.Runtime.InteropServices;
using FinancialAnalytics.Core.Composition.Unity;
using FinancialAnalytics.Formulas.Bootstrapping;

namespace FinancialAnalytics.Formulas
{
    [Guid("12AEEA34-BE04-48B2-BBC0-22DB512E96AA")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    [ProgId("FinancialAnalytics.FormulaProcessor")]
    public class FormulaProcessor : Extensibility.IDTExtensibility2, IFormulaProcessor
    {
        private readonly IServiceContainer _container;
        private FormulasBootstrapper _bootstrapper;

        public FormulaProcessor()
        {
            _container = Locator.Current.Container;
        }

        public object FA(object symbols, [OptionalAttribute]object dataItems, [OptionalAttribute]object layout, [OptionalAttribute]object destinationCell)
        {
            return _bootstrapper.FormulaHandler.FA(symbols, dataItems, layout, destinationCell);
        }

        #region IDTExtensibility2
        public void OnAddInsUpdate(ref Array custom)
        {
            
        }

        public void OnBeginShutdown(ref Array custom)
        {
          
        }

        public void OnConnection(object Application, Extensibility.ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _bootstrapper = _container.GetInstance<FormulasBootstrapper>();
            _bootstrapper.Run(Application);
        }

        public void OnDisconnection(Extensibility.ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _bootstrapper.ApplicationProvider.Dispose();
        }

        public void OnStartupComplete(ref Array custom)
        {
           
        }
        #endregion
    }
}
