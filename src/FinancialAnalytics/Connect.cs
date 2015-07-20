using System.Collections;
using System.Drawing;
using FinancialAnalytics.Bootstrapping;
using Microsoft.Win32;

namespace FinancialAnalytics
{
    using System;
    using Microsoft.Office.Core;
    using Core.Composition.Unity;

    #region Read me for Add-in installation and setup information.
    // When run, the Add-in wizard prepared the registry for the Add-in.
    // At a later time, if the Add-in becomes unavailable for reasons such as:
    //   1) You moved this project to a computer other than which is was originally created on.
    //   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
    //   3) Registry corruption.
    // you will need to re-register the Add-in by building the FinancialAnalyticsSetup project, 
    // right click the project in the Solution Explorer, then choose install.
    #endregion

    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    //[GuidAttribute("D1AE74F1-69EB-4FCE-808A-35A8C97690E6"), ProgId("FinancialAnalytics.Connect")]
    public class Connect : Object, Extensibility.IDTExtensibility2, IRibbonExtensibility
    {
        private readonly IServiceContainer _container;
        private CommonBootstrapper _bootstrapper;

        private string _udfProgId = "FinancialAnalytics.FormulaProcessor";
        private Microsoft.Office.Interop.Excel.Application _excel;
        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>
        public Connect()
        {
            _container = Locator.Current.Container;
        }

        public FinancialAnalytics.Bootstrapping.CommonBootstrapper CommonBootstrapper
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }

        /// <summary>
        ///      Implements the OnConnection method of the IDTExtensibility2 interface.
        ///      Receives notification that the Add-in is being loaded.
        /// </summary>
        /// <param term='application'>
        ///      Root object of the host application.
        /// </param>
        /// <param term='connectMode'>
        ///      Describes how the Add-in is being loaded.
        /// </param>
        /// <param term='addInInst'>
        ///      Object representing this Add-in.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
        {
            _excel = application as Microsoft.Office.Interop.Excel.Application;
            _bootstrapper = _container.GetInstance<CommonBootstrapper>();
            _bootstrapper.Run(application);
            EnableUDFAddin();
        }

        /// <summary>
        ///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
        ///     Receives notification that the Add-in is being unloaded.
        /// </summary>
        /// <param term='disconnectMode'>
        ///      Describes how the Add-in is being unloaded.
        /// </param>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref Array custom)
        {
            _bootstrapper.ApplicationProvider.Dispose();
            _bootstrapper.WindowManager.CloseAllWindows();
        }

        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>
        ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application has completed loading.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
        }

        public void RibbonLoaded(IRibbonUI ribbon)
        {
            _bootstrapper.Ribbon.RibbonLoaded(ribbon);
        }

        public string GetCustomUI(string RibbonID)
        {
            return ResourceX.RibbonExcel;
        }


        public string GetLabel(IRibbonControl control)
        {
            return _bootstrapper.Ribbon.GetLabel(control);
        }

        public bool GetEnabled(IRibbonControl control)
        {
            return _bootstrapper.Ribbon.GetEnabled(control);
        }

        public bool GetVisible(IRibbonControl control)
        {
            return _bootstrapper.Ribbon.GetVisible(control);
        }

        public string GetKeytip(IRibbonControl control)
        {
            return _bootstrapper.Ribbon.GetKeytip(control);
        }

        public string GetScreentip(IRibbonControl control)
        {
            return _bootstrapper.Ribbon.GetScreentip(control);
        }

        public string GetSupertip(IRibbonControl control)
        {
            return _bootstrapper.Ribbon.GetSupertip(control);
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            return _bootstrapper.Ribbon.GetImage(control);
        }

        public void OnButtonAction(IRibbonControl control)
        {
            _bootstrapper.Ribbon.OnButtonAction(control);
        }

        private void EnableUDFAddin()
        {
            try
            {
                IEnumerator addinsEnumerator = _excel.AddIns.GetEnumerator();
                Microsoft.Office.Interop.Excel.AddIn currentAddin = null;
                bool addinFound = false;

                while (addinsEnumerator.MoveNext())     // check if the Addin was already added
                {
                    currentAddin = addinsEnumerator.Current as Microsoft.Office.Interop.Excel.AddIn;
                    if (currentAddin.progID == _udfProgId)
                    {
                        currentAddin.Installed = false;  // disable it first so it works through automation
                        currentAddin.Installed = true;  // make sure it is enabled
                        addinFound = true;
                        break;
                    }
                }
                if (!addinFound)        // if addin wasn't added, add it
                {
                    _excel.AddIns.Add(_udfProgId, false).Installed = true;
                }
                
            }
            catch (Exception exception)
            { }
        }

        /// <summary>
        /// Creates a registry string or deletes one in the context of registering an Excel Automation Addin written in C#
        /// </summary>
        private void RegisterUDF(bool installTheAddin = true)
        {
            string totalname = _udfProgId;
            string totalnameReg = "/A " + "\"" + totalname + "\"";

            try
            {
                var excelVersion = _bootstrapper.ApplicationProvider.Application.Version;
                RegistryKey optionsKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Office\\" + excelVersion + "\\Excel\\Options", true);
                int numberValues = optionsKey.ValueCount;
                string[] valueNames = new string[numberValues];
                valueNames = optionsKey.GetValueNames();

                int numberAddins = 0;
                bool entryExists = false;
                string entryName = "";
                for (int i = 0; i <= numberValues - 1; i++)
                {
                    if (valueNames[i].Length >= 4)
                    {
                        string expression = valueNames[i].Substring(4);
                        int num;
                        bool isNumeric = int.TryParse(expression, out num);
                        if ((isNumeric) || (expression == ""))
                        {
                            numberAddins++;
                            if (String.Compare((string)optionsKey.GetValue(valueNames[i]), totalnameReg) == 0)
                            {
                                entryExists = true;
                                entryName = valueNames[i];
                            }
                        }
                    }
                }

                if (installTheAddin)
                {
                    if (!(entryExists))
                    {
                        string X = "";
                        if (numberAddins > 0) { X = numberAddins.ToString(); } else { X = ""; }
                        string stringvaluename = "OPEN" + X;
                        optionsKey.SetValue(stringvaluename, totalnameReg, RegistryValueKind.String);
                    }
                }
                else
                {
                    if (entryExists)
                    {
                        optionsKey.DeleteValue(entryName, false);
                    }
                }
            }
            catch (Exception ex)
            { }
        }
    }
}