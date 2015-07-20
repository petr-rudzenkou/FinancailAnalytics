using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using FinancialAnalytics.Wrappers.Office.Enums;
using Microsoft.Office.Core;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
    [ComVisible(true)]
    public interface IApplicationCommon : IDisposable
    {
        ICOMAddIns COMAddIns { get; }

        string Version { get; }

        bool IsStarted { get; }

        int WindowHandle { get; }

        IntPtr TopWindowHandle { get; }

        ICommandBars CommandBars { get; }

        bool Visible { get; set; }

        void InitializeApplication();

        void Quit();

        bool IsInitializedWithApplication { get; }

        bool AutoLaunch { get; set; }

        IEnumerable<string> AddInNames { get; }

        string Name { get; }

		ApplicationVersion ApplicationVersion { get; }

		IFileDialog GetFileDialog(FileDialogType fileDialogType);
    }
}
