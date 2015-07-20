using System.Collections.Generic;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface IFileDialog : IEntityWrapper<IFileDialog>
	{
		bool AllowMultiSelect { get; set; }

		FileDialogType DialogType {get;}

		int FilterIndex { get; set; }

		FileDialogFilters Filters { get; }

		string InitialFileName { get; set; }

		MsoFileDialogView InitialView { get; set; }

		string Item { get; }

		FileDialogSelectedItems SelectedItems { get; }

		string Title { get; set; }

		IEnumerable<string> SelectedFiles { get; }

		void AddFilter(string description, string extensions);

		void ClearFilters();

		int Show();
	}
}