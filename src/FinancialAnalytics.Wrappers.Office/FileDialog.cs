using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	public class FileDialog : EntityWrapperBase<IFileDialog>, IFileDialog
	{
		protected Microsoft.Office.Core.FileDialog _officeFileDialog;

		public FileDialog(EntityResolverBase entityResolver, Microsoft.Office.Core.FileDialog fileDialog)
			: base(entityResolver)
		{
			if (fileDialog == null)
				throw new ArgumentNullException("fileDialog");
			_officeFileDialog = fileDialog;
		}

		#region Disposable pattern

		private bool disposed = false;

		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				if (disposing)
				{
					//Here we must dispose managed resources
				}
				//Here we must dispose unmanaged resources and LOH objects
				ComObjectsFinalizer.ReleaseComObject(_officeFileDialog);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion


		public bool AllowMultiSelect
		{
			get { return _officeFileDialog.AllowMultiSelect; }
			set { _officeFileDialog.AllowMultiSelect = value; }
		}

		public FileDialogType DialogType
		{
			get { return MsoFileDialogTypeToFileDialogTypeConverter.Convert(_officeFileDialog.DialogType); }
		}

		public int FilterIndex
		{
			get { return _officeFileDialog.FilterIndex; }
			set { _officeFileDialog.FilterIndex = value; }
		}

		public FileDialogFilters Filters
		{
			get { return _officeFileDialog.Filters; }
		}

		public string InitialFileName
		{
			get { return _officeFileDialog.InitialFileName; }
			set { _officeFileDialog.InitialFileName = value; }
		}

		public MsoFileDialogView InitialView
		{
			get { return _officeFileDialog.InitialView; }
			set { _officeFileDialog.InitialView = value; }
		}

		public string Item
		{
			get { return _officeFileDialog.Item; }
		}

		public FileDialogSelectedItems SelectedItems
		{
			get {  return _officeFileDialog.SelectedItems; }
		}

		public string Title
		{
			get { return _officeFileDialog.Title; }
			set { _officeFileDialog.Title = value; }
		}

		public int Show()
		{
			return _officeFileDialog.Show();
		}

		public override bool Equals(IFileDialog obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			FileDialog fileDialog = (FileDialog)obj;
			return _officeFileDialog.Equals(fileDialog._officeFileDialog);
		}


		public IEnumerable<string> SelectedFiles
		{
			get
			{
				for (int i = 1; i <= _officeFileDialog.SelectedItems.Count; i++)
				{
					yield return  _officeFileDialog.SelectedItems.Item(i);
				}

			}
		}


		public void AddFilter(string description, string extensions)
		{
			_officeFileDialog.Filters.Add(description, extensions);
		}

		public void ClearFilters()
		{
			_officeFileDialog.Filters.Clear();
		}
	}
}
