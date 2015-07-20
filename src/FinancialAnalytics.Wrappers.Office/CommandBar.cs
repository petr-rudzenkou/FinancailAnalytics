using System;
using System.Runtime.InteropServices;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Converters;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	public class CommandBar : EntityWrapperBase<ICommandBar>, ICommandBar
	{
		protected Microsoft.Office.Core.CommandBar _officeCommandBar;
		private static readonly Object _locker = new object();

		public CommandBar(EntityResolverBase entityResolver, Microsoft.Office.Core.CommandBar commandBar)
			: base(entityResolver)
		{
			if (commandBar == null)
				throw new ArgumentNullException("commandBar");
			_officeCommandBar = commandBar;
		}

		public ICommandBarControl FindControl(ControlType type, object id, object tag, object visible, bool recursive)
		{
			MsoControlType msoControlType = MsoControlTypeToControlTypeConverter.ConvertBack(type);
			Microsoft.Office.Core.CommandBarControl officeControl = _officeCommandBar.FindControl(msoControlType, id, tag, visible, recursive);
			return officeControl == null ? null : EntityResolver.ResolveCommandBarControl(officeControl);
		}

		public void Reset()
		{
			_officeCommandBar.Reset();
		}

		public void Delete()
		{
			_officeCommandBar.Delete();
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
				if (_officeCommandBar != null && Marshal.IsComObject(_officeCommandBar))
				{
					lock (_locker)
					{
						ComObjectsFinalizer.ReleaseComObject(_officeCommandBar);
						_officeCommandBar = null;
					}
				}
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		public override bool Equals(ICommandBar obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			CommandBar commandBar = (CommandBar)obj;
			return _officeCommandBar.Equals(commandBar._officeCommandBar);
		}

		public bool Enabled
		{
			get { return _officeCommandBar.Enabled; }
			set { _officeCommandBar.Enabled = value; }
		}

		public bool Visible
		{
			get { return _officeCommandBar.Visible; }
			set { _officeCommandBar.Visible = value; }
		}
		
        public ICommandBarControls Controls
        {
            get { return EntityResolver.ResolveCommandBarControls(_officeCommandBar.Controls); }
        }

		public int Id
		{
			get { return _officeCommandBar.Id; }
		}
	}
}