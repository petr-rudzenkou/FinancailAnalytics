using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	public class CommandBarControl : EntityWrapperBase<ICommandBarControl>, ICommandBarControl
	{
		protected Microsoft.Office.Core.CommandBarControl _officeCommandBarControl;

		public CommandBarControl(EntityResolverBase entityResolver, Microsoft.Office.Core.CommandBarControl commandBarControl)
			: base(entityResolver)
		{
			if (commandBarControl == null)
				throw new ArgumentNullException("commandBarControl");
			_officeCommandBarControl = commandBarControl;
		}

		public bool Enabled
		{
			get { return _officeCommandBarControl.Enabled; }
			set { _officeCommandBarControl.Enabled = value; }
		}

		public void Delete(object temporary)
		{
			_officeCommandBarControl.Delete(temporary);
		}

		public bool Visible
		{
			get { return _officeCommandBarControl.Visible; }
			set { _officeCommandBarControl.Visible = value; }
		}

		public string Caption
		{
			get { return _officeCommandBarControl.Caption; }
			set { _officeCommandBarControl.Caption = value; }
		}

		public string Tag
		{
			get { return _officeCommandBarControl.Tag; }
			set { _officeCommandBarControl.Tag = value; }
		}

	    public int Index
	    {
            get { return _officeCommandBarControl.Index; }
	    }

	    public int Id
	    {
	        get { return _officeCommandBarControl.Id; }
	    }

        public int ListCount
        {
            get
            {
                try
                {
                    dynamic obj = _officeCommandBarControl;
                    object listCount = obj.ListCount;
                    if (listCount != null)
                    {
                        return (int)listCount;
                    }
                    else
                    {
                        return 0;
                    }
                }
                catch (Exception)
                {
                    return 0;
                }
            }
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
				ComObjectsFinalizer.ReleaseComObject(_officeCommandBarControl);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		public override bool Equals(ICommandBarControl obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			CommandBarControl commandBarControl = (CommandBarControl)obj;
			return _officeCommandBarControl.Equals(commandBarControl._officeCommandBarControl);
		}
	}
}
