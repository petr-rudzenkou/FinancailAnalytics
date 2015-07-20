using System;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Converters;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	internal class CommandBarControls : EntitiesCollectionWrapperBase<ICommandBarControls, ICommandBarControl>, ICommandBarControls
	{
		private readonly Microsoft.Office.Core.CommandBarControls _officeCommandBarControls;

		protected EntityResolverBase EntityResolver { get; private set; }

		public CommandBarControls(EntityResolverBase entityResolver, Microsoft.Office.Core.CommandBarControls officeCommandBarControls)
		{
			if (entityResolver == null)
			{
				throw new ArgumentNullException("entityResolver");
			}
			if (officeCommandBarControls == null)
			{
				throw new ArgumentNullException("officeCommandBarControls");
			}
			this.EntityResolver = entityResolver;
			_officeCommandBarControls = officeCommandBarControls;
			InitializeCollection();
		}

		private void InitializeCollection()
		{
			for (int i = 1; i <= _officeCommandBarControls.Count; i++)
			{
				AddItemToCollection(_officeCommandBarControls[i]);
			}
		}

		private void AddItemToCollection(Microsoft.Office.Core.CommandBarControl commandBarControl)
		{
			_items.Add(
				EntityResolver.ResolveCommandBarControl(commandBarControl)
				);
		}

		public override bool Equals(ICommandBarControls obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			CommandBarControls commandBarControls = (CommandBarControls)obj;
			return _officeCommandBarControls.Equals(commandBarControls._officeCommandBarControls);
		}

		private bool _disposed;

		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				if (disposing)
				{
					//Here we must dispose managed resources
				}
				//Here we must dispose unmanaged resources and LOH objects
				ComObjectsFinalizer.ReleaseComObject(_officeCommandBarControls);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		public ICommandBarControl this[object index]
		{
			get { return EntityResolver.ResolveCommandBarControl(_officeCommandBarControls[index]); }
		}

		public ICommandBarControl Add(
			ControlType type,
			object id,
			object parameter,
			object before,
			object temporary)
		{
			MsoControlType controlType = MsoControlTypeToControlTypeConverter.ConvertBack(type);
			Microsoft.Office.Core.CommandBarControl officeCommandBarControl = _officeCommandBarControls.Add(controlType, id, parameter, before, temporary);
			AddItemToCollection(officeCommandBarControl);
			return EntityResolver.ResolveCommandBarControl(officeCommandBarControl);
		}
	}
}