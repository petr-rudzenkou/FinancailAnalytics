using System;
using FinancialAnalytics.Wrappers.Office.EventsRouting;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
    public class CommandBars : EntitiesCollectionWrapperBase<ICommandBars, ICommandBar>, ICommandBars
    {
        protected Microsoft.Office.Core.CommandBars _officeCommandBars;
        protected EntityResolverBase EntityResolver { get; private set; }

        public CommandBars(EntityResolverBase entityResolver, Microsoft.Office.Core.CommandBars commandBars)
        {
            if (entityResolver == null)
            {
                throw new ArgumentNullException("entityResolver");
            }
           if (commandBars == null)
                throw new ArgumentNullException("commandBars");
            this.EntityResolver = entityResolver;
            _officeCommandBars = commandBars;
            InitializeCollection();
        }

        private void InitializeCollection()
        {
            for (int i = 1; i <= _officeCommandBars.Count; i ++)
            {
                AddItemToCollection(_officeCommandBars[i]);
            }
        }

        private void AddItemToCollection(Microsoft.Office.Core.CommandBar commandBar)
        {
            _items.Add(
                this.EntityResolver.ResolveCommandBar(commandBar)
                );
        }

        public ICommandBar this[string name]
        {
            get { return EntityResolver.ResolveCommandBar(_officeCommandBars[name]); }
        }

		public ICommandBarControl FindControl(object id)
		{
			Microsoft.Office.Core.CommandBarControl officeControl = _officeCommandBars.FindControl(Type.Missing, id, Type.Missing, Type.Missing);
			return officeControl == null ? null : EntityResolver.ResolveCommandBarControl(officeControl);
		}

        public void ExecuteMso(string idMso)
        {
            _officeCommandBars.ExecuteMso(idMso);
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
				RemoveEventHandlers();
				ComObjectsFinalizer.ReleaseComObject(_officeCommandBars);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public override bool Equals(ICommandBars obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            CommandBars commandBars = (CommandBars)obj;
            return _officeCommandBars.Equals(commandBars._officeCommandBars);
        }

		#region Events

		private void InitializeEventHandlers()
		{
			_officeCommandBars.OnUpdate += OfficeCommandBarsOnUpdate;
		}

		private void RemoveEventHandlers()
		{	
			_officeCommandBars.OnUpdate -= OfficeCommandBarsOnUpdate;
		}

		private void OfficeCommandBarsOnUpdate()
		{
			RaiseOnUpdateEvent();
		}

		private CommandBarsOnUpdate _onUpdate;
		public event CommandBarsOnUpdate OnUpdate
		{
			add
			{
				if (_onUpdate == null)
					InitializeEventHandlers();
				_onUpdate += value;
			}
			remove
			{	
				_onUpdate -= value;
				if (_onUpdate == null)
					RemoveEventHandlers();//remove subscribtion from native object to avoid frequent events to be raised
			}
		}

		protected void RaiseOnUpdateEvent()
		{
			if (_onUpdate != null)
			{
				_onUpdate();
			}
		}

		#endregion
	}
}
