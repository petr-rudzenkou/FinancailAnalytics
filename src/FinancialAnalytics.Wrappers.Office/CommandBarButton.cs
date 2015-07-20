using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.EventsRouting;
using FinancialAnalytics.Wrappers.Office.Interfaces;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Converters;

namespace FinancialAnalytics.Wrappers.Office
{
	public class CommandBarButton : CommandBarControl, ICommandBarButton
	{
		private readonly bool _wireEvents;
		private readonly Microsoft.Office.Core.CommandBarButton _officeCommandBarButton;
		public const string CommandBarButtonEventsGuid = "000C0351-0000-0000-C000-000000000046";

		private readonly int _cookie;

		public CommandBarButton(EntityResolverBase entityResolver, Microsoft.Office.Core.CommandBarButton commandBarButton, bool wireEvents = true)
			: base(entityResolver, commandBarButton)
		{
			_wireEvents = wireEvents;
			_officeCommandBarButton = commandBarButton;
			if (_wireEvents)
			{
				_cookie = ConnectEvents(new Guid(CommandBarButtonEventsGuid), _officeCommandBarControl, this);
			}
		}

		private static int ConnectEvents(Guid guid, object source, object target)
		{
			IConnectionPointContainer connPointContainer = (IConnectionPointContainer)source;
			IConnectionPoint connectionPoint;
			connPointContainer.FindConnectionPoint(ref guid, out connectionPoint);
			int cookie;
			connectionPoint.Advise(target, out cookie);
			return cookie;
		}

		private static void DisconnectEvents(Guid guid, object source, int cookie)
		{
			IConnectionPointContainer connPointContainer;
			try
			{
				connPointContainer = source as IConnectionPointContainer;
			}
			catch (InvalidComObjectException) // during application closing (object is disposed)
			{
				return;
			}
            if (connPointContainer == null)
            {
                return;
            }

			try
			{
				IConnectionPoint connectionPoint;
				connPointContainer.FindConnectionPoint(ref guid, out connectionPoint);
				connectionPoint.Unadvise(cookie);
			}
			catch
			{

			}
		}

		public event _CommandBarButtonEvents_ClickEventHandler Click;

		public void OnClick(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
		{
			var evt = Click;
			if (evt != null)
			{
				evt(this, ref CancelDefault);
			}
		}

		public Bitmap Picture
		{
			get { return PictureDispConverter.Convert(_officeCommandBarButton.Picture); }
			set { _officeCommandBarButton.Picture = PictureDispConverter.Convert(value); }
		}

		public Bitmap Mask
		{
			get { return PictureDispConverter.Convert(_officeCommandBarButton.Mask); }
			set { _officeCommandBarButton.Mask = PictureDispConverter.Convert(value); }
		}

		public CommandBarButtonState State
		{
			get { return MsoButtonStateToCommandBarButtonStateConverter.Convert(_officeCommandBarButton.State); }
			set { _officeCommandBarButton.State = MsoButtonStateToCommandBarButtonStateConverter.ConvertBack(value); }
		}

		public void Execute()
		{
			_officeCommandBarButton.Execute();
		}

		public void TurnOnButton()
		{
			if (State == CommandBarButtonState.ButtonUp)
			{
				Execute();
			}
		}

		public void TurnOffButton()
		{
			if (State == CommandBarButtonState.ButtonDown)
			{
				Execute();
			}
		}

		#region Disposable pattern

		private bool disposed;

		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				if (_wireEvents)
				{
					DisconnectEvents(new Guid(CommandBarButtonEventsGuid), _officeCommandBarControl, _cookie);
				}
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion
	}
}