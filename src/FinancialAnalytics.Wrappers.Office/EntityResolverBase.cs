using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	public class EntityResolverBase
	{
		public ICommandBars ResolveCommandBars(Microsoft.Office.Core.CommandBars officeCommandBars)
		{
			return new CommandBars(this, officeCommandBars);
		}

		public ICommandBar ResolveCommandBar(Microsoft.Office.Core.CommandBar officeCommandBar)
		{
			return new CommandBar(this, officeCommandBar);
		}

		public ICommandBarButton ResolveCommandBarButton(Microsoft.Office.Core.CommandBarButton officeCommandBarButton, bool wireEvents = true)
		{
			return new CommandBarButton(this, officeCommandBarButton, wireEvents);
		}

		public ICommandBarPopup ResolveCommandBarPopup(Microsoft.Office.Core.CommandBarPopup officeCommandBarPopup)
		{
			return new CommandBarPopup(this, officeCommandBarPopup);
		}

		/// <remarks>
		/// Makes cast-able wrapper, since casting direct wrapper, e.g. (CommandBarButton)commandBarControl, will fail.
		/// </remarks>
		public ICommandBarControl ResolveCommandBarControl(Microsoft.Office.Core.CommandBarControl officeCommandBarControl)
		{
			Microsoft.Office.Core.CommandBarButton officeCommandBarButton = officeCommandBarControl as Microsoft.Office.Core.CommandBarButton;
			if (officeCommandBarButton != null)
				return ResolveCommandBarButton(officeCommandBarButton);

			Microsoft.Office.Core.CommandBarPopup officeCommandBarPopup = officeCommandBarControl as Microsoft.Office.Core.CommandBarPopup;
			if (officeCommandBarPopup != null)
				return ResolveCommandBarPopup(officeCommandBarPopup);

			return new CommandBarControl(this, officeCommandBarControl);
		}

		public ICOMAddIn ResolveCOMAddIn(Microsoft.Office.Core.COMAddIn officeCOMAddIn)
		{
			return new COMAddIn(this, officeCOMAddIn);
		}

		public ICOMAddIns ResolveCOMAddIns(Microsoft.Office.Core.COMAddIns officeCOMAddIns)
		{
			return new COMAddIns(this, officeCOMAddIns);
		}

		public ICommandBarControls ResolveCommandBarControls(Microsoft.Office.Core.CommandBarControls commandBarControls)
		{
			return new CommandBarControls(this, commandBarControls);
		}

		public IGradientStop ResolveGradientStop(Microsoft.Office.Core.GradientStop gradientStop)
		{
			return new GradientStop(this, gradientStop);
		}

		public IGradientStops ResolveGradientStops(Microsoft.Office.Core.GradientStops gradientStops)
		{
			return new GradientStops(this, gradientStops);
		}

		public IColorFormat ResolveColorFormat(Microsoft.Office.Core.ColorFormat colorFormat)
		{
			return new ColorFormat(this, colorFormat);
		}

		public IFileDialog ResolveFileDialog(Microsoft.Office.Core.FileDialog fileDialog)
		{
			return new FileDialog(this, fileDialog);
		}

		public IDocumentProperty ResolveDocumentProperty(object documentProperty)
		{
			return new DocumentProperty(this, documentProperty);
		}

		public DocumentProperties ResolveDocumentProperties(object documentProperties)
		{
			return new DocumentProperties(this, documentProperties);
		}

		public ICustomXmlPart ResolveCustomXmlPart(Microsoft.Office.Core.CustomXMLPart customXmlPart)
		{
			if (customXmlPart == null)
			{
				return null;
			}
			return new CustomXmlPart(this, customXmlPart);
		}

		public ICustomXmlParts ResolveCustomXmlParts(Microsoft.Office.Core.CustomXMLParts customXmlParts)
		{
			if (customXmlParts == null)
			{
				return null;
			}
			return new CustomXmlParts(this, customXmlParts);
		}

		public ICustomXmlPrefixMappings ResolveCustomXmlPrefixMappings(Microsoft.Office.Core.CustomXMLPrefixMappings customXmlPrefixMappings)
		{
			if (customXmlPrefixMappings == null)
			{
				return null;
			}
			return new CustomXmlPrefixMappings(this, customXmlPrefixMappings);
		}

		public ICustomXmlNode ResolveCustomXmlNode(Microsoft.Office.Core.CustomXMLNode customXmlNode)
		{
			if (customXmlNode == null)
			{
				return null;
			}
			return new CustomXmlNode(this, customXmlNode);
		}

		public ICustomXmlNodes ResolveCustomXmlNodes(Microsoft.Office.Core.CustomXMLNodes customXmlNodes)
		{
			if (customXmlNodes == null)
			{
				return null;
			}
			return new CustomXmlNodes(this, customXmlNodes);
		}

		public IFont2 ResolveFont2(Microsoft.Office.Core.Font2 font2)
		{
			if (font2 == null)
			{
				return null;
			}
			return new Font2(this, font2);
		}

		public ITextRange2 ResolveTextRange2(Microsoft.Office.Core.TextRange2 textRange2)
		{
			if (textRange2 == null)
			{
				return null;
			}
			return new TextRange2(this, textRange2);
		}
	}
}