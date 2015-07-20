using System;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Resources;

namespace FinancialAnalytics.Wrappers.Excel.Converters.Localization
{
	public class LocalizedDescriptionAttribute : DescriptionAttribute
	{
		private readonly string _resourceKey;
		private readonly ResourceManager _resource;
		public LocalizedDescriptionAttribute(string resourceKey, Type resourceType)
		{
			_resource = new ResourceManager(resourceType);
			_resourceKey = resourceKey;
		}

		public override string Description
		{
			get
			{
				string displayName = _resource.GetString(_resourceKey, CultureInfo.CurrentUICulture);

				if (String.IsNullOrEmpty(displayName))
				{
					return String.Format("[[{0}]]", _resourceKey);
				}
				else
				{
					return displayName;
				}
			}
		}
	}

	public static class EnumExtensions
	{
		public static string GetDescription(this Enum enumValue)
		{
			FieldInfo fieldInfo = enumValue.GetType().GetField(enumValue.ToString());

			DescriptionAttribute[] attributes = (DescriptionAttribute[])fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);

			if (attributes.Length > 0)
			{
				return attributes[0].Description;
			}
			else
			{
				return enumValue.ToString();
			}
		}
	}
}
