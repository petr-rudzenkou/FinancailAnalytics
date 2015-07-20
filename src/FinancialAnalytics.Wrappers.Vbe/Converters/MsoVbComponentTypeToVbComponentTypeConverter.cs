using Microsoft.Vbe.Interop;
using FinancialAnalytics.Wrappers.Vbe.Enums;

namespace FinancialAnalytics.Wrappers.Vbe.Converters
{
	public class MsoVbComponentTypeToVbComponentTypeConverter
    {
		public static VbComponentType Convert(vbext_ComponentType msoComponentType)
        {
			VbComponentType componentType;
			switch (msoComponentType)
            {
				case vbext_ComponentType.vbext_ct_ActiveXDesigner:
					componentType = VbComponentType.ActiveXDesigner;
                    break;
				case vbext_ComponentType.vbext_ct_ClassModule:
					componentType = VbComponentType.ClassModule;
                    break;
				case vbext_ComponentType.vbext_ct_Document:
					componentType = VbComponentType.Document;
                    break;
				case vbext_ComponentType.vbext_ct_MSForm:
					componentType = VbComponentType.MSForm;
                    break;
				default:
					componentType = VbComponentType.StdModule;
                    break;
            }
			return componentType;
        }

		public static vbext_ComponentType ConvertBack(VbComponentType componentType)
        {
			vbext_ComponentType msoComponentType;
			switch (componentType)
            {
				case VbComponentType.ActiveXDesigner:
					msoComponentType = vbext_ComponentType.vbext_ct_ActiveXDesigner;
                    break;
				case VbComponentType.ClassModule:
					msoComponentType = vbext_ComponentType.vbext_ct_ClassModule;
                    break;
				case VbComponentType.Document:
					msoComponentType = vbext_ComponentType.vbext_ct_Document;
                    break;
				case VbComponentType.MSForm:
					msoComponentType = vbext_ComponentType.vbext_ct_MSForm;
                    break;
				default:
					msoComponentType = vbext_ComponentType.vbext_ct_StdModule;
                    break;
            }
            return msoComponentType;
        }
    }
}
