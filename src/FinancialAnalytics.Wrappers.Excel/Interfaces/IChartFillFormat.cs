using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IChartFillFormat : IEntityWrapper<IChartFillFormat>
    {
        bool Visible { get; set; }

        FillType Type { get; }

        IChartColorFormat ForeColor { get; }

        IChartColorFormat BackColor { get; }

        int GradientVariant { get; }

        float GradientDegree { get; }

        PatternType Pattern { get; }


        GradientColorType GradientColorType { get; }

        GradientStyle GradientStyle { get; }

        PresetGradientType PresetGradientType { get; }

        PresetTexture PresetTexture { get; }

        void Solid();

        void OneColorGradient(GradientStyle style, int variant, float degree);

        void TwoColorGradient(GradientStyle style, int variant, float degree);

        void PresetGradient(GradientStyle style, int variant, PresetGradientType presetGradientType);

        void PresetTextured(PresetTexture presetTexture);

        void Patterned(PatternType pattern);
    }
}
