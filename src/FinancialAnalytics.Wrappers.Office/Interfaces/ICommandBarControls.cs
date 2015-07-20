using System;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICommandBarControls : IEntitiesCollectionWrapper<ICommandBarControls, ICommandBarControl>
	{
		ICommandBarControl this[object index] { get; }

		ICommandBarControl Add(
			ControlType type,
			Object id,
			Object parameter,
			Object before,
			Object temporary);
	}
}