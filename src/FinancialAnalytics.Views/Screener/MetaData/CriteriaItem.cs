using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using DryTools;

namespace FinancialAnalytics.Views.Screener.MetaData
{
    internal class CriteriaItem : INotifyPropertyChanged, ICloneable
    {
        public CriteriaItem(Type dataType)
        {
            DataType = dataType;
        }

        public bool SupportAutocompletion { get; set; }

        public string DisplayLabel { get; set; }

        public string DisplayToolTip { get; set; }

        public string TargetBaseset { get; set; }

        public string TargetPrvtComp { get; set; }

        public Type DataType { get; private set; }

        public CriteriaDataContext DataContext { get; set; }

        public string Value
        {
            get { return _value; }
            set
            {
                if (_value != value)
                {
                    _value = value;
                    NotifyPropertyChanged(() => Value);
                }
            }
        }

        public string Value2
        {
            get { return _value2; }
            set
            {
                if (_value2 != value)
                {
                    _value2 = value;
                    NotifyPropertyChanged(() => Value2);
                }
            }
        }

        public Visibility ItemVisibility
        {
            get { return _itemVisibility; }
            set
            {
                if (_itemVisibility != value)
                {
                    _itemVisibility = value;
                    NotifyPropertyChanged(() => ItemVisibility);
                }
            }
        }

        public override string ToString()
        {
            return DisplayLabel ?? base.ToString();
        }

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        protected void NotifyPropertyChanged<TProperty>(Func<TProperty> getPropertyExpression)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(ExtractName.From(getPropertyExpression)));
        }

        #endregion

        #region IClonable

        public CriteriaItem Clone()
        {
            return (CriteriaItem)MemberwiseClone();
        }

        object ICloneable.Clone()
        {
            return Clone();
        }

        #endregion

        #region Implementation

        private string _value;
        private string _value2;
        private Visibility _itemVisibility;

        #endregion
    }
}
