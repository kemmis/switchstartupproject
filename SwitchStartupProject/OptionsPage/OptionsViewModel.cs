using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;

namespace LucidConcepts.SwitchStartupProject.OptionsPage
{
    public class OptionsViewModel : INotifyPropertyChanged
    {
        private readonly OptionPage model;

        public OptionsViewModel(OptionPage model)
        {
            this.model = model;
        }

        #region Mode

        public IEnumerable<EMode> ModeValues
        {
            get { return Enum.GetValues(typeof (EMode)).OfType<EMode>(); }
        }

        public EMode SelectedMode
        {
            get { return model.Mode; }
            set
            {
                model.Mode = value;
                _RaisePropertyChanged("SelectedMode");
                _RaisePropertyChanged("MruVisibility");
            }
        }

        public Visibility MruVisibility
        {
            get { return model.Mode == EMode.MostRecentlyUsed ? Visibility.Visible : Visibility.Hidden; }
        }

        public int MruCount
        {
            get { return model.MostRecentlyUsedCount; }
            set
            {
                model.MostRecentlyUsedCount = value;
                _RaisePropertyChanged("MruCount");
            }
        }

        #endregion

        #region INotifyPropertyChanged

        protected void _RaisePropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged = (s, e) => { };

        #endregion
    }
}
