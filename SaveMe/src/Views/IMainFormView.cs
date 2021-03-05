using System;
using System.Windows.Forms;

namespace src
{
    public interface IMainFormView
    {
        ListBox DropDownAutoSaveFrequency { get; set; }
        CheckedListBox ListOfOpenOfficeApplications { get; set; }

        event EventHandler AutoSaveFrequencySelected;

        void InitializePresenter();
    }
}