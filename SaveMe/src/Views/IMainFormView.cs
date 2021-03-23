using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace src
{
    public interface IMainFormView
    {
        ListBox DropDownAutoSaveFrequency { get; set; }
        ListBox ListOfOpenOfficeApplications { get; set; }

        event EventHandler AutoSaveFrequencySelected;

        event EventHandler Load;

        event EventHandler SaveSelectedButtonClicked;
        event EventHandler SaveAllButtonClicked;

        void InitializePresenter();

        object Invoke(Delegate method);
    }
}