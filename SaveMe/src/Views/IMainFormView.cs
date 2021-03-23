using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace src
{
    public interface IMainFormView
    {
        ListBox DropDownAutoSaveFrequencies { get; set; }
        ListBox ListOfOpenOfficeApplications { get; set; }
        Button SaveSelectedButton { get; set; }
        Button SaveAllButton { get; set; }

        event EventHandler AutoSaveFrequencySelected;

        event EventHandler Load;

        event EventHandler SaveSelectedButtonClicked;
        event EventHandler SaveAllButtonClicked;

        void InitializePresenter();

        object Invoke(Delegate method);
    }
}