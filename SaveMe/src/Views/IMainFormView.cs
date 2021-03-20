using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace src
{
    public interface IMainFormView
    {
        ListBox DropDownAutoSaveFrequency { get; set; }
        CheckedListBox ListOfOpenOfficeApplications { get; set; }
        //BindingList<string> OpenOfficeAppNames { get; set; }

        event EventHandler AutoSaveFrequencySelected;

        event EventHandler Load;

        void InitializePresenter();

        object Invoke(Delegate method);
    }
}