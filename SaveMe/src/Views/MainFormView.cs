using src.Factories;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace src
{
    /// <summary>
    ///     The main form of the application and app entry point.
    /// </summary>
    public partial class MainFormView : Form, IMainFormView
    {
        public event EventHandler AutoSaveFrequencySelected;
        public event EventHandler SaveSelectedButtonClicked;
        public event EventHandler SaveAllButtonClicked;


        public MainFormView()
        {
            InitializeComponent();
            InitializePresenter();
        }

        public void InitializePresenter()
        {
            PresenterFactory.CreateForView(this);
        }

        public ListBox ListOfOpenOfficeApplications
        {
            get { return listAllApplications; }
            set { listAllApplications = value; }
        }

        public ListBox DropDownAutoSaveFrequencies
        {
            get { return dropdownAutoSaveFrequencies; }
            set { dropdownAutoSaveFrequencies = value; }
        }

        public Button SaveSelectedButton
        {
            get { return btnSaveSelected; }
            set { btnSaveSelected = value; }
        }

        public Button SaveAllButton
        {
            get { return btnSaveAll; }
            set { btnSaveAll = value; }
        }
    }
}
