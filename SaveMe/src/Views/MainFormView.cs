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
    public partial class MainForm : Form, IMainFormView
    {
        public event EventHandler AutoSaveFrequencySelected;

        public MainForm()
        {
            InitializeComponent();
            InitializePresenter();
        }

        public void InitializePresenter()
        {
            PresenterFactory.CreateForView(this);
        }

        public CheckedListBox ListOfOpenOfficeApplications
        {
            get { return listAllApplications; }
            set { listAllApplications = value; }
        }

        public ListBox DropDownAutoSaveFrequency
        {
            get { return dropdownAutoSaveFrequencies; }
            set { dropdownAutoSaveFrequencies = value; }
        }
    }
}
