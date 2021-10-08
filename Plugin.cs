using PPL_Lib;
using OV_Overlay;
using LCI_Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;

///----------------------------------------------------------------------------
/// This example plugin demonstrates
///   1) Adding a menu item to the main O-Calc Pro menu
///   2) Enabling and disabling that menu item
///   3) Performing a task when the item is clicked
///----------------------------------------------------------------------------

namespace OCalcProPlugin
{
    public class Plugin : PPLPluginInterface
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        /// <summary>
        /// THis is the handle to the main O-Calc Pro component
        /// </summary>
        PPL_Lib.PPLMain cPPLMain = null;

        /// <summary>
        /// Declare the type of plugin as one of:
        ///         DOCKED_TAB
        ///         MENU_ITEM
        ///         BOTH_DOCKED_AND_MENU
        ///         CLEARANCE_SAG_PROVIDER
        /// </summary>
        public PLUGIN_TYPE Type
        {
            get
            {
                return PLUGIN_TYPE.BOTH_DOCKED_AND_MENU;
            }
        }

        /// <summary>
        /// Declare the name of the plugin usd for synthesizing the registry keys ect
        /// </summary>
        public String Name
        {
            get
            {
                return "Testing";
            }
        }

        /// <summary>
        /// Optionally declare a description string (defaults to the name);
        /// </summary>
        public String Description
        {
            get
            {
                return Name;
            }
        }

        /// <summary>
        /// Add a tabbed form to the tabbed window (if the plugin type is 
        /// PLUGIN_TYPE.DOCKED_TAB 
        /// or
        /// PLUGIN_TYPE.BOTH_DOCKED_AND_MENU
        /// </summary>
        /// <param name="pPPLMain"></param>
        public void AddForm(PPL_Lib.PPLMain pPPLMain)
        {
            cPPLMain = pPPLMain;
            cForm = new PluginForm(pPPLMain);
            Guid guid = new Guid(0x868e6fd5, 0x1d3, 0x4b44, 0xb0, 0x58, 0xfb, 0xbe, 0x98, 0x7a, 0xae, 0x94);
            cForm.cGuid = guid;
            cPPLMain.cDockedPanels.Add(cForm.cGuid.ToString(), cForm);
            foreach (Control ctrl in cPPLMain.Controls)
            {
                if (ctrl is WeifenLuo.WinFormsUI.Docking.DockPanel)
                {
                    cForm.Show(ctrl as WeifenLuo.WinFormsUI.Docking.DockPanel, WeifenLuo.WinFormsUI.Docking.DockState.Document);
                }
            }
            cForm.Show();
        }

        PluginForm cForm = null;

        private class SimpleSpan : INotifyPropertyChanged
        {
            private string _name;
            private string _type;
            private double _length;

            public event PropertyChangedEventHandler PropertyChanged;

            public string Name
            {
                get
                {
                    return _name;
                }
                set
                {
                    _name = value;
                    NotifyPropertyChanged();
                }
            }
            public string Type 
            {
                get
                {
                    return _type;
                }
                set
                {
                    _type = value;
                    NotifyPropertyChanged();
                }
            }
            public double Length 
            {
                get
                {
                    return Math.Round(_length / 12, 2);
                }
                set
                {
                    _length = value;
                    NotifyPropertyChanged();
                }
            }

            private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
            {
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
                }
            }
        }

        private class SimplePole : INotifyPropertyChanged
        {
            private string _poleNumber;
            private double _poleLength;
            private double _setDepth;
            private double _difference;
            private BindingList<SimpleSpan> _spans;
            private Color _mra;

            public event PropertyChangedEventHandler PropertyChanged;

            public string PoleNumber
            {
                get
                {
                    return _poleNumber;
                }
                set
                {
                    _poleNumber = value;
                    NotifyPropertyChanged();
                }
            }
            public double PoleLength
            {
                get
                {
                    return Math.Round(_poleLength / 12, 2);
                }
                set
                {
                    _poleLength = value;
                    NotifyPropertyChanged();
                }
            }
            public double SetDepth
            {
                get
                {
                    return Math.Round(_setDepth / 12, 2);
                }
                set
                {
                    _setDepth = value;
                    NotifyPropertyChanged();
                }
            }
            public double Difference
            {
                get
                {
                    return Math.Round(_difference, 2);
                }
                set
                {
                    _difference = value;
                    NotifyPropertyChanged();
                }
            }
            public BindingList<SimpleSpan> Spans
            {
                get
                {
                    return _spans;
                }
                set
                {
                    _spans = value;
                    NotifyPropertyChanged();
                }
            }
            public Color MRA
            {
                get
                {
                    return _mra;
                }
                set
                {
                    _mra = value;
                    NotifyPropertyChanged();
                }
            }
            private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
            {
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
                }
            }
        }

        class PluginForm : WeifenLuo.WinFormsUI.Docking.DockContent
        {
            private DataGridView pole_dataGridView = new DataGridView();
            private DataGridView spans_dataGridView = new DataGridView();

            private BindingSource pole_bindingSource = new BindingSource();
            private BindingSource spans_bindingSource = new BindingSource();

            public PluginForm(PPL_Lib.PPLMain pPPLMain)
            {
                AllocConsole();

                var cPPLMain = pPPLMain;

                this.Name = "pluginForm";
                this.Text = "MRA Colors";

                var myPole = new SimplePole();
                myPole.Spans = new BindingList<SimpleSpan>();

                // spans_dataGridView initial setup
                this.Controls.Add(spans_dataGridView);

                spans_dataGridView.Dock = DockStyle.Fill;
                //spans_dataGridView.Location = new Point(0, 200);
                //spans_dataGridView.Anchor = (AnchorStyles.Top & AnchorStyles.Bottom & AnchorStyles.Right);
                spans_dataGridView.AllowUserToAddRows = false;
                spans_dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
                spans_dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                spans_dataGridView.BorderStyle = BorderStyle.Fixed3D;
                spans_dataGridView.AllowUserToResizeColumns = false;
                spans_dataGridView.AllowUserToResizeRows = false;
                spans_dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

                spans_dataGridView.AutoGenerateColumns = true;

                // pole_dataGridView initial setup
                this.Controls.Add(pole_dataGridView);

                pole_dataGridView.Dock = DockStyle.Top;
                pole_dataGridView.AllowUserToAddRows = false;
                pole_dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                pole_dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                pole_dataGridView.BorderStyle = BorderStyle.Fixed3D;
                pole_dataGridView.Size = new System.Drawing.Size(0, 44);
                pole_dataGridView.AllowUserToResizeColumns = false;
                pole_dataGridView.AllowUserToResizeRows = false;
                pole_dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

                pole_dataGridView.AutoGenerateColumns = true;

                // pole_bindingSource do binding
                pole_bindingSource.DataSource = myPole;
                pole_dataGridView.DataSource = pole_bindingSource;

                // More pole_dataGridView customizing
                pole_dataGridView.RowHeadersVisible = false;
                pole_dataGridView.Columns["PoleNumber"].HeaderText = "Pole";
                pole_dataGridView.Columns["PoleLength"].HeaderText = "Pole Length ( ft. )";
                pole_dataGridView.Columns["SetDepth"].HeaderText = "Set Depth ( ft. )";
                pole_dataGridView.Columns["Difference"].HeaderText = "Difference ( ft. )";

                // spans_bindingSource do binding
                spans_bindingSource.DataSource = myPole.Spans;
                spans_dataGridView.DataSource = spans_bindingSource;

                // More spans_dataGridView customizing
                spans_dataGridView.RowHeadersVisible = false;
                spans_dataGridView.ColumnHeadersVisible = true;
                spans_dataGridView.Columns["Type"].Visible = false;
                spans_dataGridView.Columns["Length"].Visible = true;

                spans_dataGridView.Columns["Name"].HeaderText = "Spans";

                // Catch the pole_dataGridView.DataBindingComplete event
                pole_dataGridView.DataBindingComplete += (object sender, DataGridViewBindingCompleteEventArgs e) =>
                {
                    if (pole_dataGridView.Rows.Count > 0)
                    {
                        if (Convert.ToDouble(pole_dataGridView.Rows[0].Cells["Difference"].Value) < 30)
                        {
                            pole_dataGridView.Rows[0].Cells["Difference"].Style.BackColor = Color.PaleVioletRed;
                        }
                        else
                        {
                            pole_dataGridView.Rows[0].Cells["Difference"].Style.BackColor = Color.PaleGreen;
                        }

                        if (Convert.ToDouble(pole_dataGridView.Rows[0].Cells["SetDepth"].Value) < 0)
                        {
                            pole_dataGridView.Rows[0].Cells["SetDepth"].Style.BackColor = Color.PaleVioletRed;
                        }
                        else
                        {
                            pole_dataGridView.Rows[0].Cells["SetDepth"].Style.BackColor = Color.White;
                        }
                    }
                };

                // Catch the spans_dataGridView.RowsAdded event
                spans_dataGridView.RowsAdded += (object sender, DataGridViewRowsAddedEventArgs e) =>
                {
                    var spanLength = Convert.ToDouble(spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Value);

                    // Color the rows based on span lengths

                    if (spanLength < 300)
                    {
                        spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Style.BackColor = Color.PaleGreen;
                    }

                    if (spanLength >= 300 & spanLength < 350)
                    {
                        spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Style.BackColor = Color.PaleGoldenrod;
                    }

                    if (spanLength >= 350)
                    {
                        spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Style.BackColor = Color.PaleVioletRed;
                    }
                };

                cPPLMain.BeforeEvent += (object sender, PPL_Lib.PPLMain.EVENT_TYPE pEventType, ref bool pAbortEvent) =>
                {
                    //Console.WriteLine("BeforeEvent: " + pEventType.ToString());

                    if (pEventType == PPLMain.EVENT_TYPE.REBUILD_3D)
                    {
                        var pole = pPPLMain.GetPole();

                        if (pole != null)
                        {
                            myPole.Spans.Clear();

                            var elementList = pole.GetElementList();
                            var elementListLength = elementList.Count;

                            myPole.PoleNumber = pole.GetValueString("Pole Number");
                            myPole.PoleLength = pole.GetValueDouble("LengthInInches");
                            myPole.SetDepth = pole.GetValueDouble("BuryDepthInInches");

                            myPole.Difference = myPole.PoleLength - myPole.SetDepth;

                            foreach (var item in pole.GetElementList())
                            {
                                var itemType = item.GetSchemaKey();

                                if (itemType == "Span")
                                {
                                    var span = new SimpleSpan();
                                    span.Name = item.ToString();
                                    span.Type = item.GetValueString("SpanType");
                                    span.Length = item.GetValueDouble("SpanDistanceInInches");

                                    myPole.Spans.Add(span);
                                }
                            }
                        }
                    };

                    if (pEventType == PPLMain.EVENT_TYPE.REBUILD_ALL_DISPLAYS)
                    {
                        //cPPLMain.c3DView.NeutralCameraPosition();
                        //cPPLMain.c3DView.pole3Dview1.cCameraManager.UpdateCamera();

                        //var pole = cPPLMain.GetPole();

                        //var twod = cPPLMain.c2DView;
                        //var threed = cPPLMain.c3DView;
                        //var ALDE = cPPLMain.cActiveLineDesignElement;
                        //var inv = cPPLMain.cInventory;
                    }

                    if (pEventType == PPLMain.EVENT_TYPE.LD_SELECT_POLE)
                    {
                        // Triggers on pole select in 3d view only.

                        //cPPLMain.c3DView.NeutralCameraPosition();
                        //cPPLMain.c3DView.ResetCameraPosition();

                        //Console.WriteLine("LD_SELECTPOLE Is this enabled? Guess not.");
                    }

                    if (pEventType == PPLMain.EVENT_TYPE.SUBSTITUTION)
                    {
                        // TODO
                    }
                };
            }

            public Guid cGuid;
            protected override string GetPersistString()
            {
                return cGuid.ToString();
            }
        }

        /// <summary>
        /// Perform clearance analysis if type is PLUGIN_TYPE.CLEARANCE_SAG_PROVIDER
        /// </summary>
        /// <param name="pMain"></param>
        /// <returns></returns>
        public PPLClearance.ClearanceSagProvider GetClearanceSagProvider(PPL_Lib.PPLMain pMain)
        {
            System.Diagnostics.Debug.Assert(Type == PLUGIN_TYPE.CLEARANCE_SAG_PROVIDER, Name + " is not a clearance provider plugin.");
            return null;
        }

        //the toolstrip items we will add to O-Calc Pro
        ToolStripMenuItem quickString_MenuItemButton = null;
        ToolStripMenuItem checkAll_MenuItemButton = null;
        ToolStripMenuItem fixedWind_MenuItemButton = null;
        ToolStripMenuItem setElevation_MenuItemButton = null;
        ToolStripMenuItem testing_MenuItemButton = null;

        public void AddToMenu(PPL_Lib.PPLMain pPPLMain, System.Windows.Forms.ToolStrip pToolStrip)
        {
            //save the reference to the O-Calc Pro main
            cPPLMain = pPPLMain;

            //create the toolstrip buttons
            quickString_MenuItemButton = new ToolStripMenuItem("Quick String");
            quickString_MenuItemButton.AutoToolTip = true;
            quickString_MenuItemButton.ToolTipText = Description;
            quickString_MenuItemButton.Enabled = true;

            checkAll_MenuItemButton = new ToolStripMenuItem("Check All");
            checkAll_MenuItemButton.AutoToolTip = true;
            checkAll_MenuItemButton.ToolTipText = Description;
            checkAll_MenuItemButton.Enabled = true;

            fixedWind_MenuItemButton = new ToolStripMenuItem("Calculate Checked Fixed Wind");
            fixedWind_MenuItemButton.AutoToolTip = true;
            fixedWind_MenuItemButton.ToolTipText = Description;
            fixedWind_MenuItemButton.Enabled = true;

            setElevation_MenuItemButton = new ToolStripMenuItem("Set Elevation");
            setElevation_MenuItemButton.AutoToolTip = true;
            setElevation_MenuItemButton.ToolTipText = Description;
            setElevation_MenuItemButton.Enabled = true;

            testing_MenuItemButton = new ToolStripMenuItem("Testing");
            testing_MenuItemButton.AutoToolTip = true;
            testing_MenuItemButton.ToolTipText = Description;
            testing_MenuItemButton.Enabled = false;


            //find the dropdown menu we want to add the toolstrip buttons to
            int itemindex = 0;
            System.Diagnostics.Debug.Assert(pToolStrip.Items[itemindex] is ToolStripDropDownButton);
            if (pToolStrip.Items[itemindex] is ToolStripDropDownButton)
            {
                ToolStripDropDownButton tsb = pToolStrip.Items[itemindex] as ToolStripDropDownButton;
                System.Diagnostics.Debug.Assert(tsb.Text == "&File");
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, quickString_MenuItemButton);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, checkAll_MenuItemButton);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, fixedWind_MenuItemButton);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, setElevation_MenuItemButton);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, new ToolStripSeparator());
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, testing_MenuItemButton);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, new ToolStripSeparator());
            }

            //add and event handler to detect the toolstrip button bein clicked by the user
            quickString_MenuItemButton.Click += quickString_Click;
            checkAll_MenuItemButton.Click += checkAll_Click;
            fixedWind_MenuItemButton.Click += fixedWind_Click;
            setElevation_MenuItemButton.Click += setElevation_Click;
            testing_MenuItemButton.Click += testing_Click;
        }

        void quickString_Click(object sender, EventArgs e)
        {
            try
            {
                var catalogs = cPPLMain.cCatalogManager.cCatalogs;

                foreach (var catalog in catalogs)
                {
                    if (catalog.cName == "User")
                    {
                        var selected = catalog.cCatalog.GetSelectedElement();

                        if (selected != null)
                        {
                            cPPLMain.cLineDesignInteraction.ApplyStringingAssembly(selected, true, new StringBuilder("Error"));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        PPL_LineDesign.PolesListView cPolesList = null;
        void checkAll_Click(object sender, EventArgs e)
        {
            foreach (Control control in cPPLMain.cLineDesignInteraction.GetLD_MainForm().Controls)
            {
                if (control.GetType().ToString().Contains("SplitContainer"))
                {
                    foreach (Control splitterPanel in control.Controls)
                    {
                        foreach (Control sub2Control in splitterPanel.Controls)
                        {
                            foreach (Control item in sub2Control.Controls)
                            {
                                foreach (Control item2 in item.Controls)
                                {
                                    foreach (Control item3 in item2.Controls)
                                    {
                                        if (item3.Name == "cPolesList")
                                        {
                                            cPolesList = (PPL_LineDesign.PolesListView)item3;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            foreach (ListViewItem item in cPolesList.Items)
            {
                if (item.Checked == false)
                {
                    item.Checked = true;
                }
            }
        }

        private MenuStrip cLineDesignMenuStrip = null;
        private ToolStripMenuItem cfixedWind = null;

        void fixedWind_Click(object sender, EventArgs e)
        {
            foreach (Control control in cPPLMain.cLineDesignInteraction.GetLD_MainForm().Controls)
            {
                if (control.GetType().ToString().Contains("SplitContainer"))
                {
                    foreach (Control splitterPanel in control.Controls)
                    {
                        foreach (Control sub2Control in splitterPanel.Controls)
                        {
                            foreach (Control item in sub2Control.Controls)
                            {
                                if (item.Name == "lineDesignMenuStrip")
                                {
                                    cLineDesignMenuStrip = (MenuStrip)item;
                                }
                            }
                        }
                    }
                }
            }


            foreach (var item in cLineDesignMenuStrip.Items)
            {
                if (item.GetType().ToString() == "PPL_Lib.PPL_ToolStripMenuItem" && item.ToString() == "&Calculate")
                {
                    var newItem = (ToolStripMenuItem)item;

                    foreach (var item2 in newItem.DropDownItems)
                    {
                        if (item2.ToString() == "Line Analysis")
                        {
                            var newItem2 = (ToolStripMenuItem)item2;

                            foreach (var item3 in newItem2.DropDownItems)
                            {
                                if (item3.ToString() == "&Checked Poles Only")
                                {
                                    var newitem3 = (ToolStripMenuItem)item3;

                                    foreach (var item4 in newitem3.DropDownItems)
                                    {
                                        if (item4.ToString() == "&Fixed Wind")
                                        {
                                            cfixedWind = (ToolStripMenuItem)item4;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            cfixedWind.PerformClick();
        }

        private ToolStripMenuItem cSingleSetElevation = null;
        private ToolStripMenuItem cCheckedSetElevation = null;

        void setElevation_Click(object sender, EventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.Text == "OV Overlay")
                {
                    foreach (var item in form.Controls)
                    {
                        var newItem = (OV_Overlay.OverlayControl)item;

                        foreach (var item2 in newItem.Controls)
                        {
                            if (item2.GetType().ToString().Contains("MenuStrip")) {
                                var newItem2 = (MenuStrip)item2;

                                foreach (var item3 in newItem2.Items)
                                {
                                    if (item3.GetType().ToString().Contains("ToolStripMenuItem"))
                                    {
                                        var newItem3 = (ToolStripMenuItem)item3;

                                        foreach (var item4 in newItem3.DropDownItems)
                                        {
                                            Console.WriteLine($"{item4} {item4.GetType()}");

                                            if (item4.ToString() == "Set Pole &Elevation")
                                            {
                                                var newItem4 = (ToolStripMenuItem)item4;

                                                foreach (var item5 in newItem4.DropDownItems)
                                                {
                                                    Console.WriteLine($"{item5} {item5.GetType()}");

                                                    if (item5.ToString() == "Current &Pole")
                                                    {
                                                        cSingleSetElevation = (ToolStripMenuItem)item5;
                                                    }
                                                }
                                            }
                                        }
                                    }
                        
                                }
                            }

                        }
                    }
                }
            }

            cSingleSetElevation.PerformClick();
        }

        void testing_Click(object sender, EventArgs e)
        {
            PPL_ADGV.DataGridViewWithFilteringL<LCI_Lib.MakeReadyAssessment.MakeReadyAssessmentRecord> mra_adgv = null;

            foreach (Form form in Application.OpenForms)
            {
                if (form.Text == "MRA")
                {
                    foreach (Control control in form.Controls)
                    {

                        Console.WriteLine($"Control Count: {control.Controls.Count}\tType: {control.GetType()}\tName: {control.Name}\tText: {control.Text}");

                        if (control.GetType().ToString().Contains("DataGridViewWithFilteringL"))
                        {
                            Console.WriteLine("Casting the control to a custom datagridview");

                            try
                            {
                                mra_adgv = (PPL_ADGV.DataGridViewWithFilteringL<LCI_Lib.MakeReadyAssessment.MakeReadyAssessmentRecord>)control;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                            if (mra_adgv != null)
                            {
                                for (int i = 0; i <= mra_adgv.Rows.Count; i++)
                                {
                                    Console.WriteLine(mra_adgv.Rows[i].Cells["PLDBID"].Value.ToString());
                                    Console.WriteLine(cPPLMain.GetPole().GetValueString("Pole Number"));

                                    if (mra_adgv.Rows[i].Cells["PLDBID"].Value.ToString() == cPPLMain.GetPole().GetValueString("Pole Number"))
                                    {
                                        Console.WriteLine("PLDBID & POle Number match");

                                        //var color = mra_adgv.Rows[i].Cells["MRA"].Value.ToString();

                                        //if (color == "Unset")
                                        //{
                                        //    myPole.MRA = Color.Gray;
                                        //}

                                        //if (color == "Green")
                                        //{
                                        //    myPole.MRA = Color.PaleGreen;
                                        //}

                                        //if (color == "Yellow")
                                        //{
                                        //    myPole.MRA = Color.PaleGoldenrod;
                                        //}

                                        //if (color == "Red")
                                        //{
                                        //    myPole.MRA = Color.PaleVioletRed;
                                        //}
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
