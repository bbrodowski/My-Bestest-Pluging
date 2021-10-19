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
using System.Windows.Forms.Integration;
using System.Windows.Controls;
using ListViewItem = System.Windows.Forms.ListViewItem;
using Control = System.Windows.Forms.Control;
using System.Windows.Input;
using MouseEventArgs = System.Windows.Forms.MouseEventArgs;
using TabControl = System.Windows.Forms.TabControl;

///----------------------------------------------------------------------------
/// This example plugin demonstrates
///   1) Adding a menu item to the main O-Calc Pro menu
///   2) Enabling and disabling that menu item
///   3) Performing a task when the item is clicked
///----------------------------------------------------------------------------

namespace OCalcProPlugin
{
    public class SimpleSpan : INotifyPropertyChanged
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
    public class SimplePole : INotifyPropertyChanged
    {
        private string _poleNumber;
        private double _poleLength;
        private double _setDepth;
        private double _difference;
        private string _suggestedMRA;
        private BindingList<SimpleSpan> _spans;

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
        public string SuggestedMRA
        {
            get
            {
                return _suggestedMRA;
            }
            set
            {
                _suggestedMRA = value;
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
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
    public static class SharedReferences
    {
        public static PPLMain cPPLMain { get; set; }
        public static int cLineDesign { get; set; }
        public static OV_Overlay.OverlayControl cOVOverlay { get; set; }
        public static LCI_Lib.MakeReadyAssessment.MakeReadyAssesmentForm cMRA { get; set; }
        public static PPL_ADGV.DataGridViewWithFilteringL<LCI_Lib.MakeReadyAssessment.MakeReadyAssessmentRecord> cADGV { get; set; }
        public static SimplePole cMyPole { get; set; }
        public static PPLCatalogManager.CatalogDefinition cUserCatalog { get; set; }
        public static PPLCatalogManager.CatalogDefinition cMasterCatalog { get; set; }
        public static PPL_LineDesign.PolesListView cPolesList { get; set; }
        public static ToolStripMenuItem cfixedWind { get; set; }
        public static ToolStripMenuItem cSortCheckHelper { get; set; }
        public static ToolStripMenuItem cSingleSetElevation { get; set; }
        public static ToolStripMenuItem cCheckedSetElevation { get; set; }
        public static ToolStripMenuItem cAllSetElevation { get; set; }

        public static void GetReferences()
        {
            cUserCatalog = cPPLMain.cCatalogManager.cCatalogs[0];
            cMasterCatalog = cPPLMain.cCatalogManager.cCatalogs[1];

            foreach (Form form in Application.OpenForms)
            {
                if (form.Text == "LCI")
                {
                    
                }

                // Charter RDOF LCI only
                if (form.Text == "MRA")
                {
                    cMRA = (LCI_Lib.MakeReadyAssessment.MakeReadyAssesmentForm)form;

                    cADGV = (PPL_ADGV.DataGridViewWithFilteringL<LCI_Lib.MakeReadyAssessment.MakeReadyAssessmentRecord>)form.Controls[1];
                }

                // OV Overlay Form
                if (form.Text.Contains("OV Overlay"))
                {
                    cOVOverlay = (OV_Overlay.OverlayControl)form.Controls[0];

                    cPolesList = cOVOverlay.cLD_MainForm.cPolesList;

                    var lineDesignMenuStrip = cOVOverlay.cLD_MainForm.lineDesignMenuStrip;

                    var calculateToolStipMenuItem = (ToolStripMenuItem)lineDesignMenuStrip.Items[10];

                    var lineAnalysisToolStipMenuItem = (ToolStripMenuItem)calculateToolStipMenuItem.DropDownItems[2];

                    var checkPolesOnlyToolStipMenuItem = (ToolStripMenuItem)lineAnalysisToolStipMenuItem.DropDownItems[1];

                    cfixedWind = (ToolStripMenuItem)checkPolesOnlyToolStipMenuItem.DropDownItems[0];

                    var polesMenustrip = cOVOverlay.cLD_MainForm.polesMenuStrip;

                    var viewToolStripMenuItem = (ToolStripMenuItem)polesMenustrip.Items[2];

                    cSortCheckHelper = (ToolStripMenuItem)viewToolStripMenuItem.DropDownItems[6];

                    foreach (var item2 in cOVOverlay.Controls)
                    {
                        //Console.WriteLine(item2 + ", " + item2.GetType());

                        if (item2.GetType() == typeof(MenuStrip))
                        {
                            var ovoverlayMenuStrip = (MenuStrip)item2;

                            var toolsToolStripMenuItem = (ToolStripMenuItem)ovoverlayMenuStrip.Items[2];

                            var setPoleElevationToolStripMenuItem = (ToolStripMenuItem)toolsToolStripMenuItem.DropDownItems[1];

                            cSingleSetElevation = (ToolStripMenuItem)setPoleElevationToolStripMenuItem.DropDownItems[0];

                            cCheckedSetElevation = (ToolStripMenuItem)setPoleElevationToolStripMenuItem.DropDownItems[1];

                            cAllSetElevation = (ToolStripMenuItem)setPoleElevationToolStripMenuItem.DropDownItems[2];
                        }
                    }
                }
            }
        }
    }
    public class Plugin : PPLPluginInterface
    {
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

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        /// <summary>
        /// Add a tabbed form to the tabbed window (if the plugin type is 
        /// PLUGIN_TYPE.DOCKED_TAB 
        /// or
        /// PLUGIN_TYPE.BOTH_DOCKED_AND_MENU
        /// </summary>
        /// <param name="pPPLMain"></param>
        public void AddForm(PPL_Lib.PPLMain pPPLMain)
        {
            //AllocConsole();

            SharedReferences.cPPLMain = pPPLMain;
            SharedReferences.GetReferences();

            cForm = new PluginForm();
            Guid guid = new Guid(0x9eb1dc5c, 0xe2ca, 0x4a60, 0x88, 0x7e, 0xad, 0xb9, 0xa8, 0x77, 0x36, 0x2e);
            cForm.cGuid = guid;
            SharedReferences.cPPLMain.cDockedPanels.Add(cForm.cGuid.ToString(), cForm);
            foreach (Control ctrl in SharedReferences.cPPLMain.Controls)
            {
                if (ctrl is WeifenLuo.WinFormsUI.Docking.DockPanel)
                {
                    cForm.Show(ctrl as WeifenLuo.WinFormsUI.Docking.DockPanel, WeifenLuo.WinFormsUI.Docking.DockState.Document);
                }
            }
            cForm.Show();
        }

        PluginForm cForm = null;

        class PluginForm : WeifenLuo.WinFormsUI.Docking.DockContent
        {
            private DataGridView pole_dataGridView = new DataGridView();
            private DataGridView spans_dataGridView = new DataGridView();

            private BindingSource pole_bindingSource = new BindingSource();
            private BindingSource spans_bindingSource = new BindingSource();

            private int gSuggestedMRA;

            public PluginForm()
            {
                this.Name = "pluginForm";
                this.Text = "MRA Colors";

                SharedReferences.cMyPole = new SimplePole();
                SharedReferences.cMyPole.Spans = new BindingList<SimpleSpan>();

                // spans_dataGridView initial setup
                this.Controls.Add(spans_dataGridView);

                spans_dataGridView.Dock = DockStyle.Fill;
                //spans_dataGridView.Location = new Point(0, 200);
                //spans_dataGridView.Anchor = (AnchorStyles.Top & AnchorStyles.Bottom & AnchorStyles.Right);
                spans_dataGridView.AllowUserToAddRows = false;
                spans_dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
                spans_dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                spans_dataGridView.BorderStyle = BorderStyle.Fixed3D;
                spans_dataGridView.AllowUserToResizeColumns = true;
                spans_dataGridView.AllowUserToResizeRows = true;
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
                pole_dataGridView.AllowUserToResizeColumns = true;
                pole_dataGridView.AllowUserToResizeRows = true;
                pole_dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

                pole_dataGridView.AutoGenerateColumns = true;

                // pole_bindingSource do binding
                pole_bindingSource.DataSource = SharedReferences.cMyPole;
                pole_dataGridView.DataSource = pole_bindingSource;

                // More pole_dataGridView customizing
                pole_dataGridView.RowHeadersVisible = false;
                pole_dataGridView.Columns["PoleNumber"].HeaderText = "Pole";
                pole_dataGridView.Columns["PoleLength"].HeaderText = "Pole Length ( ft. )";
                pole_dataGridView.Columns["SetDepth"].HeaderText = "Set Depth ( ft. )";
                pole_dataGridView.Columns["Difference"].HeaderText = "Difference ( ft. )";

                // spans_bindingSource do binding
                spans_bindingSource.DataSource = SharedReferences.cMyPole.Spans;
                spans_dataGridView.DataSource = spans_bindingSource;

                // More spans_dataGridView customizing
                spans_dataGridView.RowHeadersVisible = false;
                spans_dataGridView.ColumnHeadersVisible = true;
                spans_dataGridView.Columns["Type"].Visible = false;
                spans_dataGridView.Columns["Length"].Visible = true;

                spans_dataGridView.Columns["Name"].HeaderText = "Spans";

                var red = Color.PaleVioletRed;
                var yellow = Color.PaleGoldenrod;
                var green = Color.PaleGreen;

                // Need to add an event on click of setdepth to substitute pole;

                // Catch the pole_dataGridView.DataBindingComplete event
                pole_dataGridView.DataBindingComplete += (object sender, DataGridViewBindingCompleteEventArgs e) =>
                {
                    if (pole_dataGridView.Rows.Count > 0)
                    {
                        if (Convert.ToDouble(pole_dataGridView.Rows[0].Cells["Difference"].Value) < 30)
                        {
                            pole_dataGridView.Rows[0].Cells["Difference"].Style.BackColor = red;
                        }
                        else
                        {
                            pole_dataGridView.Rows[0].Cells["Difference"].Style.BackColor = green;
                        }

                        if (Convert.ToDouble(pole_dataGridView.Rows[0].Cells["SetDepth"].Value) < 0)
                        {
                            pole_dataGridView.Rows[0].Cells["SetDepth"].Style.BackColor = red;
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
                        spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Style.BackColor = green;
                    }

                    if (spanLength >= 300 & spanLength < 350)
                    {
                        spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Style.BackColor = yellow;
                    }

                    if (spanLength >= 350)
                    {
                        spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Style.BackColor = red;
                    }
                };

                SharedReferences.cPPLMain.BeforeEvent += (object sender, PPL_Lib.PPLMain.EVENT_TYPE pEventType, ref bool pAbortEvent) =>
                {
                    //Console.WriteLine("BeforeEvent: " + pEventType.ToString());

                    if (pEventType == PPLMain.EVENT_TYPE.REBUILD_ALL_DISPLAYS)
                    {
                            SharedReferences.cPPLMain.c3DView.pole3Dview1.SignalCameraChanged();
                    }

                    if (pEventType == PPLMain.EVENT_TYPE.REBUILD_3D)
                    {
                        var cPole = SharedReferences.cPPLMain.GetPole();

                        if (cPole != null)
                        {
                            SharedReferences.cMyPole.Spans.Clear();

                            var elementList = cPole.GetElementList();
                            var elementListLength = elementList.Count;

                            SharedReferences.cMyPole.PoleNumber = cPole.GetValueString("Pole Number");
                            SharedReferences.cMyPole.PoleLength = cPole.GetValueDouble("LengthInInches");
                            SharedReferences.cMyPole.SetDepth = cPole.GetValueDouble("BuryDepthInInches");

                            SharedReferences.cMyPole.Difference = SharedReferences.cMyPole.PoleLength - SharedReferences.cMyPole.SetDepth;

                            foreach (var item in cPole.GetElementList())
                            {
                                var itemType = item.GetSchemaKey();

                                if (itemType == "Span")
                                {
                                    var span = new SimpleSpan();
                                    span.Name = item.ToString();
                                    span.Type = item.GetValueString("SpanType");
                                    span.Length = item.GetValueDouble("SpanDistanceInInches");

                                    SharedReferences.cMyPole.Spans.Add(span);
                                }
                            }

                            gSuggestedMRA = 0;

                            foreach (SimpleSpan span in SharedReferences.cMyPole.Spans)
                            {
                                if (span.Length < 300)
                                {
                                    var temp = 0;

                                    if (gSuggestedMRA > temp)
                                    {
                                        // Do Nothing
                                    } else
                                    {
                                        gSuggestedMRA = 0;
                                    }
                                }

                                if (span.Length >= 300 & span.Length < 350)
                                {
                                    var temp = 1;

                                    if (gSuggestedMRA > temp)
                                    {
                                        // Do Nothing
                                    }
                                    else
                                    {
                                        gSuggestedMRA = 1;
                                    }
                                }

                                if (span.Length >= 350)
                                {
                                    var temp = 2;

                                    if (gSuggestedMRA > temp)
                                    {
                                        // Do Nothing
                                    }
                                    else
                                    {
                                        gSuggestedMRA = 2;
                                    }
                                }
                            }

                            if (SharedReferences.cMyPole.Difference < 30)
                            {
                                gSuggestedMRA = 2;
                            }

                            if (gSuggestedMRA == 0)
                            {
                                SharedReferences.cMyPole.SuggestedMRA = "Green";
                                pole_dataGridView.Rows[0].Cells["SuggestedMRA"].Style.BackColor = green;
                            }

                            if (gSuggestedMRA == 1)
                            {
                                SharedReferences.cMyPole.SuggestedMRA = "Yellow";
                                pole_dataGridView.Rows[0].Cells["SuggestedMRA"].Style.BackColor = yellow;
                            }

                            if (gSuggestedMRA == 2)
                            {
                                SharedReferences.cMyPole.SuggestedMRA = "Red";
                                pole_dataGridView.Rows[0].Cells["SuggestedMRA"].Style.BackColor = red;
                            }

                        }
                    };
                };
            }

            public Guid cGuid;
            protected override string GetPersistString()
            {
                return cGuid.ToString();
            }
        }

        public PPLClearance.ClearanceSagProvider GetClearanceSagProvider(PPL_Lib.PPLMain pMain)
        {
            System.Diagnostics.Debug.Assert(Type == PLUGIN_TYPE.CLEARANCE_SAG_PROVIDER, Name + " is not a clearance provider plugin.");
            return null;
        }

        ToolStripMenuItem CreateMenuItemButton(string menuItemText, EventHandler handler)
        {
            var temp = new ToolStripMenuItem();

            temp = new ToolStripMenuItem(menuItemText);
            temp.AutoToolTip = true;
            temp.ToolTipText = Description;
            temp.Enabled = true;

            temp.Click += handler;

            return temp;
        }

        public void AddToMenu(PPL_Lib.PPLMain pPPLMain, System.Windows.Forms.ToolStrip pToolStrip)
        {
            var lineDesign_DropDown = new ToolStripMenuItem("Line Design");
            var mra_DropDown = new ToolStripMenuItem("MRA");
            var ovoverlay_DropDown = new ToolStripMenuItem("OV Overlay");
            var measurement_Dropdown = new ToolStripMenuItem("LCI Measurement Tool");

            var quickString_MenuItemButton = CreateMenuItemButton("Quick String", quickString_Click);

            var sortCheckHelper_MenuItemButton = CreateMenuItemButton("View, Sort / Check Helper", sortCheckHelper_Click);

            var currentSetElevation_MenuItemButton = CreateMenuItemButton("Set Pole Elevation, Current", currentSetElevation_Click);
            var checkedSetElevation_MenuItemButton = CreateMenuItemButton("Set Pole Elevation, Checked", checkedSetElevation_Click);
            var allSetElevation_MenuItemButton = CreateMenuItemButton("Set Pole Elevation, All", allSetElevation_Click);

            var combinedCheckAllFixedWind_MenuItemButton = CreateMenuItemButton("Combined Check All / Fixed Wind", combinedCheckAllFixedWind_Click);

            var setMRAUnset_MenuItemButton = CreateMenuItemButton("Set MRA to Unset", SetMRACell);
            var setMRAGreen_MenuItemButton = CreateMenuItemButton("Set MRA to Green", SetMRACell);
            var setMRAYellow_MenuItemButton = CreateMenuItemButton("Set MRA to Yellow", SetMRACell);
            var setMRARed_MenuItemButton = CreateMenuItemButton("Set MRA to Red", SetMRACell);
            var checkBackbone_MenuItemButton = CreateMenuItemButton("Check Backbone", SetMRACheckableCell);
            var checkDoubleWood_MenuItemButton = CreateMenuItemButton("Check Double Wood", SetMRACheckableCell);
            var checkSpanGuy_MenuItemButton = CreateMenuItemButton("Check Span Guy", SetMRACheckableCell);
            var check3rdParty_MenuItemButton = CreateMenuItemButton("Check 3rd Party", SetMRACheckableCell);
            var setMRASuggested_MenuItemButton = CreateMenuItemButton("Set MRA to Suggested", SetMRACell);

            var firstPoint_MenuItemButton = CreateMenuItemButton("First Point", firstPoint_Click);
            var secondPoint_MenuItemButton = CreateMenuItemButton("Second Point", secondPoint_Click);
            var reset_MenuItemButton = CreateMenuItemButton("Reset", secondPoint_Click);

            lineDesign_DropDown.DropDownItems.AddRange(new ToolStripItem[]
            {
                sortCheckHelper_MenuItemButton
            });

            mra_DropDown.DropDownItems.AddRange(new ToolStripItem[]
            {
                setMRAUnset_MenuItemButton,
                setMRAGreen_MenuItemButton,
                setMRAYellow_MenuItemButton,
                setMRARed_MenuItemButton,
                checkBackbone_MenuItemButton,
                checkDoubleWood_MenuItemButton,
                checkSpanGuy_MenuItemButton,
                check3rdParty_MenuItemButton,
                setMRASuggested_MenuItemButton
            });

            ovoverlay_DropDown.DropDownItems.AddRange(new ToolStripItem[]
            {
                currentSetElevation_MenuItemButton,
                checkedSetElevation_MenuItemButton,
                allSetElevation_MenuItemButton
            });

            measurement_Dropdown.DropDownItems.AddRange(new ToolStripItem[]
            {
                firstPoint_MenuItemButton,
                secondPoint_MenuItemButton,
                reset_MenuItemButton
            });;

            //find the dropdown menu we want to add the toolstrip buttons to
            int itemindex = 0;
            System.Diagnostics.Debug.Assert(pToolStrip.Items[itemindex] is ToolStripDropDownButton);
            if (pToolStrip.Items[itemindex] is ToolStripDropDownButton)
            {
                ToolStripDropDownButton tsb = pToolStrip.Items[itemindex] as ToolStripDropDownButton;
                System.Diagnostics.Debug.Assert(tsb.Text == "&File");

                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, lineDesign_DropDown);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, mra_DropDown);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, ovoverlay_DropDown);
                //tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, measurement_Dropdown);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, new ToolStripSeparator());
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, quickString_MenuItemButton);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, combinedCheckAllFixedWind_MenuItemButton);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, new ToolStripSeparator());
            }
        }

        private void CheckAll()
        {
            foreach (ListViewItem item in SharedReferences.cPolesList.Items)
            {
                if (item.Checked == false)
                {
                    item.Checked = true;
                }
            }
        }

        private bool CheckReference(ToolStripMenuItem item)
        {
            var temp = false;

            if (item != null)
                temp = true;
            else
            {
                SharedReferences.GetReferences();
                temp = true;
            }

            return temp;
        }

        private void quickString_Click(object sender, EventArgs e)
        {
            var selected = SharedReferences.cUserCatalog.cCatalog.GetSelectedElement();

            if (selected != null)
                SharedReferences.cPPLMain.cLineDesignInteraction.ApplyStringingAssembly(selected, true, new StringBuilder("Error"));
        }

        private void sortCheckHelper_Click(object sender, EventArgs e)
        {
            CheckReference(SharedReferences.cSortCheckHelper);
            SharedReferences.cSortCheckHelper.PerformClick();
        }

        private void currentSetElevation_Click(object sender, EventArgs e)
        {
            CheckReference(SharedReferences.cSingleSetElevation);

            SharedReferences.cSingleSetElevation.PerformClick();

        }

        private void checkedSetElevation_Click(object sender, EventArgs e)
        {
            CheckReference(SharedReferences.cCheckedSetElevation);

            SharedReferences.cCheckedSetElevation.PerformClick();

        }

        private void allSetElevation_Click(object sender, EventArgs e)
        {
            CheckReference(SharedReferences.cAllSetElevation);

            SharedReferences.cAllSetElevation.PerformClick();

        }

        private void combinedCheckAllFixedWind_Click(object sender, EventArgs e)
        {
            CheckReference(SharedReferences.cAllSetElevation);

            CheckAll();

            SharedReferences.cfixedWind.PerformClick();
        }

        private void SetMRACell(object sender, EventArgs e)
        {
            if (SharedReferences.cADGV != null)
            {
                for (int i = 0; i < SharedReferences.cADGV.Rows.Count; i++)
                {
                    if (SharedReferences.cADGV.Rows[i].Cells["PLDBID"].Value.ToString() == SharedReferences.cPPLMain.GetPole().GetValueString("Pole Number"))
                    {
                        var temp = sender.ToString();

                        if (temp == "Set MRA to Unset")
                        {
                            SharedReferences.cADGV.Rows[i].Cells["MRA"].Value = "Unset";
                        }

                        if (temp == "Set MRA to Green")
                        {
                            SharedReferences.cADGV.Rows[i].Cells["MRA"].Value = "Green";
                        }

                        if (temp == "Set MRA to Yellow")
                        {
                            SharedReferences.cADGV.Rows[i].Cells["MRA"].Value = "Yellow";
                        }

                        if (temp == "Set MRA to Red")
                        {
                            SharedReferences.cADGV.Rows[i].Cells["MRA"].Value = "Red";
                        }

                        if (temp == "Set MRA to Suggested")
                        {
                            SharedReferences.cADGV.Rows[i].Cells["MRA"].Value = SharedReferences.cMyPole.SuggestedMRA;
                        }
                    }
                }
            }
        }

        private void SetMRACheckableCell(Object sender, EventArgs e)
        {
            if (SharedReferences.cADGV != null)
            {
                for (int i = 0; i < SharedReferences.cADGV.Rows.Count; i++)
                {
                    if (SharedReferences.cADGV.Rows[i].Cells["PLDBID"].Value.ToString() == SharedReferences.cPPLMain.GetPole().GetValueString("Pole Number"))
                    {
                        var temp = sender.ToString();

                        if (temp == "Check Backbone")
                        {
                            var backbone = (DataGridViewCheckBoxCell)SharedReferences.cADGV.Rows[i].Cells["Bkbn"];

                            if (backbone.Value is true)
                            {
                                backbone.Value = false;
                            }
                            else
                            {
                                backbone.Value = true;
                            }
                        }

                        if (temp == "Check Double Wood")
                        {
                            var doubleWood = (DataGridViewCheckBoxCell)SharedReferences.cADGV.Rows[i].Cells["DubWood"];

                            if (doubleWood.Value is true)
                            {
                                doubleWood.Value = false;
                            }
                            else
                            {
                                doubleWood.Value = true;
                            }
                        }

                        if (temp == "Check Span Guy")
                        {
                            var spanGuy = (DataGridViewCheckBoxCell)SharedReferences.cADGV.Rows[i].Cells["SpanGuy"];

                            if (spanGuy.Value is true)
                            {
                                spanGuy.Value = false;
                            }
                            else
                            {
                                spanGuy.Value = true;
                            }
                        }

                        if (temp == "Check 3rd Party")
                        {
                            var thirdParty = (DataGridViewCheckBoxCell)SharedReferences.cADGV.Rows[i].Cells["Col3dPty"];

                            if (thirdParty.Value is true)
                            {
                                thirdParty.Value = false;
                            }
                            else
                            {
                                thirdParty.Value = true;
                            }
                        }
                    }
                }
            }
        }

        private void firstPoint_Click(object sender, EventArgs e)
        {
            
        }

        private void secondPoint_Click(object sender, EventArgs e)
        {

        }
    }
}