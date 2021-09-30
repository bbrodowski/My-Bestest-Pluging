using PPL_Lib;
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
                return PLUGIN_TYPE.DOCKED_TAB;
            }
        }

        /// <summary>
        /// Declare the name of the plugin usd for synthesizing the registry keys ect
        /// </summary>
        public String Name
        {
            get
            {
                return "MRA Colors";
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
            Guid guid = new Guid(0xed199143, 0x4947, 0x4aad, 0xbd, 0x1, 0x66, 0xa, 0x4f, 0x8, 0x61, 0xc5);
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
                //AllocConsole();

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
                            pole_dataGridView.Rows[0].Cells["Difference"].Style.BackColor = Color.Red;
                        } 
                        else
                        {
                            pole_dataGridView.Rows[0].Cells["Difference"].Style.BackColor = Color.Green;
                        }

                        if (Convert.ToDouble(pole_dataGridView.Rows[0].Cells["SetDepth"].Value) < 0)
                        {
                            pole_dataGridView.Rows[0].Cells["SetDepth"].Style.BackColor = Color.Red;
                        }
                        else
                        {
                            pole_dataGridView.Rows[0].Cells["SetDepth"].Style.BackColor = Color.Green;
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
                        spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Style.BackColor = Color.Green;
                    }

                    if (spanLength >= 300 & spanLength < 350)
                    {
                        spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Style.BackColor = Color.Yellow;
                    }

                    if (spanLength >= 350)
                    {
                        spans_dataGridView.Rows[e.RowIndex].Cells["Length"].Style.BackColor = Color.Red;
                    }
                };

                cPPLMain.BeforeEvent += (object sender, PPL_Lib.PPLMain.EVENT_TYPE pEventType, ref bool pAbortEvent) =>
                {
                    //Console.Clear();
                    Console.WriteLine(pEventType);

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

        //the toolstrip item we will add to O-Calc Pro
        ToolStripMenuItem cToolStripMenuItemButton = null;

        public void AddToMenu(PPL_Lib.PPLMain pPPLMain, System.Windows.Forms.ToolStrip pToolStrip)
        {
            //save the reference to the O-Calc Pro main
            cPPLMain = pPPLMain;

            //create the toolstrip button
            cToolStripMenuItemButton = new ToolStripMenuItem(Name);
            cToolStripMenuItemButton.AutoToolTip = true;
            cToolStripMenuItemButton.ToolTipText = Description;

            //find the dropdown menu we want to add the toolsrip button to
            int itemindex = 0;
            System.Diagnostics.Debug.Assert(pToolStrip.Items[itemindex] is ToolStripDropDownButton);
            if (pToolStrip.Items[itemindex] is ToolStripDropDownButton)
            {
                ToolStripDropDownButton tsb = pToolStrip.Items[itemindex] as ToolStripDropDownButton;
                System.Diagnostics.Debug.Assert(tsb.Text == "&File");
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, cToolStripMenuItemButton);
                tsb.DropDownItems.Insert(tsb.DropDownItems.Count - 1, new ToolStripSeparator());

                //add an event handler when the dropdown menu is opened to allow us
                //to enable or disble the toolstrip button (optional)
                tsb.DropDownOpened += tsb_DropDownOpened;
            }
            //add and event handler to detect the toolstrip button bein clicked by the user
            cToolStripMenuItemButton.Click += button_Click;
        }

        /// <summary>
        /// The menu containing our tool is being displayed.  Optionally
        /// enable or disable the toolstrip button depending on
        /// our criteria
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void tsb_DropDownOpened(object sender, EventArgs e)
        {
            bool enabled = true;
            //if (cPPLMain != null)
            //{
            //    enabled = (cPPLMain.GetMainStructure() is PPLPole);
            //}
            cToolStripMenuItemButton.Enabled = enabled;
        }

        /// <summary>
        /// The menu item was clicked by the user
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void button_Click(object sender, EventArgs e)
        {
            try
            {
                DoPluginOperation();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        /// <summary>
        /// Perform the plugin tool's operation
        /// </summary>
        private void DoPluginOperation()
        {
            try
            {
                PPLMessageBox.Show("Hello World!");
            }
            catch (Exception ex)
            {
                PPLMessageBox.Show(ex.Message, "Error in " + Name);
            }
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
    }
}
