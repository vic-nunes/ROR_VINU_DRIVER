 using ASCOM.Utilities;
using ASCOM.Utilities.Interfaces;
using System;
using System.Windows.Forms;

namespace ASCOM.ROR_VINU.Dome
{
    public partial class SetupDialogForm : Form
    {
        private const string NO_PORTS_MESSAGE = "No COM ports found";
        private TraceLogger tl;

        public SetupDialogForm()
        {
            InitializeComponent();
            tl = new TraceLogger();
            InitUI();
        }

        public SetupDialogForm(TraceLogger tlDriver)
        {
            InitializeComponent();
            tl = tlDriver;
            InitUI();
        }

        private void CmdOK_Click(object sender, EventArgs e)
        {
            tl.Enabled = chkTrace.Checked;
            DomeHardware.SetTimeout((int)numericTimeout.Value);

            if (comboBoxComPort.SelectedItem is null)
            {
                tl.LogMessage("Setup OK", "New configuration values - COM Port: Not selected");
            }
            else if (comboBoxComPort.SelectedItem.ToString() == NO_PORTS_MESSAGE)
            {
                tl.LogMessage("Setup OK", "New configuration values - NO COM ports detected on this PC.");
            }
            else
            {
                DomeHardware.comPort = (string)comboBoxComPort.SelectedItem;
                tl.LogMessage("Setup OK", $"New configuration values - COM Port: {comboBoxComPort.SelectedItem}, Timeout: {numericTimeout.Value}s");
            }

            DomeHardware.WriteProfileSettings();
        }

        private void CmdCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BrowseToAscom(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("https://ascom-standards.org/");
            }
            catch (System.ComponentModel.Win32Exception noBrowser)
            {
                if (noBrowser.ErrorCode == -2147467259)
                    MessageBox.Show(noBrowser.Message);
            }
            catch (Exception other)
            {
                MessageBox.Show(other.Message);
            }
        }

        private void InitUI()
        {
            DomeHardware.ReadProfileSettings();
            chkTrace.Checked = tl.Enabled;

            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Dome";
                numericTimeout.Value = Convert.ToInt32(driverProfile.GetValue(DomeHardware.DriverProgId, DomeHardware.timeoutProfileName, string.Empty, DomeHardware.timeoutDefault));
            }

            comboBoxComPort.Items.Clear();
            using (Serial serial = new Serial())
            {
                comboBoxComPort.Items.AddRange(serial.AvailableCOMPorts);
            }

            if (comboBoxComPort.Items.Count == 0)
            {
                comboBoxComPort.Items.Add(NO_PORTS_MESSAGE);
                comboBoxComPort.SelectedItem = NO_PORTS_MESSAGE;
            }

            if (comboBoxComPort.Items.Contains(DomeHardware.comPort))
            {
                comboBoxComPort.SelectedItem = DomeHardware.comPort;
            }

            tl.LogMessage("InitUI", $"Set UI controls to Trace: {chkTrace.Checked}, COM Port: {comboBoxComPort.SelectedItem}, Timeout: {numericTimeout.Value}s");
        }

        private void SetupDialogForm_Load(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
                WindowState = FormWindowState.Normal;
            else
            {
                TopMost = true;
                Focus();
                BringToFront();
                TopMost = false;
            }
        }
    }
}