using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace MidiButtonBoard
{
    public class MidiBoard : Form
    {
        private Button[] buttons = new Button[16];
        private IntPtr midiOutHandle = IntPtr.Zero;
        private NumericUpDown sizeInput;
        private TextBox[] labelInputs = new TextBox[16];
        private NumericUpDown[] noteInputs = new NumericUpDown[16];
        private NumericUpDown[] centerVolumeInputs = new NumericUpDown[16]; // Per-button center volume
        private NumericUpDown[] midVolumeInputs = new NumericUpDown[16];   // Per-button mid volume
        private NumericUpDown[] edgeVolumeInputs = new NumericUpDown[16]; // Per-button edge volume
        private int buttonSize = 120;
        private const int GridSize = 4;
        private const int GridMargin = 10;
        private const int Spacing = 10;
        private readonly Color ButtonColor = Color.FromArgb(11, 20, 150);
        private readonly Color TextBoxForeColor = Color.FromArgb(61, 171, 22);
        private bool[] isButtonProcessing = new bool[16];
        private Timer[] noteOffTimers = new Timer[16];

        [DllImport("winmm.dll")]
        private static extern int midiOutOpen(out IntPtr handle, int deviceID, IntPtr callback, IntPtr instance, int flags);

        [DllImport("winmm.dll")]
        private static extern int midiOutClose(IntPtr handle);

        [DllImport("winmm.dll")]
        private static extern int midiOutShortMsg(IntPtr handle, uint msg);

        [DllImport("winmm.dll")]
        private static extern int midiOutGetNumDevs();

        [DllImport("winmm.dll")]
        private static extern int midiOutGetDevCaps(int deviceID, out MIDIOUTCAPS caps, int size);

        [StructLayout(LayoutKind.Sequential)]
        private struct MIDIOUTCAPS
        {
            public ushort wMid;
            public ushort wPid;
            public uint vDriverVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szPname;
            public ushort wTechnology;
            public ushort wVoices;
            public ushort wNotes;
            public ushort wChannelMask;
            public uint dwSupport;
        }

        public MidiBoard()
        {
            this.Text = "MIDI Board";
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.BackColor = Color.Black;
            this.MaximizeBox = false;

            int midiYokeDeviceID = FindMidiYokeDevice();
            if (midiYokeDeviceID == -1)
            {
                MessageBox.Show("MIDI Yoke not found. Please install MIDI Yoke.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }

            try
            {
                int result = midiOutOpen(out midiOutHandle, midiYokeDeviceID, IntPtr.Zero, IntPtr.Zero, 0);
                if (result != 0)
                    throw new Exception(String.Format("Failed to open MIDI Yoke (Device ID: {0}). Error code: {1}", midiYokeDeviceID, result));
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("MIDI output failed: {0}\nEnsure MIDI Yoke is installed.", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }

            InitializeControls();
            LayoutControls();
            UpdateFormSize();
        }

        private int FindMidiYokeDevice()
        {
            int numDevices = midiOutGetNumDevs();
            string deviceList = "Available MIDI output devices:\n";
            int midiYokeDeviceID = -1;

            for (int i = 0; i < numDevices; i++)
            {
                MIDIOUTCAPS caps;
                int result = midiOutGetDevCaps(i, out caps, Marshal.SizeOf(typeof(MIDIOUTCAPS)));
                if (result == 0)
                {
                    deviceList += String.Format("Device {0}: {1}\n", i, caps.szPname);
                    if (caps.szPname.Contains("MIDI Yoke"))
                    {
                        midiYokeDeviceID = i;
                    }
                }
            }

            if (midiYokeDeviceID == -1)
            {
                MessageBox.Show(String.Format("MIDI Yoke not found. {0}\nPlease install MIDI Yoke.", deviceList), 
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return midiYokeDeviceID;
        }

        private void InitializeControls()
        {
            try
            {
                for (int i = 0; i < 16; i++)
                {
                    buttons[i] = new Button();
                    buttons[i].Text = (i + 1).ToString();
                    buttons[i].Tag = i;
                    buttons[i].FlatStyle = FlatStyle.Flat;
                    buttons[i].BackColor = ButtonColor;
                    buttons[i].ForeColor = Color.White;
                    buttons[i].Font = new Font("Arial", 10, FontStyle.Bold);
                    buttons[i].Click += new EventHandler(Button_Click);
                    this.Controls.Add(buttons[i]);
                    isButtonProcessing[i] = false;

                    noteOffTimers[i] = new Timer();
                    noteOffTimers[i].Interval = 100;
                    noteOffTimers[i].Tag = i;
                    noteOffTimers[i].Tick += new EventHandler(NoteOff_Tick);

                    labelInputs[i] = new TextBox();
                    labelInputs[i].Text = (i + 1).ToString();
                    labelInputs[i].Width = 60;
                    labelInputs[i].Tag = i;
                    labelInputs[i].BackColor = Color.Black;
                    labelInputs[i].ForeColor = TextBoxForeColor;
                    labelInputs[i].TextChanged += new EventHandler(LabelInput_TextChanged);
                    this.Controls.Add(labelInputs[i]);

                    noteInputs[i] = new NumericUpDown();
                    noteInputs[i].Minimum = 0;
                    noteInputs[i].Maximum = 127;
                    noteInputs[i].Value = 36 + i; // C2 to G#3
                    noteInputs[i].Width = 60;
                    noteInputs[i].Tag = i;
                    noteInputs[i].BackColor = Color.Black;
                    noteInputs[i].ForeColor = TextBoxForeColor;
                    this.Controls.Add(noteInputs[i]);

                    centerVolumeInputs[i] = new NumericUpDown();
                    centerVolumeInputs[i].Minimum = 0;
                    centerVolumeInputs[i].Maximum = 127;
                    centerVolumeInputs[i].Value = 127;
                    centerVolumeInputs[i].Width = 60;
                    centerVolumeInputs[i].Tag = i;
                    centerVolumeInputs[i].BackColor = Color.Black;
                    centerVolumeInputs[i].ForeColor = TextBoxForeColor;
                    this.Controls.Add(centerVolumeInputs[i]);

                    midVolumeInputs[i] = new NumericUpDown();
                    midVolumeInputs[i].Minimum = 0;
                    midVolumeInputs[i].Maximum = 127;
                    midVolumeInputs[i].Value = 80;
                    midVolumeInputs[i].Width = 60;
                    midVolumeInputs[i].Tag = i;
                    midVolumeInputs[i].BackColor = Color.Black;
                    midVolumeInputs[i].ForeColor = TextBoxForeColor;
                    this.Controls.Add(midVolumeInputs[i]);

                    edgeVolumeInputs[i] = new NumericUpDown();
                    edgeVolumeInputs[i].Minimum = 0;
                    edgeVolumeInputs[i].Maximum = 127;
                    edgeVolumeInputs[i].Value = 40;
                    edgeVolumeInputs[i].Width = 60;
                    edgeVolumeInputs[i].Tag = i;
                    edgeVolumeInputs[i].BackColor = Color.Black;
                    edgeVolumeInputs[i].ForeColor = TextBoxForeColor;
                    this.Controls.Add(edgeVolumeInputs[i]);
                }

                sizeInput = new NumericUpDown();
                sizeInput.Minimum = 30;
                sizeInput.Maximum = 200;
                sizeInput.Value = buttonSize;
                sizeInput.Width = 60;
                sizeInput.BackColor = Color.Black;
                sizeInput.ForeColor = TextBoxForeColor;
                sizeInput.ValueChanged += new EventHandler(SizeInput_ValueChanged);
                this.Controls.Add(sizeInput);
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Failed to initialize controls: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LayoutControls()
        {
            int clientWidth = Math.Max(this.ClientSize.Width, 300);
            int clientHeight = Math.Max(this.ClientSize.Height, 300);

            for (int i = 0; i < 16; i++)
            {
                int row = i / GridSize;
                int col = i % GridSize;
                buttons[i].Size = new Size(buttonSize, buttonSize);
                buttons[i].Location = new Point(
                    GridMargin + col * (buttonSize + Spacing),
                    GridMargin + row * (buttonSize + Spacing)
                );
                buttons[i].Visible = true;
            }

            int gridWidth = GridSize * buttonSize + (GridSize - 1) * Spacing + 2 * GridMargin;
            sizeInput.Location = new Point(gridWidth + GridMargin, GridMargin);
            sizeInput.Visible = true;

            int inputY = GridMargin + sizeInput.Height + Spacing;
            for (int i = 0; i < 16; i++)
            {
                int x = gridWidth + GridMargin;
                labelInputs[i].Location = new Point(x, inputY + i * (labelInputs[i].Height + Spacing));
                labelInputs[i].Visible = true;
                x += labelInputs[i].Width + Spacing;
                noteInputs[i].Location = new Point(x, inputY + i * (noteInputs[i].Height + Spacing));
                noteInputs[i].Visible = true;
                x += noteInputs[i].Width + Spacing;
                centerVolumeInputs[i].Location = new Point(x, inputY + i * (centerVolumeInputs[i].Height + Spacing));
                centerVolumeInputs[i].Visible = true;
                x += centerVolumeInputs[i].Width + Spacing;
                midVolumeInputs[i].Location = new Point(x, inputY + i * (midVolumeInputs[i].Height + Spacing));
                midVolumeInputs[i].Visible = true;
                x += midVolumeInputs[i].Width + Spacing;
                edgeVolumeInputs[i].Location = new Point(x, inputY + i * (edgeVolumeInputs[i].Height + Spacing));
                edgeVolumeInputs[i].Visible = true;
            }
        }

        private void UpdateFormSize()
        {
            int gridWidth = GridSize * buttonSize + (GridSize - 1) * Spacing + 2 * GridMargin;
            int gridHeight = GridSize * buttonSize + (GridSize - 1) * Spacing + 2 * GridMargin;
            int inputWidth = 5 * (labelInputs[0].Width + Spacing) + GridMargin; // Label + note + 3 volume columns
            int inputHeight = sizeInput.Height + Spacing + 16 * (labelInputs[0].Height + Spacing);
            this.ClientSize = new Size(gridWidth + inputWidth, Math.Max(gridHeight, inputHeight + GridMargin));
        }

        private void SizeInput_ValueChanged(object sender, EventArgs e)
        {
            buttonSize = (int)sizeInput.Value;
            LayoutControls();
            UpdateFormSize();
        }

        private void LabelInput_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;
            int index = (int)txt.Tag;
            buttons[index].Text = txt.Text;
        }

        private void Button_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            int index = (int)btn.Tag;

            if (isButtonProcessing[index])
                return;

            isButtonProcessing[index] = true;

            MouseEventArgs mouseArgs = e as MouseEventArgs;
            int velocity = (int)midVolumeInputs[index].Value;

            if (mouseArgs != null)
            {
                float centerX = btn.Width / 2f;
                float centerY = btn.Height / 2f;
                float clickX = mouseArgs.X;
                float clickY = mouseArgs.Y;
                float dx = (clickX - centerX) / (btn.Width / 2f);
                float dy = (clickY - centerY) / (btn.Height / 2f);
                float distance = (float)Math.Sqrt(dx * dx + dy * dy);

                if (distance < 0.3f)
                    velocity = (int)centerVolumeInputs[index].Value;
                else if (distance < 0.7f)
                    velocity = (int)midVolumeInputs[index].Value;
                else
                    velocity = (int)edgeVolumeInputs[index].Value;
            }

            try
            {
                int noteNumber = (int)noteInputs[index].Value;
                int channel = 0;

                uint noteOn = (uint)(0x90 | channel | (noteNumber << 8) | (velocity << 16));
                int result = midiOutShortMsg(midiOutHandle, noteOn);
                if (result != 0)
                    throw new Exception(String.Format("Failed to send Note On. Error code: {0}", result));

                noteOffTimers[index].Tag = new MidiData(noteNumber, channel);
                noteOffTimers[index].Start();

                Color feedbackColor = velocity == (int)centerVolumeInputs[index].Value ? Color.Red :
                                     velocity == (int)midVolumeInputs[index].Value ? Color.Yellow :
                                     Color.Green;
                btn.BackColor = feedbackColor;

                Timer colorTimer = new Timer();
                colorTimer.Interval = 100;
                colorTimer.Tick += new EventHandler(delegate(object s, EventArgs args)
                {
                    btn.BackColor = ButtonColor;
                    colorTimer.Stop();
                    colorTimer.Dispose();
                });
                colorTimer.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("MIDI error: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isButtonProcessing[index] = false;
            }
        }

        private class MidiData
        {
            public int NoteNumber;
            public int Channel;

            public MidiData(int noteNumber, int channel)
            {
                NoteNumber = noteNumber;
                Channel = channel;
            }
        }

        private void NoteOff_Tick(object sender, EventArgs e)
        {
            Timer timer = (Timer)sender;
            MidiData data = timer.Tag as MidiData;
            int index = Array.IndexOf(noteOffTimers, timer);

            if (data != null && index >= 0)
            {
                try
                {
                    uint noteOff = (uint)(0x80 | data.Channel | (data.NoteNumber << 8) | (0 << 16));
                    int result = midiOutShortMsg(midiOutHandle, noteOff);
                    if (result != 0)
                        throw new Exception(String.Format("Failed to send Note Off. Error code: {0}", result));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(String.Format("MIDI error in NoteOff: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            timer.Stop();
            if (index >= 0)
                isButtonProcessing[index] = false;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (midiOutHandle != IntPtr.Zero)
                midiOutClose(midiOutHandle);
            foreach (Timer timer in noteOffTimers)
                timer.Dispose();
            base.OnClosing(e);
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new MidiBoard());
        }
    }
}