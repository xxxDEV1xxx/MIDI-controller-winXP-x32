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
        private NumericUpDown centerVelocityInput;
        private NumericUpDown midVelocityInput;
        private NumericUpDown edgeVelocityInput;
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
                    noteOffTimers[i].Tag = i; // Store index in Tag
                    noteOffTimers[i].Tick += new EventHandler(NoteOff_Tick);
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

                for (int i = 0; i < 16; i++)
                {
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
                    noteInputs[i].Value = 36 + i; // C2 to G#3, common for drum pads
                    noteInputs[i].Width = 60;
                    noteInputs[i].Tag = i;
                    noteInputs[i].BackColor = Color.Black;
                    noteInputs[i].ForeColor = TextBoxForeColor;
                    this.Controls.Add(noteInputs[i]);
                }

                centerVelocityInput = new NumericUpDown();
                centerVelocityInput.Minimum = 0;
                centerVelocityInput.Maximum = 127;
                centerVelocityInput.Value = 127;
                centerVelocityInput.Width = 60;
                centerVelocityInput.BackColor = Color.Black;
                centerVelocityInput.ForeColor = TextBoxForeColor;
                this.Controls.Add(centerVelocityInput);

                midVelocityInput = new NumericUpDown();
                midVelocityInput.Minimum = 0;
                midVelocityInput.Maximum = 127;
                midVelocityInput.Value = 80;
                midVelocityInput.Width = 60;
                midVelocityInput.BackColor = Color.Black;
                midVelocityInput.ForeColor = TextBoxForeColor;
                this.Controls.Add(midVelocityInput);

                edgeVelocityInput = new NumericUpDown();
                edgeVelocityInput.Minimum = 0;
                edgeVelocityInput.Maximum = 127;
                edgeVelocityInput.Value = 40;
                edgeVelocityInput.Width = 60;
                edgeVelocityInput.BackColor = Color.Black;
                edgeVelocityInput.ForeColor = TextBoxForeColor;
                this.Controls.Add(edgeVelocityInput);
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
                labelInputs[i].Location = new Point(gridWidth + GridMargin, inputY + i * (labelInputs[i].Height + Spacing));
                labelInputs[i].Visible = true;
                noteInputs[i].Location = new Point(gridWidth + GridMargin + labelInputs[i].Width + Spacing, inputY + i * (noteInputs[i].Height + Spacing));
                noteInputs[i].Visible = true;
            }

            int velocityY = GridMargin;
            centerVelocityInput.Location = new Point(gridWidth + GridMargin + 2 * (labelInputs[0].Width + Spacing), velocityY);
            centerVelocityInput.Visible = true;
            midVelocityInput.Location = new Point(gridWidth + GridMargin + 2 * (labelInputs[0].Width + Spacing), velocityY + centerVelocityInput.Height + Spacing);
            midVelocityInput.Visible = true;
            edgeVelocityInput.Location = new Point(gridWidth + GridMargin + 2 * (labelInputs[0].Width + Spacing), velocityY + 2 * (centerVelocityInput.Height + Spacing));
            edgeVelocityInput.Visible = true;
        }

        private void UpdateFormSize()
        {
            int gridWidth = GridSize * buttonSize + (GridSize - 1) * Spacing + 2 * GridMargin;
            int gridHeight = GridSize * buttonSize + (GridSize - 1) * Spacing + 2 * GridMargin;
            int inputWidth = 2 * (labelInputs[0].Width + Spacing) + GridMargin;
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
            int velocity = (int)midVelocityInput.Value;

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
                    velocity = (int)centerVelocityInput.Value;
                else if (distance < 0.7f)
                    velocity = (int)midVelocityInput.Value;
                else
                    velocity = (int)edgeVelocityInput.Value;
            }

            try
            {
                int noteNumber = (int)noteInputs[index].Value;
                int channel = 0;

                uint noteOn = (uint)(0x90 | channel | (noteNumber << 8) | (velocity << 16));
                int result = midiOutShortMsg(midiOutHandle, noteOn);
                if (result != 0)
                    throw new Exception(String.Format("Failed to send Note On. Error code: {0}", result));

                // Store note and channel in a custom object for NoteOff
                noteOffTimers[index].Tag = new MidiData(noteNumber, channel);
                noteOffTimers[index].Start();

                Color feedbackColor = velocity == (int)centerVelocityInput.Value ? Color.Red :
                                     velocity == (int)midVelocityInput.Value ? Color.Yellow :
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
            int index = Array.IndexOf(noteOffTimers, timer); // Find timer index

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