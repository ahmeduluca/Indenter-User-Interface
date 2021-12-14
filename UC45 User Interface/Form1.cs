using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Timers;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using MathNet.Numerics;
using NationalInstruments;
using NationalInstruments.DAQmx;
using Task = System.Threading.Tasks.Task;
using System.Diagnostics;
using System.Windows.Forms.DataVisualization.Charting;
using System.Collections;
using Python.Runtime;

namespace UC45_User_Interface
{
    public partial class Form1 : Form
    {
        private AnalogMultiChannelReader analogInReader;
        private AsyncCallback myAsyncCallback;
        private NationalInstruments.DAQmx.Task myTask;
        private NationalInstruments.DAQmx.Task runningTask;
        private AnalogWaveform<double>[] data;
        private AnalogWaveform<double>[] dataR;
        public Form1()
        {
            InitializeComponent();
        }
        int defCount = 0;
        string[] connString = { "E", "$1E", "$1E", "$0E", "$0E" };
        string[] respConnString = { "Y", "X", "X", "X", "X", "X", "Y", "X", "'d3'gx" };
        Boolean isConnected = false;
        public string vapplied = "";
        double kp = Convert.ToDouble(Properties.Settings.Default["Piezo"]);
        double vnull = Convert.ToDouble(Properties.Settings.Default["Vmax"]);
        double vmin = Convert.ToDouble(Properties.Settings.Default["Vmin"]);
        double vmax = Convert.ToDouble(Properties.Settings.Default["Vmax"]);
        double zrange = (Convert.ToDouble(Properties.Settings.Default["Vmax"]) -
            Convert.ToDouble(Properties.Settings.Default["Vmin"])) * 20 * Convert.ToDouble(Properties.Settings.Default["Piezo"]);
        /*
 Connection
 $1E  X
 $1E  X
 $0E  X
 $0E  X
 V1E  Y
 $0E  X
 R1E  ..... Mode gir
 $0E  X
 R3E  Y
 R1E  ...... Mode gir
 R2E  Y
 Paramaters       
 Vapp gönder      X
 PID gönder       X
 Filtre çeşit     X
 Cuttoff1 gönder  X
 cuttoff2 gönder  X
 Vmin gönder      X
 Vmax gönder      X
 Delta +-S gönder X

  Her veri arası 50ms 


   */
        //protected override void Dispose(bool disposing)
        //{
        //    if (disposing)
        //    {
        //        if (components != null)
        //        {
        //            components.Dispose();
        //        }
        //        if (myTask != null)
        //        {
        //            runningTask = null;
        //            myTask.Dispose();
        //        }
        //    }
        //    base.Dispose(disposing);
        //}
        int daqprob = 0;
        List<double> datas = new List<double>();
        private void AnalogInCallback(IAsyncResult ar)
        {
            try
            {
                if (runningTask != null && runningTask == ar.AsyncState)
                {
                    data = analogInReader.EndReadWaveform(ar);
                    for (int i = 0; i < data.Length; i++)
                    {
                        datas[i] = data[i].Samples.Last().Value;
                    }
                    analogInReader.BeginMemoryOptimizedReadWaveform(10,
                        myAsyncCallback, runningTask, data);
                }
            }
            catch (DaqException exception)
            {
                analogInReader.EndReadWaveform(ar);
                timer2.Stop();
                MessageBox.Show(exception.Message);
                runningTask = null;
                myTask.Dispose();
                startButton.Enabled = true;
                stopButton.Enabled = false;
                if (daqprob == 0)
                {
                    daqprob = 1;
                    button9.PerformClick();
                }
            }

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            myAsyncCallback = new AsyncCallback(AnalogInCallback);
            physicalChannelComboBox.Items.AddRange(DaqSystem.Local.GetPhysicalChannels(PhysicalChannelTypes.AI, PhysicalChannelAccess.External));
            tempChan.SelectedIndex = 0;
            gageChan.SelectedIndex = 0;
            if (physicalChannelComboBox.Items.Count > 0)
                physicalChannelComboBox.SelectedIndex = 0;
            tcNo.SelectedIndex = 0;
            bridgeBox.SelectedIndex = 0;
            excBox.SelectedIndex = 0;
            rtdBox.SelectedIndex = 1;
            rtdExcType.SelectedIndex = 0;
            comboBox2.SelectedIndex = 1;
            chart1.MouseWheel += chart1_MouseWheel;
            pathsave = Properties.Settings.Default["PathtoSave"].ToString();
            if (!pathsave.Contains(DateTime.Today.ToString("dd-MM-yyyy")))
            {
                string zaman = @"\" + DateTime.Today.ToString("dd-MM-yyyy");
                pathsave = pathsave.Remove(pathsave.Length - 11);
                pathsave += zaman;
                Directory.CreateDirectory(pathsave);
                Properties.Settings.Default["PathtoSave"] = pathsave;
                Properties.Settings.Default.Save();
            }
            g1.Text = Properties.Settings.Default["GageCo"].ToString();
            textBox17.Text = Properties.Settings.Default["Piezo"].ToString();
            textBox1.Text = Properties.Settings.Default["Vapp"].ToString();
            textBox8.Text = Properties.Settings.Default["Vmax"].ToString();
            textBox9.Text = Properties.Settings.Default["Vmin"].ToString();
            textBox2.Text = Properties.Settings.Default["Kp"].ToString();
            textBox3.Text = Properties.Settings.Default["Ki"].ToString();
            textBox4.Text = Properties.Settings.Default["Kd"].ToString();
            textBox10.Text = Properties.Settings.Default["Z"].ToString();
            zpos = Convert.ToDouble(textBox10.Text);
            textBox5.Text = Properties.Settings.Default["deltaS"].ToString();
            comboBox1.SelectedIndex = Convert.ToInt16(Properties.Settings.Default["filterType"]);
            //tempChan.SelectedIndex = Convert.ToInt16(Properties.Settings.Default["TempCon"]);
            loadThres.Text = Properties.Settings.Default["PressThres"].ToString();
            pressure = Convert.ToDouble(loadThres.Text);
            appSpeed.Text = Properties.Settings.Default["AppSpeed"].ToString();
            textBox7.Text = Properties.Settings.Default["cutOff1"].ToString();
            textBox6.Text = Properties.Settings.Default["cutOff2"].ToString();
            VappNumeric.Value = Convert.ToDecimal(Properties.Settings.Default["Vapp"].ToString()) * 1000000;
            VappNumeric.Increment = Convert.ToDecimal(Properties.Settings.Default["VoltInc"].ToString()) * 1000;
            textBox32.Text = Properties.Settings.Default["VoltInc"].ToString();
            textBox33.Text = Properties.Settings.Default["PosInc"].ToString();
            //verticalProgressBar1.Value = Convert.ToInt16(200 * (Convert.ToDouble(textBox1.Text) + 1) / 17);
            //gageChan.SelectedIndex = Convert.ToInt16(Properties.Settings.Default["GageChan"]);
            motorPos.Text = Properties.Settings.Default["motpos"].ToString();
            stepAng.Text = Properties.Settings.Default["ang"].ToString();
            stepSpeed.Text = Properties.Settings.Default["motorspeed"].ToString();
            motorMax.Text = Properties.Settings.Default["motormax"].ToString();
            threadPitch.Text = Properties.Settings.Default["pitch"].ToString();
            motpos = Convert.ToDouble(motorPos.Text) * 1000000;
            motmax = Convert.ToDouble(motorMax.Text) * 1000000;
            motspeed = Convert.ToDouble(stepSpeed.Text);
            stepang = Convert.ToDouble(stepAng.Text);
            micstepBox.SelectedIndex = Convert.ToInt16(Properties.Settings.Default["micros"].ToString());
            micstep = Convert.ToInt16(micstepBox.SelectedItem);
            pitch = Convert.ToDouble(threadPitch.Text) * 1000000;
            stepIncbox.Text = Convert.ToString(pitch * stepang / (360 * micstep));
            stepinc = pitch * stepang / (360 * micstep);
            motfreq = 1000000 / (motspeed / stepinc);
            holdDur.Text = Properties.Settings.Default["holdDur"].ToString();
            forceCo.Text = Properties.Settings.Default["forceCo"].ToString();
            xTour.Text = Properties.Settings.Default["xenc"].ToString();
            yTour.Text = Properties.Settings.Default["yenc"].ToString();
            xinc = Convert.ToDouble(xTour.Text);
            yinc = Convert.ToDouble(yTour.Text);
            selectedHeater.SelectedIndex = 0;
            tipName.Text = Properties.Settings.Default["tipName"].ToString();
            sampleName.Text = Properties.Settings.Default["sampleName"].ToString();
            procedureName.Text = Properties.Settings.Default["procedureName"].ToString();
            if (Properties.Settings.Default["Mode"].ToString() == "1")
            {
                checkBox10.Checked = true;
            }
            else
            {
                checkBox10.Checked = false;
            }
            if (Properties.Settings.Default["inverse"].ToString() == "0")
            {
                checkBox2.Checked = false;
            }
            else
            {
                checkBox2.Checked = true;
            }
            if (Properties.Settings.Default["Enable"].ToString() == "0")
            {
                checkBox1.Checked = false;
            }
            else
            {
                checkBox1.Checked = true;
            }
            this.Width = 769 + sizeadd;
            cjcSourceComboBox.SelectedIndex = 1;
            thermocoupleTypeComboBox.SelectedIndex = 3;
            serialPort1.Parity = Parity.None;
            serialPort1.StopBits = StopBits.One;
            serialPort1.DataBits = 8;
            serialPort1.BaudRate = 57600;
            stopExp.Enabled = false;
            groupBox5.Enabled = false;
            groupBox8.Enabled = false;
            groupBox19.Enabled = false;
            giveWhile.SelectionMode = SelectionMode.One;
            heatChannels[0] = -1;
            heatChannels[1] = -1;
            heatChannels[2] = -1;
        }
        string response;
        double extdata = 0;
        //int byteCounter = 0;
        int zCount = 0;
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            response = serialPort1.ReadLine();
        }
        string texts = "";
        bool sending = false;
        string directMes = "";
        private void Send(string text, string prefix)
        {
            sending = true;
            string temp = text;
            string id = "";
            if (isConnected && prefix != "")
            {
                temp = temp.Insert(text.Length, "E");
                temp = prefix.Insert(prefix.Length, temp);
                int length = temp.Length;
                directMes = temp;
                id = $"5{length:D2}00";
                char[] buffer = new char[id.Length];
                for (int i = 0; i < id.Length; i++)
                {
                    buffer[i] = Convert.ToChar(id[i]);
                }
                serialPort2.Write(buffer, 0, id.Length);
                texts = temp;
                directpass = 1;
            }
            else if (isConnected && prefix == "" && serialPort2.IsOpen)
            {
                serialPort2.Write(temp);
                lastwords = temp;
            }
            sending = false;
        }
        int p = -1;
        DialogResult resulttextform = DialogResult.Ignore;
        private string TextFormat(string text, string setting, double min, double max)
        {
            //string temp = String.Format("{0:0.#####0}",text);//yazılanı aldık.
            //int tempL = 0;
            double tem1 = 0;
            int tem2 = 0;
            decimal onda = 0;
            string ondas = "";
            int eks = 0;
            if (text != "")
            {
                try
                {
                    if (text.Contains("-"))
                    {
                        text = text.Remove(0, 1);
                        eks = 1;
                    }
                    tem2 = (int)Convert.ToDouble(text);
                    tem1 = Convert.ToDouble(text);
                    onda = (decimal)tem1 - tem2;
                    if (onda != 0)
                    {
                        onda = (int)(onda * 1000000);
                        ondas = string.Format("{0:D}.{1:D6}", tem2, Convert.ToInt32(onda));
                    }
                    else
                    {
                        ondas = string.Format("{0:D}.000000", tem2);
                    }
                    if (eks == 1)
                    {
                        ondas = ondas.Insert(0, "-");
                    }
                    tem1 = Convert.ToDouble(ondas);
                }
                catch (Exception ex)
                {
                    label11.Text = ex.Message;
                    ondas = "0.000000";
                }
            }
            else
            {
                ondas = "0.000000";
            }
            if ((tem1 > max || tem1 < min))
            {
                if (setting != "")
                {
                    ondas = Properties.Settings.Default[setting].ToString();
                    resulttextform = MessageBox.Show("Out of range! If this message has been shown during experiment, Please press ABORT!", "Limits of System", MessageBoxButtons.AbortRetryIgnore);
                }
                else
                {
                    if (resulttextform == DialogResult.Abort)
                    {
                        emCounter = 0;
                    }
                }
            }
            return ondas;
        }
        double vol = 0;
        bool voltageplus = false;
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (vol > Convert.ToDouble(textBox1.Text))
                {
                    voltageplus = true;
                }
                vapplied = textBox1.Text;
                if (!textBox10.Enabled && textBox1.Text != "" && resulttextform == DialogResult.Ignore)
                {
                    textBox10.Text = ((vnull - Convert.ToDouble(vapplied)) * 20 * kp).ToString();
                }
                else if (!textBox10.Enabled && textBox1.Text != "" && resulttextform == DialogResult.Ignore)
                {
                    if (voltageplus)
                    {
                        textBox10.Text = TextFormat((-Math.Pow(vol, 2) * alpha + vol * betha + 73.731).ToString(), "Z", -90, 90);
                    }
                    else
                    {
                        textBox10.Text = TextFormat((Math.Pow(vol, 2) * alpha + vol * betha + 73.731).ToString(), "Z", -90, 90);
                    }
                }
                resulttextform = DialogResult.Ignore;
            }
            catch (Exception ex)
            {
                label11.Text = ex.Message;
                textBox1.Text = "0.000000";
            }

        }
        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            labelVmax.Text = TextFormat(textBox8.Text, "", -1, 7.5);
            //M.....E
            //genliği -1,7.5 e kadar regex
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            labelVmin.Text = TextFormat(textBox9.Text, "", -1, 7.5);
            //N.....E
            //genliği -1,7.5 e kadar regex
        }
        int emCounter = 1;
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            //Herşeyi durdur
            // open loop
            // mod internal a al 
            //W0.000000E gönder 0 volta çek
            try
            {
                emCounter++;
                if (emCounter % 2 == 0 && checkBox15.Checked)
                {
                    Send("1", "T"); //T1E
                    Thread.Sleep(55);
                    Send("0.000000", "W"); //W0.000000E
                    textBox1.Text = "0.000000";
                    Thread.Sleep(55);
                    Send("0", "B");
                    if (comread == "X")
                    {
                        MessageBox.Show("Emergency Stop! Voltage has been set to 0!" + Environment.NewLine +
                            "Click again emergency button to activate program.", "Emergency", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    groupBox1.Enabled = false;
                    groupBox2.Enabled = false;
                    groupBox5.Enabled = false;
                    groupBox8.Enabled = false;
                    groupBox11.Enabled = false;
                    checkBox1.Enabled = false;
                }
                else if (emCounter % 2 == 0 && !checkBox15.Checked)
                {
                    serialPort2.Write("Q0000");
                    MessageBox.Show("Emergency Stop! Voltage has been set to 0!" + Environment.NewLine +
                            "Click again emergency button to activate program.", "Emergency", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    groupBox1.Enabled = false;
                    groupBox2.Enabled = false;
                    groupBox5.Enabled = false;
                    groupBox8.Enabled = false;
                    groupBox11.Enabled = false;
                    checkBox1.Enabled = false;
                }
                else
                {
                    groupBox1.Enabled = true;
                    groupBox2.Enabled = true;
                    groupBox5.Enabled = true;
                    groupBox8.Enabled = true;
                    groupBox11.Enabled = true;
                    checkBox1.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                label11.Text = ex.Message;
                MessageBox.Show("Communication Lost! Please unplug high voltage connections & take sample to safe distance!", "Emergency", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }   // tekrar basılmadığı sürece herşey kapalı
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            // close loop kapat
            // mode internal a al
            //
            // W7.500000E basılmadığı sürece
            try
            {
                emCounter++;
                if (emCounter % 2 == 0 && checkBox15.Checked)
                {
                    Send("1", "T"); //T1E
                    Thread.Sleep(55);
                    Send("7.500000", "W");
                    textBox1.Text = "7.500000";
                    Thread.Sleep(55);
                    Send("0", "B"); //B0E
                    groupBox1.Enabled = false;
                    groupBox2.Enabled = false;
                    groupBox5.Enabled = false;
                    groupBox8.Enabled = false;
                    groupBox11.Enabled = false;
                    checkBox1.Enabled = false;
                    MessageBox.Show("Emergency Stop! Actuator has been pulled full-back!" + Environment.NewLine +
                            "Click again emergency button to activate program.", "Emergency", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (emCounter % 2 == 0 && !checkBox15.Checked)
                {
                    serialPort2.Write("*0000");
                    autopass = 1;
                    //verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                    groupBox1.Enabled = false;
                    groupBox2.Enabled = false;
                    groupBox5.Enabled = false;
                    groupBox8.Enabled = false;
                    groupBox11.Enabled = false;
                    checkBox1.Enabled = false;
                    MessageBox.Show("Emergency Stop! Actuator has been pulled full-back!" + Environment.NewLine +
                            "Click again emergency button to activate program.", "Emergency", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    groupBox1.Enabled = true;
                    groupBox2.Enabled = true;
                    groupBox5.Enabled = true;
                    groupBox8.Enabled = true;
                    groupBox11.Enabled = true;
                    checkBox1.Enabled = true;

                }
            }
            catch (Exception ex)
            {
                label11.Text = ex.Message;
                MessageBox.Show("Communication Lost! Please unplug high voltage connections and take sample to safe distance!",
                    "Emergency", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = TextFormat(textBox2.Text, "Kp", 0, 10);
            //P.......E
            //p 0,10 arası regex
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox3.Text = TextFormat(textBox3.Text, "Ki", 0, (double)32000);
            //I.....E
            // 0,32000 arası regex
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            textBox4.Text = TextFormat(textBox4.Text, "Kd", 0, 0.01);
            //D....E
            //0,0.010000 arası
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            textBox5.Text = TextFormat(textBox5.Text, "deltaS", 0, 100);
            //Radia button durumuna göre sayının durumunu belirle
            //G.....E
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                groupBox1.Enabled = true;
                Send("1", "B");
                Properties.Settings.Default["Enable"] = "1";
                Properties.Settings.Default.Save();
            }
            else
            {
                groupBox1.Enabled = false;
                Send("0", "B");
                Properties.Settings.Default["Enable"] = "0";
                Properties.Settings.Default.Save();
            }
            //B0E unchecked
            //B1E checked

        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Send(comboBox1.SelectedIndex.ToString(), "C");
            Properties.Settings.Default["filterType"] = comboBox1.SelectedIndex.ToString();
            Properties.Settings.Default.Save();
            // index 0 C0E gönder
            // index 1 C1E ....
            // index 2 .
            // index 3
            // index 4
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            textBox7.Text = TextFormat(textBox7.Text, "cutOff1", 0, 1000000);
            // F....E 
            // sınır +
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            textBox6.Text = TextFormat(textBox6.Text, "cutOff2", 0, 1000000);
            // S....E
            //sınır +
        }

        private void UsePy()
        {
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic scipy = Py.Import("scipy.optimize.curve_fit");
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            clickinc = 0;
            if (zCount % 2 != 0)
            {
                VappNumeric.Increment = (decimal)Convert.ToDouble(Convert.ToDouble(Properties.Settings.Default["VoltInc"].ToString()) * 1000);
                VappNumeric.Value = (decimal)(Convert.ToDouble(textBox1.Text) * 1000000);
                textBox1.Enabled = true;
                textBox10.Enabled = false;
                button6.Text = "Position Control";
            }
            else
            {
                VappNumeric.Increment = (decimal)Convert.ToDouble(Properties.Settings.Default["PosInc"].ToString());
                VappNumeric.Value = (decimal)(Convert.ToDouble(textBox10.Text) * 1000);
                textBox1.Enabled = false;
                textBox10.Enabled = true;
                button6.Text = "Voltage Control";
            }
            zCount++;
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                textBox8.Text = labelVmax.Text;
                Send(labelVmax.Text, "M");
                vmax = Convert.ToDouble(textBox8.Text);
                zrange = kp * 20 * (vnull - vmin);
                Properties.Settings.Default["Vmax"] = textBox8.Text;
                Properties.Settings.Default.Save();
                textBox1.Text = TextFormat(textBox1.Text, "Vapp", vmin, vmax);
                //verticalProgressBar1.Value = Convert.ToInt16(200 * (Convert.ToDouble(textBox1.Text) + 1) / 17);
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                textBox9.Text = labelVmin.Text;
                Send(labelVmin.Text, "N");
                vmin = Convert.ToDouble(textBox9.Text);
                zrange = kp * 20 * (vnull - vmin);
                Properties.Settings.Default["Vmin"] = textBox9.Text;
                Properties.Settings.Default.Save();
                textBox1.Text = TextFormat(textBox1.Text, "Vapp", vmin, vmax);
            }
        }
        string sendz;
        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                try
                {
                    labelZ.Text = TextFormat(textBox10.Text, "Z", (vmin - (vmax - vnull) + 1) * 20 * kp, (vnull - vmin) * 20 * kp);
                    textBox10.Text = labelZ.Text;
                    if (zrange >= Convert.ToDouble(textBox10.Text))//hysteresis on/off-> cb8.
                    {
                        sendz = ((vnull * 20 - Convert.ToDouble(textBox10.Text) / kp) / 20).ToString();
                        textBox1.Text = TextFormat(sendz, "Vapp", vmin, vnull);
                    }
                    else
                    {
                        sendz = textBox1.Text;
                    }
                    //verticalProgressBar1.Value = Convert.ToInt16(200 * (Convert.ToDouble(textBox1.Text) + 1) / 17);
                    clickinc = 0;
                    Send(TextFormat(sendz, "Vapp", -1, 7.5), "W");
                    Properties.Settings.Default["Z"] = textBox10.Text;
                    Properties.Settings.Default["Vapp"] = textBox1.Text;
                    Properties.Settings.Default.Save();
                    VappNumeric.Value = (decimal)(Convert.ToDouble(textBox10.Text) * 1000);
                }
                catch (Exception ex)
                {
                    label11.Text = ex.Message;

                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                try
                {
                    labelVapp.Text = TextFormat(textBox1.Text, "Vapp", vmin, vnull);
                    textBox1.Text = labelVapp.Text;
                    // verticalProgressBar1.Value = Convert.ToInt16(200 * (Convert.ToDouble(textBox1.Text) + 1) / 17);
                    clickinc = 0;
                    Send(labelVapp.Text, "W");
                    labelZ.Text = textBox10.Text;
                    Properties.Settings.Default["Vapp"] = textBox1.Text;
                    Properties.Settings.Default["Z"] = textBox10.Text;
                    Properties.Settings.Default.Save();
                    VappNumeric.Value = (decimal)(Convert.ToDouble(textBox1.Text) * 1000000);
                }
                catch (Exception ex)
                {
                    label11.Text = ex.Message;

                }

            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                labelDeltaS.Text = textBox5.Text;
                if (checkBox2.Checked)
                {
                    Send(textBox5.Text, "G-");
                }
                else
                {
                    Send(textBox5.Text, "G");
                }
                Properties.Settings.Default["deltaS"] = textBox5.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                labelKp.Text = textBox2.Text;
                Send(labelKp.Text, "P");
                Properties.Settings.Default["Kp"] = textBox2.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                labelKi.Text = textBox3.Text;
                Send(labelKi.Text, "I");
                Properties.Settings.Default["Ki"] = textBox3.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                labelKd.Text = textBox4.Text;
                Send(labelKd.Text, "D");
                Properties.Settings.Default["Kd"] = textBox4.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                labelCutOff1.Text = textBox7.Text;
                Send(labelCutOff1.Text, "F");
                Properties.Settings.Default["cutOff1"] = textBox7.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                labelCutOff2.Text = textBox6.Text;
                Send(labelCutOff2.Text, "S");
                Properties.Settings.Default["cutOff2"] = textBox6.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                Send(textBox5.Text, "G-");
                Properties.Settings.Default["inverse"] = "1";
                Properties.Settings.Default.Save();
            }
            else
            {
                Send(textBox5.Text, "G");
                Properties.Settings.Default["inverse"] = "0";
                Properties.Settings.Default.Save();
            }
        }
        double zpos = 0;
        bool zplus = false;
        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            //textBox10.Text = TextFormat(textBox10.Text, "Z", (vmin - (vmax - vnull)+20) * kp, (vnull-vmin)*kp);
            //(vmin - (vmax - vnull) + 20) * kp, (vnull + 20) * kp
            //, "Z", (vmin - (vmax - vnull)+20) * kp, (vnull - vmin) * kp)
            try
            {
                if (zpos > Convert.ToDouble(textBox10.Text))
                {
                    zplus = true;
                }
                else
                {
                    zplus = false;
                }
                zpos = Convert.ToDouble(textBox10.Text);
                if (!textBox1.Enabled && resulttextform == DialogResult.Ignore && zrange >= Convert.ToDouble(textBox10.Text))
                {
                    //textBox1.Text = TextFormat(((vnull - (Convert.ToDouble(textBox10.Text) / kp)) / 20).ToString(), "Vapp", vmin/20,vmax/20);
                    textBox1.Text = ((vnull * 20 - (Convert.ToDouble(textBox10.Text) / kp)) / 20).ToString();
                }
                else if (!textBox1.Enabled && resulttextform == DialogResult.Ignore)
                {
                    if (zplus)
                    {
                        textBox1.Text = TextFormat(((betha - (Math.Sqrt(Math.Pow(betha, 2) + 4 * alpha * (73.731 - zpos)))) / (40 * alpha)).ToString(), "Vapp", -1, 7.5);
                    }
                    else
                    {
                        textBox1.Text = TextFormat(((-betha - (Math.Sqrt(Math.Pow(betha, 2) - 4 * alpha * (73.731 - zpos)))) / (40 * alpha)).ToString(), "Vapp", -1, 7.5);
                    }
                }
                resulttextform = DialogResult.Ignore;
            }
            catch (Exception ex)
            {
                label11.Text = ex.Message;
                textBox10.Text = "0.000000";
            }

        }


        bool combochange = false;
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex != 0)
            {
                groupBox3.Enabled = true;
                groupBox4.Enabled = true;
                groupBox5.Text = "Step" + comboBox2.SelectedIndex.ToString();
                groupBox6.Enabled = false;
                combochange = true;
                if (depth.Count >= comboBox2.SelectedIndex)
                {
                    textBox11.Text = depth[comboBox2.SelectedIndex - 1];
                    textBox12.Text = speed[comboBox2.SelectedIndex - 1];
                    indXpst.Text = Xpos[comboBox2.SelectedIndex - 1];
                    indYpst.Text = Ypos[comboBox2.SelectedIndex - 1];
                    dtRamp.Text = dTemp[comboBox2.SelectedIndex - 1];
                    dtSpeed.Text = dTspeed[comboBox2.SelectedIndex - 1];
                    dtAmp.Text = dTamp[comboBox2.SelectedIndex - 1];
                    dtF.Text = dTfreq[comboBox2.SelectedIndex - 1];
                    dtPhase.Text = dTlag[comboBox2.SelectedIndex - 1];
                    retractStep.Checked = Convert.ToBoolean(retStep[comboBox2.SelectedIndex - 1]);
                    holdDur.Text = retHold[comboBox2.SelectedIndex - 1];
                    holdPercent.Value = holdAt[comboBox2.SelectedIndex - 1];
                    for (int i = 0; i < 3; i++)
                    {
                        if (i == dTwhen[comboBox2.SelectedIndex - 1])
                        {
                            giveWhile.SetItemChecked(dTwhen[comboBox2.SelectedIndex - 1], true);
                        }
                        else
                        {
                            giveWhile.SetItemCheckState(i, CheckState.Unchecked);
                        }
                    }
                }
                else
                {
                    textBox11.Text = "";
                    textBox12.Text = "";
                    indXpst.Text = "";
                    indYpst.Text = "";
                    dtRamp.Text = "";
                    dtSpeed.Text = "";
                    dtAmp.Text = "";
                    dtF.Text = "";
                    dtPhase.Text = "";
                    holdDur.Text = "";
                    holdPercent.Value = 0;
                    giveWhile.SetItemCheckState(0, CheckState.Unchecked);
                    giveWhile.SetItemCheckState(1, CheckState.Unchecked);
                    giveWhile.SetItemCheckState(2, CheckState.Unchecked);
                    retractStep.Checked = false;
                }
                if (duration.Count >= comboBox2.SelectedIndex)
                {
                    textBox13.Text = duration[comboBox2.SelectedIndex - 1];
                    indXpst.Text = Xpos[comboBox2.SelectedIndex - 1];
                    indYpst.Text = Ypos[comboBox2.SelectedIndex - 1];
                }
                else
                {
                    textBox13.Text = "";
                    indXpst.Text = "";
                    indYpst.Text = "";
                }
                if (amplitude.Count >= comboBox2.SelectedIndex)
                {
                    textBox14.Text = amplitude[comboBox2.SelectedIndex - 1];
                    indXpst.Text = Xpos[comboBox2.SelectedIndex - 1];
                    indYpst.Text = Ypos[comboBox2.SelectedIndex - 1];
                }
                else
                {
                    textBox14.Text = "";
                    indXpst.Text = "";
                    indYpst.Text = "";
                }
                if (frequency.Count >= comboBox2.SelectedIndex)
                {
                    textBox15.Text = frequency[comboBox2.SelectedIndex - 1];
                }
                else
                {
                    textBox15.Text = "";
                }
                if (interval.Count >= comboBox2.SelectedIndex)
                {
                    textBox16.Text = interval[comboBox2.SelectedIndex - 1];

                }
                else
                {
                    textBox16.Text = "";
                }
                if (box5.Count >= comboBox2.SelectedIndex)
                {
                    checkBox5.Checked = box5[comboBox2.SelectedIndex - 1];
                }
                else
                {
                    checkBox5.Checked = false;
                }

                if (box11.Count >= comboBox2.SelectedIndex)
                {
                    checkBox11.Checked = box11[comboBox2.SelectedIndex - 1];
                }
                else
                {
                    checkBox11.Checked = false;
                }
                if (!checkBox11.Checked && !checkBox5.Checked)
                {
                    checkBox3.Checked = true;
                }
                else
                {
                    checkBox3.Checked = false;
                }
                combochange = false;
            }
            else if (comboBox2.SelectedIndex == 0)
            {
                groupBox3.Enabled = false;
                groupBox4.Enabled = false;
                groupBox6.Enabled = true;
                groupBox5.Text = "Approach";
            }
        }
        bool presscon = false;
        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
        }
        bool tempcon = false;
        private string Experimentlog(string now)
        {
            string format = "";
            try
            {
                // DATE TIME & Medium Temperature and Humidity
                format = "Indentation Experiment" + " " + now + Environment.NewLine +
                    "(DHT11 DATA) Medium Temperature: " + humTemp.Text + " Celcius - Medium Relative Humidity: %" + relHum.Text + Environment.NewLine;

                //Tip & Sample Info
                format = format + "Tip: " + tipName.Text + "  Sample: " + sampleName.Text + "\t Experiment Procedure: " + procedureName.Text + Environment.NewLine;

                //BALANCE &DEVIATION &VIB CHECK
                format = format + "\nBalance of Indenter: " + Environment.NewLine + "Trapezoid Screw Holder (Head) (phi, theta) :" + balance[0, 0].ToString() + ", "
                    + balance[0, 1].ToString() + Environment.NewLine + "Indenter Holder (Moving Part) (phi, theta): " + balance[1, 0].ToString() + ", " +
                    balance[1, 1].ToString() + Environment.NewLine + "Sample Holder (Base) (phi, theta): " + balance[2, 0].ToString() + ", "
                    + balance[2, 1].ToString() + Environment.NewLine + "Sample Deviation (9 pt grid check): Sx = " + surfDevX.Text + " Sy = " + surfDevY.Text +
                    //"\nMax Vibration Frequency & Amplitude:"+vibrationModes.Rows[0].Cells[0].Value.ToString()+", "+ vibrationModes.Rows[0].Cells[1].Value.ToString() + !!TAMAMLANMADI HENUZ
                    Environment.NewLine;

                //STEP MOTOR
                format = format + $"\nStep Motor Position:{motorPos.Text} mm\nParameters:" + Environment.NewLine + "Stepper Angle: " + stepAng.Text + " Thread Pitch Length: "
                    + threadPitch.Text + Environment.NewLine + "Driver Micro Step(*256 movement resolution): " + micstepBox.SelectedItem.ToString() + Environment.NewLine +
                    "Calculated Step Control: " + stepIncbox.Text + " nm" + Environment.NewLine;

                //XYMOTOR 
                format = format + "\nXY Motor Settings: (KT180 DC)" + Environment.NewLine + "X Position (mm): " + xkon.ToString() + "; Y Position (mm): " + ykon.ToString() +
                    Environment.NewLine + "X Encoder Coefficient: (um/#) " + xinc.ToString() + "; Y Encoder Coefficient: (um/#) " + yinc.ToString() + Environment.NewLine;

                //ACTUATOR
                format = format + "\nActuator APA60S parameters:" + Environment.NewLine + "Piezo Constant: " + kp.ToString() +
                    "um/V; Applied Voltage -prior to experiment- (V): " + vol.ToString() + Environment.NewLine +
                "Close Loop Control: " + checkBox1.Checked.ToString() + "; ";
                if (checkBox1.Checked)
                {
                    format = format + "Close Loop Parameters(Kp, Ki, Kd, Filter): " + "(" + textBox2.Text + ", " +
                textBox3.Text + ", " + textBox4.Text + ", " + comboBox1.SelectedItem.ToString() + ", F1:" + textBox7.Text + " Hz, F2:" +
                textBox6.Text + " Hz, " + ")" + "; " + Environment.NewLine;
                }

                //Heater Settings - active channels: set temp -act temp-duty to temp func- act dut- control type and params..
                for (int i = 0; i < heatChannels.Count(); i++)
                {
                    if (activeHeat[i] || duty[i] != 0)
                    {
                        format = format + "\n Heater Properties: Internal Heater used via MCU channel " + (i + 1).ToString() + Environment.NewLine + "Heater Duty Cycle (f:1kHz) %: " +
                            duty[i].ToString() + " for holding at T: " + setTemper[i] + Environment.NewLine;
                        if (activeHeat[i])
                        {
                            format = isSampleheat[i] ? (format + " Used as Sample Heater\n") : (format + "Used as an extra Curremt Source/Heater\n");
                            format = !feedMode[i] ? (format + "Feedback through NI-PCIe\n") : (format + "Feedback through MCU\n");
                            format = isBand[i] ? (format + "Inside Band PI Control ") : (format + "All PI Control ");
                            format = isBand[i] ? (format + "Band Interval: " + bandInter[i].ToString()) : (format);
                            format = format + $"\nProportional Gain: {proGains[i]:0.00} applied within Error larger than Sensor Deviation: {heaterSensorDev[i]:0.00}\n";
                            format = format + $"Heater to sensor thermal mass dependent time constant given: {heatTime[i] / 10.0:0.00} seconds.\n";
                            format = (slopeHeat[i] != null) ? (format + "T = %Duty " + slopeHeat[i].ToString("0.000") + " + " + cnstHeat[i].ToString("0.000")) : (format);
                        }

                    }
                }

                //Indenter type -stepmotor /actuator
                format = format + "Indentations were made with: ";
                if (motorDrive.Checked)
                {
                    format = format + "Step Motor";
                    format = format + " & Approached with only Step Motor";
                }
                else
                {
                    format = format + "Actuator";
                    if (motorApp.Checked)
                    {
                        format = format + "& Approached with only Step Motor";
                    }
                    else if (actAppOnly.Checked)
                    {
                        format = format + "& Approached with only Actuator";
                    }
                    else
                    {
                        format = format + "& Approached with Actuator-Step Motor Combination Routine";
                    }
                }

                //DAQ Settings
                //NI-PCIe
                format = format + Environment.NewLine + "\nNI-DAQmx -PCI-e6321- Channel Properties" + Environment.NewLine
    + "Rate: " + rateNumeric.Value.ToString() + "Hz; Samples per Channel: " +
    (rateNumeric.Value / 10).ToString() + "; \nInput Channels:" + Environment.NewLine;
                for (int i = 0; i < phychannel.Count; i++)
                {
                    format = format + "\nANALOG INPUT(AI)" + phychannel[i];
                    if (chtype[i] == 0)
                    {
                        format = format + ": " + "Voltage Measurement" + "Interval V(min, max): " + "(" + minval[i].ToString()
                            + ", " + maxval[i].ToString() + ")" + Environment.NewLine;
                    }
                    else if (chtype[i] == 1)
                    {
                        format = format + ": " + "Temperature Measurement" + "Interval C(min, max): " + "(" + minval[i].ToString()
        + ", " + maxval[i].ToString() + "); Thermocouple Type: " + thermocoupleTypeComboBox.Items[tctype[i]].ToString()
        + Environment.NewLine + "Cold Junction Compensation:" +
        "Source: " + cjctype[i] + "; Channel: " + cjcchannel[i] + "; Cold Junction Temperature (C): "
                            + cjcvalue[i].ToString() + Environment.NewLine;
                    }
                    else if (chtype[i] == 3)
                    {
                        format = format + ": " + "Strain Gage Measurement" + "Interval Strain(min, max): " + "(" + minval[i].ToString()
                            + ", " + maxval[i].ToString() + ")" + Environment.NewLine + "Bridge Type: " + bridgeBox.Items[gagetype[i]].ToString() +
                            " Excitation Type: " + excBox.Items[exctype[i]].ToString() + " Excitation Voltage (mV): " + gagexc[i].ToString() + Environment.NewLine
                            + "Nominal Resistance (Ohm): " + gageres[i].ToString() + " Wire Resistance (mOhm): " + gagewire[i].ToString() + Environment.NewLine +
                            "Initial Voltage (mV): " + gageinit[i].ToString() + "Poisson's Ratio: " + gagepois[i].ToString() + "Gage Factor: "
                            + gagefac[i].ToString() + Environment.NewLine;
                    }
                    else if (chtype[i] == 2)
                    {
                        format = format + ": " + "Temperature Measurement" + "Interval C(min, max): " + "(" + minval[i].ToString()
        + ", " + maxval[i].ToString() + "); 2 Wire-RTD- Type: " + rtdBox.Items[rtdtype[i]].ToString() + Environment.NewLine + " Excitation Source:"
        + rtdExcType.Items[rtdexctype[i]].ToString() + " External Excitation Current (mA):" + rtdexc[i].ToString()
        + " 0 degree Resistance (mOhm):" + rtdres[i].ToString() + Environment.NewLine;
                    }
                    if (isFeed[i])
                    {
                        format = format + "Channel was used as Feedback for heater /& load control!" + Environment.NewLine;
                    }
                }

                //Loadcell & Other Sensors
                if (controlExt.Checked)
                {
                    format = format + "\nLoadcell Measurement via HX711 24 Bit ADC 10 Hz through MCU COM" + Environment.NewLine;
                }
                else if (useGagePress.Checked)//!!FORCE COEFF & GAGE CONVERSION
                {
                    format = format + "\nLoadcell Measurement via Strain Gage Measurement through NI-PCIe" + Environment.NewLine;
                }
                else if (isMcuAdc.Checked)
                {
                    format = format + "\nLoadcell Measurement with MCU 12 Bit ADC -up to 76 MHz conversion rate-" + Environment.NewLine;
                }

                // ++EXTRA COARSE PST SENSOR && EXTERNAL USART COM PROP
                if (digiCon.Checked)
                {
                    //sensorCon
                }
                if (serialAct.Checked)
                {
                    if (saveExt.Checked)
                    {

                    }
                }

                //Experiment Parameters +++!!Control type : Voltage /Disp/Force..
                if (deney == 1)//Calibration -equal step push&pull
                {
                    format = format + Environment.NewLine + "Calibration / Equal Step Experiment Settings " + Environment.NewLine
                        + "Total Depth: " + textBox20.Text + "um; # of Steps: " + textBox21.Text + "; Time Interval between each step: "
                        + textBox22.Text + "ms; Retract to First position: " + returnCheck.Checked.ToString() + ";" + Environment.NewLine
                        + "Is Calibration: " + checkBox13.Checked.ToString() + ";";
                }
                else if (deney == 2) //Indentation Exp with Osc
                {
                    format = format + Environment.NewLine + "Indent " + button7.Text + "\nSurface Approach Control Properties" + Environment.NewLine +
                "Temperature Channel: AI " + tempChan.SelectedItem.ToString() +
                "Pressure Threshold: " + loadThres.Text + " mN; " +
                "\nApproach Speed: " + appSpeed.Text + "um/s; " + Environment.NewLine + "Load Controlled Approach Enabled: " + autoApp.Checked;
                    format = format + Environment.NewLine + "Experiment Step Settings:" + Environment.NewLine;
                    for (int i = 0; i < depth.Count; i++)
                    {
                        format = format + "\nStep " + (i + 1).ToString() + Environment.NewLine +
                            "X displacement: " + Xpos[i] + " um; Y displacement: " + Ypos[i] +
                        " um Depth: " + depth[i] + " um; " +
                       "Speed: " + speed[i] + " um/s; " +
                       "Hold Time: " + duration[i] + " s; ";
                        if (Convert.ToDouble(dTemp[i]) != 0)
                        {
                            format = format + " Temperature ramp given: " + "during:" + "with speed:";
                        }
                        if (box5[i] || box11[i])//square wave or triangle
                        {
                            format = format + "Oscillation Type: " + checkBox5.Text + "; "
                        + "Amplitude: " + amplitude[i] + " um; " +
                       "Frequency: " + frequency[i] + " Hz; " +
                       "Interval: " + interval[i] + " s; " + Environment.NewLine;
                            format = format + "Oscillation Average Position at";
                            if (oscDown.Checked)
                            {
                                format = format + " beyond one amplitude wrt last position." + Environment.NewLine;
                            }
                            else
                            {
                                format = format + " last position." + Environment.NewLine;
                            }
                        }
                        else
                        {
                            format = format + "Oscillation Type: " + "None" + "; " + Environment.NewLine;
                        }
                        if (Convert.ToDouble(dTamp[i]) != 0)
                        {
                            format = format + "Temperature Oscillation given: " + "with frequency:" + "phase lag wrt mechanical oscillation:";
                        }
                        if (retStep[i] == 1)
                        {
                            format = format + "Auto retracted after completion of step;";
                            format = format + Environment.NewLine + "Hold at " + holdAt[i] + "% of Removal for " + retHold[i] + " seconds." + Environment.NewLine;
                        }
                    }

                }
                return format;
            }
            catch
            {
                return format;
            }

        }
        int startexp = 0;
        List<string> textexp = new List<string>();
        private void button3_Click(object sender, EventArgs e)
        {
            int dirind = 0;
            textexp = new List<string>();
            int osctype = 0;
            int st = depth.Count;
            DialogResult result = DialogResult.No;
            result = MessageBox.Show("Are experiment settings ready?", "Set Experiment", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if (st > 0 && result == DialogResult.Yes)
            {
                if (autoApp.Checked && loadExt)
                {
                    Send($"401{st:D2}", "");
                }
                else if (autoApp.Checked)
                {
                    Send($"402{st:D2}", "");
                }
                else
                {
                    Send($"400{st:D2}", "");
                }
                for (int i = 0; i < depth.Count; i++)
                {
                    if (box5[i])
                    {
                        osctype = 1;
                    }
                    else if (box11[i])
                    {
                        osctype = 2;
                    }
                    else
                    {
                        osctype = 0;
                    }
                    var dtov = 0.0;
                    if (depth[i] != "")
                    {
                        if (motorDrive.Checked)
                        {
                            dirind = Math.Sign(Convert.ToDouble(depth[i]));
                            if (btn7say == 0)//Piezo Voltage  Control
                            {
                                dtov = Math.Abs(Convert.ToDouble(depth[i]) * 1000 / stepinc);
                            }
                            else if (btn7say == 1)//Force Control
                            {
                                dtov = Convert.ToDouble(depth[i]) * 1000;
                            }
                            else if (btn7say == 2)//Depth Control
                            {
                                dtov = Math.Abs(Convert.ToDouble(depth[i]) * 1000 / stepinc);
                            }
                        }
                        else
                        {
                            dirind = Math.Sign(Convert.ToDouble(depth[i]));
                            if (btn7say == 0)//Piezo Voltage  Control
                            {
                                dtov = Math.Abs((Convert.ToDouble(depth[i]) / (kp * 20)) * 1000000);
                            }
                            else if (btn7say == 1)//Force Control
                            {
                                dtov = Convert.ToDouble(depth[i]) * 1000;
                            }
                            else if (btn7say == 2)
                            {//Displacement Control
                                dtov = Math.Abs((Convert.ToDouble(depth[i]) / (kp * 20)) * 1000000);
                            }
                        }
                        if (dirind == -1)
                        {
                            dirind = 2;
                        }
                    }
                    var depthsend = string.Format("{0:000000000}|0", dtov);
                    var stov = 0.0;
                    if (speed[i] != "")
                    {
                        if (motorDrive.Checked && Convert.ToDouble(speed[i]) != 0)
                        {
                            stov = 1000 * (Math.Abs(Convert.ToDouble(depth[i])) / Convert.ToDouble(speed[i])) / dtov;
                        }
                        else
                        {
                            stov = (Convert.ToDouble(speed[i]) / (kp * 20)) * 1000000;
                        }
                    }
                    var speedsend = string.Format("{0:000000000}|0", stov);
                    var xsend = Xpos[i].Contains("-") ? (string.Format("{0:D8}|0", Convert.ToInt16(Convert.ToDouble(Xpos[i]) / xinc)))
                        : (string.Format("{0:D9}|0", Convert.ToInt16(Convert.ToDouble(Xpos[i]) / xinc)));
                    var ysend = Ypos[i].Contains("-") ? (string.Format("{0:D8}|0", Convert.ToInt16(Convert.ToDouble(Ypos[i]) / yinc)))
                        : (string.Format("{0:D9}|0", Convert.ToInt16(Convert.ToDouble(Ypos[i]) / yinc)));
                    var dtsend = string.Format("{0:00000}|0", Convert.ToDouble(dTemp[i]) * 100);
                    var dthiz = string.Format("{0:00000}|0", Convert.ToDouble(dTspeed[i]) * 1000);
                    var dtgen = string.Format("{0:00000}|0", Convert.ToDouble(dTamp[i]) * 100);
                    var dtper = string.Format("{0:00000}|0", Convert.ToDouble(dTfreq[i]) * 100);
                    var dtfaz = string.Format("{0:00000}|0", Convert.ToDouble(dTlag[i]) * 100);
                    var durasend = "000000000|0";
                    if (duration[i] != "")
                    {
                        durasend = string.Format("{0:000000000}|0", Convert.ToDouble(duration[i]) * 1000);
                    }
                    var atov = (Convert.ToDouble(amplitude[i]) / (kp * 20)) * 1000000;
                    if (motorDrive.Checked)
                    {
                        atov = Convert.ToDouble(amplitude[i]) * 1000 / stepinc;
                    }
                    var ampsend = string.Format("{0:000000000}|0", atov);
                    var freqsend = "000000000|0";
                    if (Convert.ToDouble(frequency[i]) > 0)
                    {
                        if (motorDrive.Checked)
                        {
                            freqsend = string.Format("{0:000000000}|0", 1000000 / Convert.ToDouble(frequency[i]));
                        }
                        else
                        {
                            freqsend = string.Format("{0:000000000}|0", 1000 / Convert.ToDouble(frequency[i]));
                        }
                    }
                    var intsend = string.Format("{0:000000000}|0", Convert.ToDouble(interval[i]) * 1000);
                    if (motorDrive.Checked)
                    {
                        intsend = string.Format("{0:000000000}|0", Convert.ToDouble(interval[i]) * Convert.ToDouble(frequency[i]));
                    }
                    var removeHold = Convert.ToDouble(retHold[i]) * 1000;
                    if (removeHold > 999999)
                    {
                        removeHold = 10;
                    }
                    var removalHoldSend = string.Format("{0:000000}|0", removeHold);
                    var removeHoldPercent = string.Format("{0:000}|0", holdAt[i]);
                    textexp.Insert(i, $"{i + 1:D2}" + "|" + dirind + depthsend + speedsend + durasend +
                         osctype + "|0" + ampsend + freqsend + intsend + xsend + ysend + dtsend + dthiz + dtgen + dtper + dtfaz + dTwhen[i] +
                         "|0" + retStep[i] + "|0" + removalHoldSend + removeHoldPercent);
                    if (textexp[i].Length == 149)
                    {
                        startexp = 1;
                        continue;
                    }
                    else
                    {
                        MessageBox.Show("Program Error!");
                        exppass = 0;
                        startexp = 0;
                        textexp.Clear();
                        break;
                    }
                }
                if (startexp == 1)
                {
                    exppass = 1;
                }
            }
            else
            {
                startexp = 0;
                exppass = 0;
            }
        }
        List<string> depth = new List<string>();
        List<string> speed = new List<string>();
        List<string> duration = new List<string>();
        List<string> amplitude = new List<string>();
        List<string> frequency = new List<string>();
        List<string> interval = new List<string>();
        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            labelPiezo.Text = textBox17.Text;
        }
        int directpass = 0;
        int exppass = 0;
        int calpass = 0;
        int expcounter = 0;
        int microsayac = 0;
        int rep = 0;
        int initialize = -1;
        int motorpass = 0;
        int tim1say = 0;
        int autopass = 0;
        int abqi = 0;
        int comsay = 0;
        int rhcom = 0;
        int hxcom = 0;
        int nocon = 0;
        private void textBox17_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                kp = Convert.ToDouble(textBox17.Text);
                zrange = kp * 20 * (vnull - vmin);
                textBox17.Text = labelPiezo.Text;
                Properties.Settings.Default["Piezo"] = textBox17.Text;
                Properties.Settings.Default.Save();
                textBox10.Text = TextFormat(((vnull - vol) * 20 * kp).ToString(), "Z", (vmin - (vmax - vnull) + 1) * 20 * kp, (vnull - vmin) * 20 * kp);
            }
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                textBox11.Text = TextFormat(textBox11.Text, "", -800000, 800000);
                depth.Insert(comboBox2.SelectedIndex - 1, textBox11.Text);
                if (depth.Count > comboBox2.SelectedIndex)
                {
                    depth.RemoveAt(comboBox2.SelectedIndex);
                }
                if (speed.Count < depth.Count)
                {
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (Xpos.Count < depth.Count)
                {
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                    holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                    retHold.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (Ypos.Count < depth.Count)
                {
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (duration.Count < depth.Count)
                {
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (amplitude.Count < depth.Count)
                {
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (frequency.Count < depth.Count)
                {
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (interval.Count < depth.Count)
                {
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (box5.Count < depth.Count)
                {
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                }
                if (box11.Count < depth.Count)
                {
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                }
                if (box3.Count < depth.Count)
                {
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                }
                if (depth.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }


            }
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                if (depth.Count >= comboBox2.SelectedIndex)
                {
                    textBox12.Text = TextFormat(textBox12.Text, "", 0, 100);
                    speed.Insert(comboBox2.SelectedIndex - 1, textBox12.Text);
                    if (speed.Count > comboBox2.SelectedIndex)
                    {
                        speed.RemoveAt(comboBox2.SelectedIndex);
                    }
                }
                else
                {
                    MessageBox.Show("Please enter depth first!", "Indentation Settings");
                }
            }
        }

        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                textBox13.Text = TextFormat(textBox13.Text, "", 0, 999);
                if (duration.Count == comboBox2.Items.Count - 2)
                {
                    duration.Insert(comboBox2.SelectedIndex - 1, textBox13.Text);
                }
                else
                {
                    for (int i = duration.Count; i < comboBox2.SelectedIndex - 1; i++)
                    {
                        duration.Insert(duration.Count, "0");
                    }
                    duration.Insert(comboBox2.SelectedIndex - 1, textBox13.Text);
                }
                if (duration.Count > comboBox2.SelectedIndex)
                {
                    duration.RemoveAt(comboBox2.SelectedIndex);
                }
                if (depth.Count < duration.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                    holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                    retHold.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (Xpos.Count < duration.Count)
                {
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (Ypos.Count < duration.Count)
                {
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (amplitude.Count < duration.Count)
                {
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (duration.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }

            }
        }

        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                if (checkBox5.Checked || checkBox11.Checked)
                {
                    textBox14.Text = TextFormat(textBox14.Text, "", 0, 400000);
                    if (amplitude.Count == comboBox2.Items.Count - 2)
                    {
                        amplitude.Insert(comboBox2.SelectedIndex - 1, textBox14.Text);
                    }
                    else
                    {
                        for (int i = amplitude.Count; i < comboBox2.SelectedIndex - 1; i++)
                        {
                            amplitude.Insert(amplitude.Count, "0");
                        }
                        amplitude.Insert(comboBox2.SelectedIndex - 1, textBox14.Text);
                    }
                    if (amplitude.Count > comboBox2.SelectedIndex)
                    {
                        amplitude.RemoveAt(comboBox2.SelectedIndex);
                    }
                    if (depth.Count < amplitude.Count)
                    {
                        depth.Insert(comboBox2.SelectedIndex - 1, "0");
                        speed.Insert(comboBox2.SelectedIndex - 1, "0");
                        dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                        dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                        dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                        dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                        dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                        dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                        retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                        holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                        retHold.Insert(comboBox2.SelectedIndex - 1, "0");
                    }
                    if (duration.Count < amplitude.Count)
                    {
                        duration.Insert(comboBox2.SelectedIndex - 1, "0");
                    }
                    if (Xpos.Count < amplitude.Count)
                    {
                        Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    }
                    if (Ypos.Count < amplitude.Count)
                    {
                        Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                    }
                    if (frequency.Count < amplitude.Count)
                    {
                        frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    }
                    if (interval.Count < amplitude.Count)
                    {
                        interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    }
                    if (amplitude.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                    {
                        comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                    }
                }
                else
                {
                    MessageBox.Show("Please select wave type as square or triangular", "Oscillation Settings");
                }
            }
        }

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                if (amplitude.Count >= depth.Count)
                {
                    if (amplitude.Count > comboBox2.SelectedIndex - 1)
                    {
                        if (amplitude[comboBox2.SelectedIndex - 1] != "" || amplitude[comboBox2.SelectedIndex - 1] != "0")
                        {
                            textBox15.Text = TextFormat(textBox15.Text, "", 0, 5000.0);
                            if (frequency.Count == comboBox2.Items.Count - 2)
                            {
                                frequency.Insert(comboBox2.SelectedIndex - 1, textBox15.Text);

                            }
                            else
                            {
                                for (int i = frequency.Count; i < comboBox2.SelectedIndex - 1; i++)
                                {
                                    frequency.Insert(frequency.Count, "");
                                }
                                frequency.Insert(comboBox2.SelectedIndex - 1, textBox15.Text);
                            }
                            if (frequency.Count > comboBox2.SelectedIndex)
                            {
                                frequency.RemoveAt(comboBox2.SelectedIndex);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please enter amplitude first!", "Oscillation Settings");
                        }

                    }
                    else
                    {
                        MessageBox.Show("Please enter amplitude first!", "Oscillation Settings");
                    }
                }
            }
        }

        private void textBox16_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                if (frequency.Count >= amplitude.Count)
                {
                    if (frequency.Count > comboBox2.SelectedIndex - 1)
                    {
                        if (frequency[comboBox2.SelectedIndex - 1] != "" || frequency[comboBox2.SelectedIndex - 1] != "0")
                        {
                            textBox16.Text = TextFormat(textBox16.Text, "", 0, 999);
                            if (interval.Count == comboBox2.Items.Count - 2)
                            {
                                interval.Insert(comboBox2.SelectedIndex - 1, textBox16.Text);
                            }
                            else
                            {
                                for (int i = interval.Count; i < comboBox2.SelectedIndex - 1; i++)
                                {
                                    interval.Insert(interval.Count, "0");
                                }
                                interval.Insert(comboBox2.SelectedIndex - 1, textBox16.Text);
                            }
                            if (interval.Count > comboBox2.SelectedIndex)
                            {
                                interval.RemoveAt(comboBox2.SelectedIndex);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please enter frequency first!", "Oscillation Settings");
                    }
                }
                else
                {
                    MessageBox.Show("Please enter frequency first!", "Oscillation Settings");
                }

            }
        }
        double alpha = 0.01;
        double betha = 0.9;
        private void button2_Click(object sender, EventArgs e)
        {
            if (defCount % 2 == 0)
            {
                button2.Text = "Close Define Experiment <<";
                this.Width = 769;
            }
            else
            {
                button2.Text = "Open Define Experiment >>";
                this.Width = 370 + sizeadd;
            }
            defCount++;
        }

        string volplus;
        string volmin;
        int t;
        int numosc;
        private async Task OscillationAsync(string amp, string freq, string inter, bool type, bool check)
        {

            var time = new Stopwatch();
            double w;
            int progressval = progressBar1.Value;
            if (amp != "" && freq != "" && inter != "")
            {
                double gen = Convert.ToDouble(amp);
                w = Convert.ToDouble(freq);
                double sur = Convert.ToDouble(inter);
                if (gen > 0 && w > 0 && sur > 0)
                {
                    t = Convert.ToInt32(1000 / (2 * w));
                    double lastpos = Convert.ToDouble(textBox1.Text);//bunu iyi belirlemek lazım!
                    numosc = Convert.ToInt32(Convert.ToDouble(w * sur));
                    if (oscDown.Checked)
                    {
                        volmin = TextFormat((lastpos).ToString(), "", -1, 7.5);
                        volplus = TextFormat(((lastpos * 20 - gen / kp) / 20).ToString(), "", -1, 7.5);
                    }
                    else
                    {
                        volmin = TextFormat(((lastpos * 20 + gen / kp) / 20).ToString(), "", -1, 7.5);
                        volplus = TextFormat(((lastpos * 20 - gen / kp) / 20).ToString(), "", -1, 7.5);
                    }
                    if (type)
                    {
                        time.Start();
                        for (int i = 0; i < numosc; i++)
                        {
                            if (emCounter % 2 != 0)
                            {
                                if (checkBox15.Checked)
                                {
                                    Send(volplus, "W");
                                }
                                textBox1.Text = volplus;
                                //textBox10.Update();
                                //textBox1.Update();
                                await Task.Delay(t);
                                timeshow.Text = time.ElapsedMilliseconds.ToString();
                                time.Restart();
                                if (checkBox15.Checked)
                                {
                                    Send(volmin, "W");
                                }
                                textBox1.Text = volmin;
                                //textBox1.Update();
                                //textBox10.Update();
                                await Task.Delay(t);
                                timeshow.Text = time.ElapsedMilliseconds.ToString();
                            }
                            else
                            {
                                time.Reset();
                                break;
                            }
                        }
                        time.Reset();
                    }
                    else if (check)
                    {
                        string step = (5 / w).ToString();
                        returnCheck.Checked = true;
                        time.Start();
                        for (int i = 0; i < numosc; i++)
                        {
                            await CalibrationAsync(amp, step, "");
                            timeshow.Text = time.ElapsedMilliseconds.ToString();
                            time.Restart();
                        }
                        time.Stop();
                    }

                    string lp = TextFormat(Convert.ToString(lastpos), "Vapp", -1, 7.5);
                    Send(lp, "W");
                }

            }
        }
        string totaldepth;
        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                textBox20.Text = labelCalDepth.Text;
                totaldepth = textBox20.Text;
                if (motorDrive.Checked)
                {
                    textBox21.Text = Convert.ToString((int)(Convert.ToDouble(textBox20.Text) * 1000 / stepinc));
                    labelSteps.Text = textBox21.Text;
                    steps = textBox21.Text;
                }
            }
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            labelCalDepth.Text = TextFormat(textBox20.Text, "", 0, 8000000);
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            textBox21.Text = textBox21.Text;
        }
        string steps;
        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                labelSteps.Text = textBox21.Text;
                steps = textBox21.Text;
            }
        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            labelCalTime.Text = TextFormat(textBox22.Text, "", 0, 100);
        }
        string caltime;
        private void textBox22_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                textBox22.Text = labelCalTime.Text;
                caltime = textBox22.Text;
            }
        }

        private async Task IndentAsync(int i)
        {
            double a = 0;
            double depthform = 0;
            double voltage = 0;
            double zind = 0;
            if (depth[i] != "")
            {
                depthform = Convert.ToDouble(depth[i]);
            }
            if (depthform != 0)
            {
                voltage = (Convert.ToDouble(textBox1.Text) * 20 - (depthform / kp)) / 20;
                if (speed[i] == "" || speed[i] == "0.000000")
                {
                    if (depthform != 0 && emCounter % 2 != 0)
                    {
                        Properties.Settings.Default["Vapp"] = TextFormat(voltage.ToString(), "Vapp", -1, 7.5);
                        Properties.Settings.Default.Save();
                        textBox1.Text = TextFormat(voltage.ToString(), "Vapp", -1, 7.5);
                        textBox1.Update();
                        textBox10.Update();
                        if (checkBox15.Checked)
                        {
                            Send(TextFormat(voltage.ToString(), "Vapp", -1, 7.5), "W");
                        }
                        await Task.Delay(55);
                    }
                }
                else
                {
                    double stepno = depthform / (Convert.ToDouble(speed[i]) * 0.1);
                    await CalibrationAsync(depth[i], stepno.ToString(), "");
                }
            }
            if (duration[i] != "")
                a = 1000 * Convert.ToDouble(duration[i]);
            if (a != 0)
            {
                await Task.Delay(Convert.ToInt32(a));
            }


        }
        int deney = -1;
        public static bool calib_but = false;
        string header = "";
        int convertibl = 0;
        private async void button8_ClickAsync(object sender, EventArgs e)
        {
            timer3.Stop();
            double tb20 = 0;
            int tb22 = 50;
            stopCal.Enabled = true;
            calib_but = true;
            groupBox8.Enabled = false;
            tim2i = 0;
            graphcount++;
            try
            {
                tb20 = Convert.ToDouble(textBox20.Text);
                convertibl = 1;
            }
            catch (Exception ex)
            {
                connection = ex.Message;
                convertibl = 0;
            }
            if (((!motorDrive.Checked && (zrange - zpos >= tb20 && !firstDir.Checked) || (zpos >= tb20 && firstDir.Checked)) ||
                (motorDrive.Checked && (motmax - motpos >= tb20 * 1000 && !firstDir.Checked) || (motpos >= tb20 * 1000 && firstDir.Checked)))
                && textBox21.Text != "" && convertibl == 1)
            {
                string zaman = "\\" + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss");
                directorset(zaman);
                deney = 1;
                emCounter = 1;
                loadd.RemoveRange(0, loadd.Count);
                vapp.RemoveRange(0, vapp.Count);
                if (cal)
                {
                    volcal2.RemoveRange(0, volcal2.Count);
                    volvs2.RemoveRange(0, volvs2.Count);
                }
                else
                {
                    volcal1.RemoveRange(0, volcal1.Count);
                    volvs2.RemoveRange(0, volvs2.Count);
                }
                if (rateNumeric.Value <= 10000)
                {
                    timer4.Interval = Convert.ToUInt16(10000 / rateNumeric.Value);
                    timer5.Interval = Convert.ToUInt16(10000 / rateNumeric.Value);
                    timer2.Interval = 100;
                }
                else
                {
                    timer4.Interval = 100;
                    timer5.Interval = 100;
                    timer2.Interval = 100;
                }
                chart1.Enabled = false;
                if (!checkBox15.Checked)
                {
                    var totdep = (tb20 / (kp * 20)) * 1000000;
                    if (motorDrive.Checked)
                    {
                        totdep = tb20 * 1000 / stepinc;
                    }
                    var depthsend = string.Format("{0:000000000}|0", totdep);
                    var steps = string.Format("{0:000000000}|0", Convert.ToInt16(textBox21.Text));
                    var intervalcal = 0.0;
                    var intersend = "";
                    var calhold = string.Format("{0:000000}|0", calHoldDur);
                    DialogResult result = DialogResult.None;
                    try
                    {
                        intervalcal = Convert.ToDouble(textBox22.Text) * 1000;
                        if (intervalcal >= 0.0002 && motorDrive.Checked && emCounter % 2 != 0)
                        {
                            intersend = string.Format("{0:000000000}|0", intervalcal);
                            tb22 = (int)intervalcal;
                            texts = depthsend + steps + intersend + calhold;
                            calpass = 1;
                        }
                        else if (motorDrive.Checked && emCounter % 2 != 0)
                        {
                            MessageBox.Show("Time interval is set to 200us because entered value lower than mechanical limit!" +
    " \n Do you want to continue?", "UC45 Secure Communication Limits", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            intervalcal = 0.2;
                            intersend = string.Format("{0:000000000}|0", intervalcal);
                            tb22 = (int)intervalcal;
                            texts = depthsend + steps + intersend + calhold;
                            calpass = 1;
                        }
                        else if (!motorDrive.Checked && intervalcal >= 0.05 && emCounter % 2 != 0)
                        {
                            intersend = string.Format("{0:000000000}|0", intervalcal);
                            tb22 = (int)intervalcal;
                            texts = depthsend + steps + intersend + calhold;
                            calpass = 1;
                        }
                        else if (emCounter % 2 != 0)
                        {
                            result = MessageBox.Show("Time interval is set to 50ms because entered value lower than communication limit!" +
    " \n Do you want to continue?", "UC45 Secure Communication Limits", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        }
                    }
                    catch (Exception ex)
                    {
                        connection = ex.Message;
                        result = MessageBox.Show("Time interval is set to 50ms because communication limit of UC45! \n Do you want to continue?",
                            "UC45 Secure Communication Limits", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    }
                    if (result == DialogResult.Yes)
                    {
                        intersend = "000000000|0";
                        texts = depthsend + steps + intersend + calhold;
                        calpass = 1;
                    }
                    if (returnCheck.Checked)
                    {
                        if (firstDir.Checked)
                        {
                            Send("41011", "");
                        }
                        else
                        {
                            Send("41010", "");
                        }
                    }
                    else
                    {
                        if (firstDir.Checked)
                        {
                            Send("41001", "");
                        }
                        else
                        {
                            Send("41000", "");
                        }
                        //41-00-returntype-dirtype
                    }
                }
                appret = 0;
                motpos0 = motpos;
                zpos0 = zpos;
                if (!checkBox15.Checked)
                {
                    while (appret == 0 && emCounter % 2 != 0)
                    {
                        await Task.Delay(50);
                    }
                    if (btn9 == 1 || loadExt)
                    {
                        Tdms_Saver("Calibration");
                        await Task.Delay(timer2.Interval);
                    }
                    timer2.Start();
                    if (motorDrive.Checked && emCounter % 2 != 0)
                    {
                        Send("DOIT!100", "");
                    }
                    else if (emCounter % 2 != 0)
                    {
                        Send("DOIT!_00", "");
                    }
                    expfin = 0;
                    File.WriteAllText(fileopen + zaman + "_ExperimentLog.txt", Experimentlog(zaman));
                    explogPath = fileopen + zaman + "_ExperimentLog.txt";
                    fileopen = pathtemp + "\\Calibration";
                    while (expfin == 0 && emCounter % 2 != 0)
                    {
                        await Task.Delay(10);
                        if (motorDrive.Checked)
                        {
                            motorPos.Text = Convert.ToString(motpos / 1000000);
                            //verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                        }
                        else
                        {
                            textBox1.Text = Convert.ToString(vol);
                        }
                    }
                    Properties.Settings.Default["motpos"] = motorPos.Text;
                    Properties.Settings.Default["Vapp"] = textBox1.Text;
                    Properties.Settings.Default.Save();
                }
                else if (checkBox15.Checked && emCounter % 2 != 0)
                {
                    await CalibrationAsync(textBox20.Text, textBox21.Text, textBox22.Text);
                }
                if (emCounter % 2 != 0)
                {
                    Send("CLFIN", "");
                }
                if (btn9 == 1)
                {
                    runningTask = null;
                    myTask.Dispose();
                }
                timer2.Stop();
                timer5.Stop();
                timer4i = 0;
                timer5i = 0;
                chart1.Enabled = true;
                emCounter = 1;
                deney = -1;
                tb1.RemoveRange(0, tb1.Count);
                tb10.RemoveRange(0, tb10.Count);
            }
            else
            {
                MessageBox.Show("Please enter depth and step value appropriately!",
                    "Experiment Range Controller", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            groupBox8.Enabled = true;
            stopCal.Enabled = false;
            calib_but = false;
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked)
            {
                Send("1", "T");
                checkBox10.Text = "Internal/ Servo On!";
                Properties.Settings.Default["Mode"] = "1";
                Properties.Settings.Default.Save();
                VappNumeric.Enabled = true;
                if (zCount % 2 == 0)
                {
                    textBox1.Enabled = true;
                }
                else
                {
                    textBox10.Enabled = true;

                }
                button10.Enabled = true;
                button6.Enabled = true;
            }
            else
            {
                VappNumeric.Enabled = false;
                textBox1.Enabled = false;
                textBox10.Enabled = false;
                button10.Enabled = false;
                button6.Enabled = false;
                Send("0", "T");
                checkBox10.Text = "External/ Servo On|Off";
                Properties.Settings.Default["Mode"] = "0";
                Properties.Settings.Default.Save();
            }
        }
        bool cb5 = false;
        List<bool> box5 = new List<bool>();
        List<bool> box3 = new List<bool>();
        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (!combochange)
            {
                if (checkBox5.Checked)
                {
                    checkBox11.Checked = false;
                    checkBox3.Checked = false;
                    checkBox3.Enabled = false;
                    box5.Insert(comboBox2.SelectedIndex - 1, true);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, false);
                }
                else
                {
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    checkBox3.Checked = true;
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    checkBox3.Enabled = true;
                    textBox14.Enabled = false;
                    textBox15.Enabled = false;
                    textBox16.Enabled = false;
                    checkBox11.Enabled = false;
                    checkBox5.Enabled = false;
                }
                if (box5.Count > comboBox2.SelectedIndex)
                {
                    box3.RemoveAt(comboBox2.SelectedIndex);
                    box5.RemoveAt(comboBox2.SelectedIndex);
                }
            }
        }
        bool cb11 = false;
        List<bool> box11 = new List<bool>();
        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (!combochange)
            {
                if (checkBox11.Checked)
                {
                    checkBox5.Checked = false;
                    checkBox3.Checked = false;
                    checkBox3.Enabled = false;
                    box11.Insert(comboBox2.SelectedIndex - 1, true);
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, false);
                }
                else
                {
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    checkBox3.Checked = true;
                    checkBox3.Enabled = true;
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    textBox14.Enabled = false;
                    textBox15.Enabled = false;
                    textBox16.Enabled = false;
                    checkBox11.Enabled = false;
                    checkBox5.Enabled = false;
                }
                if (box11.Count > comboBox2.SelectedIndex)
                {
                    box3.RemoveAt(comboBox2.SelectedIndex);
                    box11.RemoveAt(comboBox2.SelectedIndex);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int s = comboBox2.SelectedIndex;
            if (comboBox2.SelectedIndex > 0)
            {
                if (depth.Count >= s)
                {
                    depth.RemoveAt(s - 1);
                    speed.RemoveAt(s - 1);
                    duration.RemoveAt(s - 1);
                    amplitude.RemoveAt(s - 1);
                    frequency.RemoveAt(s - 1);
                    interval.RemoveAt(s - 1);
                    box5.RemoveAt(s - 1);
                    box11.RemoveAt(s - 1);
                    box3.RemoveAt(s - 1);
                    Xpos.RemoveAt(s - 1);
                    Ypos.RemoveAt(s - 1);
                    dTspeed.RemoveAt(s - 1);
                    dTamp.RemoveAt(s - 1);
                    dTfreq.RemoveAt(s - 1);
                    dTemp.RemoveAt(s - 1);
                    dTwhen.RemoveAt(s - 1);
                    dTlag.RemoveAt(s - 1);
                    retStep.RemoveAt(s - 1);
                    holdAt.RemoveAt(s - 1);
                    retHold.RemoveAt(s - 1);
                }
                else if (amplitude.Count >= s)
                {
                    amplitude.RemoveAt(s - 1);
                    frequency.RemoveAt(s - 1);
                    interval.RemoveAt(s - 1);
                    box5.RemoveAt(s - 1);
                    box11.RemoveAt(s - 1);
                    box3.RemoveAt(s - 1);
                    Xpos.RemoveAt(s - 1);
                    Ypos.RemoveAt(s - 1);
                    if (duration.Count >= s)
                    {
                        duration.RemoveAt(s - 1);
                    }
                }
                else if (duration.Count >= s)
                {
                    duration.RemoveAt(s - 1);
                    Xpos.RemoveAt(s - 1);
                    Ypos.RemoveAt(s - 1);
                }
                if (comboBox2.Items.Count != 2)
                    comboBox2.Items.RemoveAt(s);
                for (int i = 1; i < comboBox2.Items.Count; i++)
                {
                    comboBox2.Items[i] = i;
                }
                comboBox2.SelectedIndex = s - 1;
            }
        }
        string initialpos = "";
        string abq = "";
        int approaching = 0;
        string fileopentdms = Application.StartupPath + @"\tdms";
        string fileopen = "";
        bool forward = true;
        double zexp = 0;
        public static bool fin = false;
        double zpos0 = 0;
        int expfin = 0;
        int appret = 0;
        string explogPath = "";
        private async void ExecuteExp_ClickAsync(object sender, EventArgs e)
        {
            if (deney == 3)
            {
                showData.PerformClick();
            }
            timer3.Stop();
            var time = new Stopwatch();
            groupBox5.Enabled = false;
            stopExp.Enabled = true;
            string zaman = "\\" + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss");
            directorset(zaman);
            deney = 2;
            fin = false;
            emCounter = 1;
            approaching = 0;
            progressBar1.Value = 0;
            vapp.RemoveRange(0, vapp.Count);
            loadd.RemoveRange(0, loadd.Count);
            volcal1.RemoveRange(0, volcal1.Count);
            volvs1.RemoveRange(0, volvs1.Count);
            chart1.Enabled = false;
            tim2i = 0;
            graphcount++;
            File.WriteAllText(fileopen + zaman + "_ExperimentLog.txt", Experimentlog(zaman));
            explogPath = fileopen + zaman + "_ExperimentLog.txt";
            if (rateNumeric.Value <= 10000)
            {
                timer4.Interval = Convert.ToUInt16(10000 / rateNumeric.Value);
                timer5.Interval = Convert.ToUInt16(10000 / rateNumeric.Value);
                timer2.Interval = 100;
            }
            else
            {
                timer4.Interval = 100;
                timer5.Interval = 100;
                timer2.Interval = 100;
            }
            if (!checkBox15.Checked)//stm32 ile kullanmak için
            {
                string triv = "DOIT";
                if (motorApp.Checked)//approach with motor
                {
                    triv = triv + "1";
                }
                else
                {
                    triv = triv + "0";
                }
                if (motorDrive.Checked)//experiment with motor
                {
                    triv = triv + "1";
                }
                else
                {
                    triv = triv + "0";
                }
                if (actAppOnly.Checked)
                {
                    triv = triv + "1";
                }
                else
                {
                    triv = triv + "0";
                }
                if (oscDown.Checked)//oscillation only downward
                {
                    triv = triv + "1";
                }
                else
                {
                    triv = triv + "0";
                }
                double speedsend = 0;
                if (motorApp.Checked)
                {
                    speedsend = 1000000 * stepinc / (speedApp * 2);
                    if (speedsend > 999999)
                    {
                        speedsend = 500000;
                    }
                }
                else
                {
                    speedsend = 50 * speedApp / (5 * kp); // SPEED SEND IS INCREMENT GIVEN AS 200MS STEPS; so divided by 5 to send exact number.
                    if (speedsend > 999999)
                    {
                        speedsend = 500000;
                    }
                }
                triv = triv + string.Format("{0:000000}|0{1:000000}", loadAppThres, speedsend);
                if (triv.Length == 22)
                {
                    Send(triv, "");
                }
                else
                {
                    DialogResult result = MessageBox.Show("Program Error!" + Environment.NewLine + "Do yo want to actuator controlled, no autoretract, normal oscillation experiment?",
                        "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                    if (result == DialogResult.Yes)
                    {
                        Send($"DOIT0000{Convert.ToInt32(loadAppThres):D6}|0{ Convert.ToInt32(speedsend):D6}", "");
                        File.AppendAllText(explogPath, "Experiment Error Log: " + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss") + Environment.NewLine +
                            "Program Error Occurred! Experiment was continued with actuator controlled-no auto retract- nomal oscillation mode with given steps");
                    }
                    else
                    {
                        Send($"Z0000000{Convert.ToInt32(loadAppThres):D6}|0{ Convert.ToInt32(speedsend):D6}", "");
                        File.AppendAllText(explogPath, "Experiment Error Log: " + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss") + Environment.NewLine +
                          "Program Error Occurred! Experiment was aborted.");
                        emCounter++;
                    }
                }
            }
            motpos0 = motpos;
            zpos0 = zpos;
            int syc = 0;
            maxtemp = 0;
            maxload = 0;
            while (appret == 0 && emCounter % 2 != 0)
            {
                await Task.Delay(10);
                syc++;
                if (comread.Contains("Stop"))
                {
                    File.AppendAllText(explogPath, "Experiment Error Log: " + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss") + Environment.NewLine +
  "Communication Error Occurred! Experiment was stopped through MCU.");
                    break;
                }
            }
            ExperimentRangeController(depth, amplitude, retStep);
            initialpos = textBox1.Text;
            if (doit)
            {
                await Task.Delay(20);
                Send("START", "");
                time.Start();
                for (int i = 0; i < depth.Count; i++)
                {
                    if (emCounter % 2 != 0)
                    {
                        apfin = 0;
                        if ((autoApp.Checked && i == 0) || (retState == 1 && autoApp.Checked))
                        {
                            approaching = 1;
                            if ((btn9 == 1 || loadExt))
                            {
                                Tdms_Saver("Step " + (i + 1) + " Approach");
                                await Task.Delay(timer2.Interval);
                                timer2.Start();
                            }
                            fileopen = pathtemp + "\\Step" + (i + 1) + "_Approach";
                            if (checkBox15.Checked)
                            {
                                await AutoConAsync(pressure, tempthres, gradthres, forward);
                            }
                            else
                            {
                                while (apfin == 0 && emCounter % 2 != 0)
                                {
                                    await Task.Delay(10);
                                    try
                                    {
                                        motorPos.Text = Convert.ToString(motpos / 1000000);
                                        textBox1.Text = Convert.ToString(vol);
                                    }
                                    catch (Exception ex)
                                    {
                                        label11.Text = ex.Message;
                                        connection = comread;
                                    }
                                }
                                if (runningTask != null)
                                {
                                    timer2.Stop();
                                    await Task.Delay(timer2.Interval);
                                    runningTask = null;
                                    myTask.Dispose();
                                }
                                if (emCounter % 2 != 0)
                                {
                                    Send("APFIN", "");
                                }
                            }
                            timer3.Stop();
                            approaching = 0;
                            retState = 0;
                        }
                        if (((depth[i] != "0.000000" && depth[i] != "0") || (duration[i] != "0.000000" && duration[i] != "0")) && (btn9 == 1 || loadExt))
                        {
                            Tdms_Saver("Step " + (i + 1) + " Indentation");
                            await Task.Delay(timer2.Interval);
                            fileopen = pathtemp + "\\Step" + (i + 1) + "_Indentation";
                            timer2.Start();
                        }
                        infin = 0;
                        if (checkBox15.Checked)
                        {
                            await Task.Delay(timer2.Interval);
                            await IndentAsync(i);
                        }
                        else
                        {
                            while (infin == 0 && emCounter % 2 != 0)
                            {
                                await Task.Delay(10);
                                try
                                {
                                    if (motorDrive.Checked)
                                    {
                                        motorPos.Text = Convert.ToString(motpos / 1000000);
                                        //verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                                    }
                                    else if (!motorDrive.Checked)
                                    {
                                        textBox1.Text = Convert.ToString(vol);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    label11.Text = ex.Message;
                                    connection = comread;
                                }
                            }
                            if (runningTask != null)
                            {
                                timer2.Stop();
                                await Task.Delay(timer2.Interval);
                                runningTask = null;
                                myTask.Dispose();
                            }
                            if (emCounter % 2 != 0)
                            {
                                Send("INFIN", "");
                            }
                            Properties.Settings.Default["motpos"] = motorPos.Text;
                            Properties.Settings.Default["Vapp"] = textBox1.Text;
                            Properties.Settings.Default.Save();
                        }
                        if ((btn9 == 1 || loadExt) && amplitude[i] != "0.000000" && amplitude[i] != "0" && emCounter % 2 != 0)
                        {
                            Tdms_Saver("Step " + (i + 1) + " Oscillation");
                            await Task.Delay(timer2.Interval);
                            timer2.Start();
                            fileopen = pathtemp + "\\Step" + (i + 1) + "_Oscillation";
                        }
                        osfin = 0;
                        if (checkBox15.Checked)
                        {
                            await OscillationAsync(amplitude[i], frequency[i], interval[i], box5[i], box11[i]);
                        }
                        else
                        {
                            while (osfin == 0 && emCounter % 2 != 0)
                            {
                                await Task.Delay(10);
                                try
                                {
                                    if (motorDrive.Checked)
                                    {
                                        motorPos.Text = Convert.ToString(motpos / 1000000);
                                        //verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                                    }
                                    else if (!motorDrive.Checked)
                                    {
                                        textBox1.Text = Convert.ToString(vol);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    label11.Text = ex.Message;
                                    connection = comread;
                                }
                            }
                            if (runningTask != null)
                            {
                                timer2.Stop();
                                await Task.Delay(timer2.Interval);
                                runningTask = null;
                                myTask.Dispose();
                            }
                            if (emCounter % 2 != 0)
                            {
                                Send("OSFIN", "");
                            }
                            Properties.Settings.Default["Vapp"] = textBox1.Text;
                            Properties.Settings.Default["motpos"] = motorPos.Text;
                            Properties.Settings.Default.Save();
                        }
                        if (Xpos.Count > i + 1)
                        {
                            if ((Convert.ToDouble(Xpos[i + 1]) != 0 || Convert.ToDouble(Ypos[i + 1]) != 0))
                            {
                                retStep[i] = 1;
                            }
                        }
                        if (retStep[i] == 1)
                        {
                            retState = 1;
                            if ((btn9 == 1 || loadExt))
                            {
                                Tdms_Saver("Step " + (i + 1) + " Retract");
                                await Task.Delay(timer2.Interval);
                                timer2.Start();
                            }
                            fileopen = pathtemp + "\\Step" + (i + 1) + "_Retract";
                            refin = 0;
                            if (checkBox15.Checked)
                            {
                                await AutoConAsync(pressure, tempthres, tempthres, !forward);
                            }
                            else
                            {
                                while (refin == 0 && emCounter % 2 != 0)
                                {
                                    await Task.Delay(10);
                                    try
                                    {
                                        motorPos.Text = Convert.ToString(motpos / 1000000);
                                        textBox1.Text = Convert.ToString(vol);
                                    }
                                    catch (Exception ex)
                                    {
                                        label11.Text = ex.Message;
                                        connection = comread;
                                    }
                                }
                                if (runningTask != null)
                                {
                                    timer2.Stop();
                                    await Task.Delay(timer2.Interval);
                                    runningTask = null;
                                    myTask.Dispose();
                                }
                                if (emCounter % 2 != 0)
                                {
                                    Send("REFIN", "");

                                }
                            }
                        }
                        Properties.Settings.Default["Vapp"] = textBox1.Text;
                        Properties.Settings.Default["motpos"] = motorPos.Text;
                        Properties.Settings.Default.Save();
                    }
                    else
                    {
                        File.AppendAllText(explogPath, "Experiment Error Log: " + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss") + Environment.NewLine +
$"Experiment was stopped at Step {i:D2}.");
                        break;
                    }
                }
                if (emCounter % 2 != 0)
                {
                    while (expfin == 0 && emCounter % 2 != 0)
                    {
                        await Task.Delay(20);
                    }
                    File.AppendAllText(explogPath, "Experiment Error Log: " + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss") + Environment.NewLine +
"Experiment was finished successfully.");
                }
            }
            else
            {
                Send("Z0000", "");
                if (emCounter % 2 != 0)
                {
                    MessageBox.Show("Experiment Range Exceeds the Indenter Range!", "Experiment Range Controller",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                    File.AppendAllText(explogPath, "Experiment Error Log: " + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss") + Environment.NewLine +
"Experiment Range Exceeded the Indenter Range! Experiment was stopped.");
                }
            }
            if (btn9 == 1)
            {
                runningTask = null;
                myTask.Dispose();
            }
            await Task.Delay(timer2.Interval);
            timer2.Stop();
            timer5.Stop();
            timer4.Stop();
            timer4i = 0;
            timer5i = 0;
            progressBar1.Value = 100;
            depth.RemoveRange(0, depth.Count);
            speed.RemoveRange(0, speed.Count);
            duration.RemoveRange(0, duration.Count);
            amplitude.RemoveRange(0, amplitude.Count);
            frequency.RemoveRange(0, frequency.Count);
            interval.RemoveRange(0, interval.Count);
            box5.RemoveRange(0, box5.Count);
            box11.RemoveRange(0, box11.Count);
            box3.RemoveRange(0, box3.Count);
            Xpos.RemoveRange(0, Xpos.Count);
            Ypos.RemoveRange(0, Ypos.Count);
            dTemp.RemoveRange(0, dTemp.Count);
            dTamp.RemoveRange(0, dTamp.Count);
            dTwhen.RemoveRange(0, dTwhen.Count);
            dTfreq.RemoveRange(0, dTfreq.Count);
            dTlag.RemoveRange(0, dTlag.Count);
            dTspeed.RemoveRange(0, dTspeed.Count);
            retStep.RemoveRange(0, retStep.Count);
            retState = 0;
            for (int j = comboBox2.Items.Count - 1; j > 0; j--)
            {
                comboBox2.Items.RemoveAt(j);
            }
            if (comboBox2.Items.Count < 2)
            {
                comboBox2.Items.Add(1);
            }
            comboBox2.SelectedIndex = 1;
            Task.WaitAll();
            fin = true;
            tb1.RemoveRange(0, tb1.Count);
            tb10.RemoveRange(0, tb10.Count);
            emCounter = 1;
            deney = -1;
            groupBox5.Enabled = true;
            stopExp.Enabled = false;
            executeExp.Enabled = false;
            chart1.Enabled = true;
            progressBar1.Value = 0;
        }
        double holdposition = 0;
        int expresscon = -1;
        int holdtime = 1000;
        string receive = "";
        string receives = "";
        int pass = 0;
        int repsay = 0;
        public string comread = "";
        string lastwords = "";
        int heatlow = 0;
        int retState = 0;
        int refin = 0;
        int apfin = 0;
        int osfin = 0;
        int infin = 0;
        bool[] xylimits = {true,true,true,true};  // {X_Minus,X_Positive,Y_Minus,Y_Positive} holds whether xy positioner is in between
        //movement ranges wrt MCU ext-int (endstop) message || for now it is accepted that xy positioner starts inside the ranges.
        private void serialPort2_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            comread = "";
            int nbyt = 0;
            char[] buffer = null;
            //serialPort2.BytesToRead
            if (serialPort2.IsOpen)
            {
                try
                {
                    nbyt = serialPort2.BytesToRead;
                    buffer = new char[nbyt];
                    serialPort2.Read(buffer, 0, nbyt);
                    //comread = serialPort2.ReadLine();
                    if (buffer[0] == 0x24)
                    {
                        for (int i = 1; i < nbyt; i++)
                        {
                            comread = comread + buffer[i];
                        }
                    }
                    nbyt -= 1;
                }
                catch (Exception ex)
                { connection = ex.Message; }
                if (comread.Count() != 0)
                {
                    if (comread == "OK")
                    {
                        try
                        {
                            serialPort2.Write("OK");
                            isConnected = true;
                            tim1say = 0;
                            timer6.Start();
                        }
                        catch (Exception ex)
                        {
                            connection = ex.Message;
                        }
                    }
                    else if (comread.Contains("AGAIN"))
                    {
                        try
                        {
                            serialPort2.Write("I");
                            isConnected = true;
                            tim1say = 0;
                            timer6.Start();
                        }
                        catch (Exception ex)
                        {
                            connection = ex.Message;
                        }
                    }
                    else if (jsBox.Checked && comread.Contains("MSP"))
                    {
                        spedmode = Convert.ToInt16(comread[3].ToString());
                        jspass = 1;
                        timer6.Start();
                    }
                    else if (comread.Contains("Finished") || comread.Contains("Retfin") || comread.Contains("Complete"))
                    {
                        if (comread.Contains("Osc"))
                        {
                            osfin = 1;
                        }
                        else if (comread.Contains("Ind"))
                        {
                            infin = 1;
                        }
                        else if (comread.Contains("Ret"))
                        {
                            refin = 1;
                        }
                        else if (comread.Contains("App"))
                        {
                            apfin = 1;
                        }
                        else if (comread.Contains("Exp")|| comread.Contains("Cal"))
                        {
                            expfin = 1;
                        }
                    }
                    else if (comread.Contains("APP") || comread.Contains("Exe") || comread.Contains("?"))
                    {
                        appret = 1;
                    }
                    else if(comread.Contains("Restart UC45"))
                    {
                        MessageBox.Show("UC45 Connection Lost! Please Reset CEDRAT's Electronics after take the tip to a secure position.");
                    }
                    else if(comread.Contains("Load Problem"))
                    {
                        MessageBox.Show("Load Measurement Problem! Please stop the experiment and check the loadcell data.");
                    }
                    else if (comread.Contains("Not_App") || comread.Contains("OutRan"))
                    {
                        //Not_Approached ; OutRange states..
                        groupBox5.Enabled = true;
                        executeExp.Enabled = false;
                        emCounter = 0;
                        texts = "";
                        textexp.Clear();
                    }
                    else if (comread.Contains("OutX"))
                    {
                        timer6.Start();
                        pass = 3;
                        timer6.Start();
                        if (comread.Contains("M"))
                        {
                            xylimits[0] = false;
                        }
                        else if (comread.Contains("P"))
                        {
                            xylimits[1] = false;
                        }
                    }
                    else if (comread.Contains("OutY"))
                    {
                        timer6.Start();
                        pass = 3;
                        if (comread.Contains("M"))
                        {
                            xylimits[2] = false;
                        }
                        else if (comread.Contains("P"))
                        {
                            xylimits[3] = false;
                        }
                    }
                    else if (comread.Contains("InX"))
                    {
                        timer6.Start();
                        pass = 3;
                        if (comread.Contains("M"))
                        {
                            xylimits[0] = true;
                        }
                        else if (comread.Contains("P"))
                        {
                            xylimits[1] = true;
                        }
                    }
                    else if (comread.Contains("InY"))
                    {
                        timer6.Start();
                        pass = 3;
                        if (comread.Contains("M"))
                        {
                            xylimits[2] = true;
                        }
                        else if (comread.Contains("P"))
                        {
                            xylimits[3] = true;
                        }
                    }
                    else if (comread.Contains("UPMOT"))
                    {
                        receive = "UPMOT";
                        if (autopass == 1)
                        {
                            pass = 1;
                        }
                        if (approaching == 2)
                        {
                            pass = 1;
                            timer6.Start();
                        }
                    }
                    else if (comread.Contains("CLOSe"))
                    {
                        receive = comread;
                        if (approaching == 2)
                        {
                            pass = 1;
                            timer6.Start();
                        }
                    }
                    else if (comread.Contains("TSET") && heat_com == 1)
                    {
                        Send(heatSender, "");
                    }
                    else if (comread.Contains("THRESHOLD"))
                    {
                        Send(hxsend, "");
                    }
                    else if (comread.Contains("Heater"))
                    {
                        Send("HEFIN", "");
                        heat_com = 0;
                        heatSender = "";
                        if (comread == "HalfDuty")
                        {
                            heatlow = 1;
                        }
                        if (comread != "DutyLowest")
                        {
                            heatlow = 2;
                        }
                    }
                    else if (comread[nbyt - 1].ToString() == "E")
                    {
                        try
                        {
                            abq = "";
                            comsay = 0;
                            if (comread.Contains("LM"))
                            {
                                for (int i = 0; i < nbyt; i++)
                                {
                                    comsay++;
                                    if (comread[i].ToString() == "L")
                                    {
                                        break;
                                    }
                                    abq += comread[i].ToString();
                                }
                                loadd.Add(Convert.ToDouble(abq));
                                if (maxload < loadd.Last())
                                {
                                    maxload = loadd.Last();
                                }
                                if (deney != -1 && deney != 0)
                                {
                                    File.AppendAllText(fileopen + "_Load.txt", loadd.Last().ToString() + Environment.NewLine);
                                }
                                comsay++;
                            }
                            abq = "";
                            for (int i = comsay; i < nbyt; i++)
                            {

                                if (comread[i].ToString() == "E")
                                {
                                    break;
                                }
                                abq += comread[i].ToString();
                            }
                            vol = Convert.ToDouble(abq);
                            if (nbyt >= 7)
                            {
                                receive = Convert.ToString(vol);
                                if (deney != -1 && deney != 0)
                                {
                                    File.AppendAllText(fileopen + "_Actuator_Voltage.txt", receive + Environment.NewLine);
                                }
                            }
                            comsay = 0;
                            repsay = 0;
                        }
                        catch (Exception ex)
                        {
                            connection = ex.Message;
                        }
                    }
                    else if (comread.Contains("PM"))
                    {
                        try
                        {
                            abq = "";
                            comsay = 0;
                            for (int i = 0; i < nbyt; i++)
                            {
                                comsay++;
                                if (comread[i].ToString() == "P")
                                {
                                    break;
                                }
                                abq += comread[i].ToString();
                            }
                            try
                            {
                                abqi = Convert.ToInt16(abq);
                            }
                            catch
                            {
                                abqi = 0;
                            }
                            if (feed != 1)
                            {
                                motpos = motpos + abqi * stepinc / 2;
                                abqi = 0;
                            }
                            if (deney != 2 || deney != 1)
                            {
                                receive = abq;
                            }
                            else
                            {
                                File.AppendAllText(fileopen + "_Motor_Position.txt", abq + Environment.NewLine);
                            }
                            abq = "";
                            if (comread.Contains("LM"))
                            {
                                for (int i = comsay + 2; i < nbyt; i++)
                                {
                                    if (comread[i].ToString() == "L")
                                    {
                                        break;
                                    }
                                    abq += comread[i].ToString();
                                }
                                loadd.Add(Convert.ToDouble(abq));
                                if (maxload < loadd.Last())
                                {
                                    maxload = loadd.Last();
                                }
                                if (deney != -1 && deney != 0)
                                {
                                    File.AppendAllText(fileopen + "_Load.txt", loadd.Last().ToString() + Environment.NewLine);
                                }
                            }
                            if (feed == 1 || autopass == 1)
                            {
                                pass = 1;
                                timer6.Start();
                            }
                            if (jsBox.Checked)
                            {
                                pass = 1;
                                autopass = 1;
                                timer6.Start();
                            }
                            abq = "";
                            repsay = 0;
                        }
                        catch (Exception ex)
                        {
                            connection = ex.Message;
                            //if (repsay < 3)
                            //{
                            //    Send("Repet", "");
                            //    repsay++;
                            //}
                            //else
                            //{
                            //    repsay = 0;
                            //}
                        }
                    }
                    else if (comread.Contains("LM"))
                    {
                        abq = "";
                        for (int i = 0; i < nbyt; i++)
                        {
                            if (comread[i].ToString() == "L")
                            {
                                break;
                            }
                            abq += comread[i].ToString();
                        }
                        try
                        {
                            receives = abq;
                            loadd.Add(Convert.ToDouble(abq));
                            if (maxload < loadd.Last())
                            {
                                maxload = loadd.Last();
                            }
                            if (deney != -1 && deney != 0)
                            {
                                File.AppendAllText(fileopen + "_Load.txt", loadd.Last().ToString() + Environment.NewLine);
                                File.AppendAllText(fileopen + "_Load.txt", loadd.Last().ToString() + Environment.NewLine);
                            }
                            repsay = 0;
                        }
                        catch (Exception ex)
                        {
                            connection = ex.Message;
                            //if (repsay < 5)
                            //{
                            //    Send("Repet", "");
                            //    repsay++;
                            //}
                            //else
                            //{
                            //    repsay = 0;
                            //}
                        }
                        abq = "";
                        if (hxcom == 1)
                        {
                            pass = 1;
                            timer6.Start();
                        }

                    }
                    else if (comread.Contains("R5"))
                    {
                        if (directpass == 1 && plug % 2 == 1)
                        {
                            Send(texts, "");
                            directpass = 0;
                            texts = "";
                            expresscon = 1;
                            tim1say = 0;
                        }
                        else
                        {
                            string stopmes = directMes;
                            char initialch = stopmes[0];
                            stopmes = stopmes.Replace(initialch, Convert.ToChar("Z"));
                            Send(stopmes, "");
                        }
                    }
                    else if (comread.Contains("R3"))
                    {
                        if (motorpass == 1 && plug % 2 == 1)
                        {
                            timer6.Start();
                            Send(texts, "");
                            motorpass = 0;
                            texts = "";
                            autopass = 1;
                            comread = "";
                            tim1say = 0;
                        }

                        else
                        {
                            Send("Z0000", "");
                        }
                    }
                    else if ((comread.Contains("R0") || comread.Contains("Received")))
                    {
                        if(startexp==1 && exppass == 0)
                        {

                        }
                        else if (exppass == 1 && plug % 2 == 1)
                        {
                            Send(textexp[expcounter], "");
                            expcounter++;
                            if (expcounter == depth.Count)
                            {
                                expcounter = 0;
                                exppass = -1;
                                textexp.Clear();
                            }
                            tim1say = 0;
                            comread = "";
                        }
                        else
                        {
                            Send($"Z{0:D49}{0:99}", "");
                            MessageBox.Show("Unauthorized Experiment Start Detected!");
                        }
                    }
                    else if (comread.Contains("R1"))
                    {
                        if (calpass == 1 && plug % 2 == 1)
                        {
                            Send(texts, "");
                            calpass = 0;
                            texts = "";
                            comread = "";
                            tim1say = 0;
                            timer6.Stop();
                        }
                        else
                        {
                            Send("Z0000000000000000000000000000000000000000", "");
                            MessageBox.Show("Unauthorized Experiment Start Detected!");
                        }
                    }
                    else if (comread.Contains("StepReady"))
                    {
                        if (exppass == -1 && plug % 2 == 1)
                        {
                            pass = 1;
                            timer6.Start();
                            tim1say = 0;
                            comread = "";
                        }
                        else if (exppass == -2)
                        {
                        }
                        else
                        {
                            Send($"Z0000000000000|0000000", "");
                            MessageBox.Show("Unauthorized Experiment Start Detected!");
                        }
                    }
                    else if ((comread.Contains("+") || comread.Contains("-")))
                    {
                        if (xymotor == 1 || jsBox.Checked||deney==1)
                        {
                            pass = 1;
                            timer6.Start();
                            receive = comread;
                            comread = "";
                        }
                    }
                    else if (comread[nbyt - 1].ToString() == "C")
                    {
                        if (senscom == 1)
                        {
                            abq = "";
                            for (int i = 0; i < nbyt; i++)
                            {
                                if (comread[i].ToString() == "C")
                                {
                                    break;
                                }
                                abq += comread[i].ToString();
                            }
                            receives = abq;
                            pass = 1;
                            timer6.Start();
                            comread = "";
                        }
                    }
                    else if (comread[nbyt - 1].ToString() == "T")
                    {
                        abq = "";
                        comsay = 0;
                        for (int i = 0; i < nbyt; i++)
                        {
                            comsay++;
                            if (comread[i].ToString() == "H")
                            {
                                break;
                            }
                            abq += comread[i].ToString();
                        }
                        receives = abq;
                        try
                        {
                            humid = Convert.ToDouble(abq);
                        }
                        catch
                        {

                        }
                        abq = "";
                        for (int i = comsay; i < nbyt; i++)
                        {
                            if (comread[i].ToString() == "T")
                            {
                                break;
                            }
                            abq += comread[i].ToString();
                        }
                        receives += abq;
                        comsay = 0;
                        try
                        {
                            tmed = Convert.ToDouble(abq);
                        }
                        catch
                        {

                        }
                        abq = "";
                        comread = "";
                        if (rhcom == 1)
                        {
                            pass = 1;
                            timer6.Start();
                        }
                    }
                    else if (comread.Contains("GY"))
                    {
                        if (balancing == 1)
                        {
                            abq = "";
                            if (datacomp == 0)
                            {
                                for (int i = 0; i < nbyt; i++)
                                {
                                    if (comread[i].ToString() == "X")
                                    {
                                        break;
                                    }
                                    abq += comread[i].ToString();
                                }
                                try
                                {
                                    double tray = Convert.ToDouble(abq);
                                    if (abq + "XGY" == comread)
                                    {
                                        recx = abq;
                                        abq = "";
                                        datacomp++;
                                        Send("CDAT1", "");
                                    }
                                    else
                                    {
                                        //Send("Repet", "");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    //Send("Repet", "");
                                }

                            }
                            else if (datacomp == 1)
                            {
                                for (int i = 0; i < nbyt; i++)
                                {
                                    if (comread[i].ToString() == "Y")
                                    {
                                        break;
                                    }
                                    abq += comread[i].ToString();
                                }
                                try
                                {
                                    double tray = Convert.ToDouble(abq);
                                    if (abq + "YGY" == comread)
                                    {
                                        recy = abq;
                                        abq = "";
                                        datacomp++;
                                        Send("CDAT2", "");
                                    }
                                    else
                                    {
                                        Send("CDAT1", "");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Send("CDAT1", "");
                                }

                            }
                            else if (datacomp == 2)
                            {
                                for (int i = 0; i < nbyt; i++)
                                {
                                    if (comread[i].ToString() == "Z")
                                    {
                                        break;
                                    }
                                    abq += comread[i].ToString();
                                }
                                try
                                {
                                    double tray = Convert.ToDouble(abq);
                                    if (abq + "ZGY" == comread)
                                    {
                                        recz = abq;
                                        pass = 2;
                                        timer6.Start();
                                        comread = "";
                                        datacomp = 0;
                                        balancing = 0;
                                        Send("CDFIN", "");
                                    }
                                    else
                                    {
                                        Send("CDAT2", "");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    //Send("CDAT2", "");
                                }
                            }
                        }
                    }
                    else if (comread.Contains("PROCESS"))
                    {
                        timer1.Stop();
                        receive = comread;
                        texts = "";
                        directpass = 0;
                        exppass = 0;
                        calpass = 0;
                        motorpass = 0;
                        connection = comread;
                        comread = "";
                        approaching = 0;
                        textexp.Clear();
                        repsay = 0;
                        timer6.Start();
                    }
                    else
                    {
                        connection = comread;
                        repsay = 0;
                        timer6.Start();
                    }
                }
            }
        }
        double humid = 0;
        double tmed = 0;
        int datacomp = 0;
        string recx, recy, recz;
        double[,] balance=new double[3, 2] { {0.0, 0.0}, {0.0, 0.0}, {0.0 , 0.0} };
        public List<double> vapp = new List<double>();
        public List<double> loadd = new List<double>();
        private async Task CalibrationAsync(string caldepth, string calstep, string calinter)
        {
            var elapsed = new Stopwatch();
            int time = 100;
            double stepdepth = 0;
            double stepvolt = 0;
            int stepnum = 0;
            double firstpt = Convert.ToDouble(textBox1.Text);
            double firstz = Convert.ToDouble(textBox10.Text);
            double zvar = 0;
            if (caldepth != "" && calstep != "")
            {
                stepnum = Convert.ToInt32(Convert.ToDouble(calstep));
                if (stepnum != 0)
                {
                    stepdepth = Math.Round(Convert.ToDouble(caldepth) / (stepnum * kp * 20), 7);
                }
            }
            if (calinter != "")
            {
                if (Convert.ToDouble(calinter) >= 0.1)
                {
                    time = Convert.ToInt32(Convert.ToDouble(calinter) * 1000);
                }
            }
            elapsed.Start();
            for (int i = 0; i < 2 * stepnum; i++)
            {
                if (i < stepnum)
                {
                    if (emCounter % 2 != 0)
                    {
                        stepvolt = Math.Round(firstpt - stepdepth * (i + 1), 7);
                        Send(TextFormat(stepvolt.ToString(), "", -1, 7.5), "W");
                        await Task.Delay(time);
                        timeshow.Text = elapsed.ElapsedMilliseconds.ToString();
                        elapsed.Restart();
                        textBox1.Text = stepvolt.ToString();
                        //textBox1.Update();
                        //textBox10.Update();
                        textBox23.Text = textBox1.Text;
                        //textBox23.Update();
                    }
                    else
                    {
                        elapsed.Reset();
                        break;
                    }
                }
                else if (i >= stepnum && returnCheck.Checked)
                {
                    if (emCounter % 2 != 0)
                    {
                        stepvolt = Math.Round(stepvolt + stepdepth, 7);
                        if (checkBox15.Checked)
                        {
                            Send(TextFormat(stepvolt.ToString(), "", -1, 7.5), "W");
                        }
                        await Task.Delay(time);
                        timeshow.Text = elapsed.ElapsedMilliseconds.ToString();
                        elapsed.Restart();
                        textBox1.Text = stepvolt.ToString();
                        //textBox1.Update();
                        //textBox10.Update();
                        textBox23.Text = textBox1.Text;
                        //textBox23.Update();
                    }
                    else
                    {
                        elapsed.Reset();
                        break;
                    }
                }
                else
                {
                    elapsed.Reset();
                    break;
                }

            }
            groupBox8.Enabled = true;
            textBox1.Text = stepvolt.ToString();
        }

        private void cjcSourceComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cjcSourceComboBox.SelectedIndex)
            {
                case 0: // Channel
                    cjcChannelTextBox.Enabled = true;
                    cjcValueNumeric.Enabled = false;
                    break;
                case 1: // Constant 
                    cjcChannelTextBox.Enabled = false;
                    cjcValueNumeric.Enabled = true;
                    break;
                case 2: // Internal
                    cjcChannelTextBox.Enabled = false;
                    cjcValueNumeric.Enabled = false;
                    break;
            }
        }
        int samplePerchannel = 10;
        private void startButton_Click(object sender, EventArgs e)
        {
            deney = 0;
            startButton.Enabled = false;
            stopButton.Enabled = true;
            try
            {
                myTask = new NationalInstruments.DAQmx.Task();
                if (measureType.SelectedIndex == 1)
                {
                    AIThermocoupleType thermocoupleType;
                    switch (thermocoupleTypeComboBox.SelectedIndex)
                    {
                        case 0:
                            thermocoupleType = AIThermocoupleType.B;
                            break;
                        case 1:
                            thermocoupleType = AIThermocoupleType.E;
                            break;
                        case 2:
                            thermocoupleType = AIThermocoupleType.J;
                            break;
                        case 3:
                            thermocoupleType = AIThermocoupleType.K;
                            break;
                        case 4:
                            thermocoupleType = AIThermocoupleType.N;
                            break;
                        case 5:
                            thermocoupleType = AIThermocoupleType.R;
                            break;
                        case 6:
                            thermocoupleType = AIThermocoupleType.S;
                            break;
                        case 7:
                        default:
                            thermocoupleType = AIThermocoupleType.T;
                            break;
                    }

                    switch (cjcSourceComboBox.SelectedIndex)
                    {
                        case 0: // Channel
                            myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Text,
                               "", Convert.ToDouble(minimumValueNumeric.Value), Convert.ToDouble(maximumValueNumeric.Value),
                               thermocoupleType, AITemperatureUnits.DegreesC, cjcChannelTextBox.Text);
                            break;
                        case 1: // Constant
                            myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Text,
                               "", Convert.ToDouble(minimumValueNumeric.Value), Convert.ToDouble(maximumValueNumeric.Value),
                               thermocoupleType, AITemperatureUnits.DegreesC, Convert.ToDouble(cjcValueNumeric.Value));
                            break;
                        case 2: // Internal
                            myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Text,
                                "", Convert.ToDouble(minimumValueNumeric.Value), Convert.ToDouble(maximumValueNumeric.Value),
                                thermocoupleType, AITemperatureUnits.DegreesC);
                            break;
                        default:
                            myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Text,
                               "", Convert.ToDouble(minimumValueNumeric.Value), Convert.ToDouble(maximumValueNumeric.Value),
                               thermocoupleType, AITemperatureUnits.DegreesC, 25);
                            break;
                    }
                }
                else if (measureType.SelectedIndex == 0)
                {
                    myTask.AIChannels.CreateVoltageChannel(physicalChannelComboBox.Text, "", (AITerminalConfiguration)(-1), Convert.ToDouble(minimumValueNumeric.Value), Convert.ToDouble(maximumValueNumeric.Value), AIVoltageUnits.Volts);
                }
                else if (measureType.SelectedIndex == 2)
                {
                    AIRtdType aIRtd;
                    AIExcitationSource aIExc;
                    switch (rtdBox.SelectedIndex)
                    {
                        case 0:
                            aIRtd = AIRtdType.Pt3750;
                            break;
                        case 1:
                            aIRtd = AIRtdType.Pt3851;
                            break;
                        case 2:
                            aIRtd = AIRtdType.Pt3911;
                            break;
                        case 3:
                            aIRtd = AIRtdType.Pt3916;
                            break;
                        case 4:
                            aIRtd = AIRtdType.Pt3920;
                            break;
                        case 5:
                            aIRtd = AIRtdType.Pt3928;
                            break;
                        default:
                            aIRtd = AIRtdType.Pt3851;
                            break;
                    }
                    switch (rtdExcType.SelectedIndex)
                    {
                        case 0:
                            aIExc = AIExcitationSource.External;
                            break;
                        case 1:
                            aIExc = AIExcitationSource.Internal;
                            break;
                        default:
                            aIExc = AIExcitationSource.External;
                            break;
                    }
                    myTask.AIChannels.CreateRtdChannel(physicalChannelComboBox.Text, "", Convert.ToDouble(minimumValueNumeric.Value),
                        Convert.ToDouble(maximumValueNumeric.Value), AITemperatureUnits.DegreesC, aIRtd,
                        AIResistanceConfiguration.TwoWire, aIExc, Convert.ToDouble(rtdExc.Text), Convert.ToDouble(rtdRes.Text));
                }
                else if (measureType.SelectedIndex == 3)
                {
                    AIStrainGageConfiguration aIStrain;
                    AIExcitationSource aIExcitation;
                    switch (bridgeBox.SelectedIndex)
                    {
                        case 0:
                            aIStrain = AIStrainGageConfiguration.FullBridgeI;
                            break;
                        case 1:
                            aIStrain = AIStrainGageConfiguration.FullBridgeII;
                            break;
                        case 2:
                            aIStrain = AIStrainGageConfiguration.FullBridgeIII;
                            break;
                        case 3:
                            aIStrain = AIStrainGageConfiguration.HalfBridgeI;
                            break;
                        case 4:
                            aIStrain = AIStrainGageConfiguration.HalfBridgeII;
                            break;
                        case 5:
                            aIStrain = AIStrainGageConfiguration.QuarterBridgeI;
                            break;
                        case 6:
                            aIStrain = AIStrainGageConfiguration.QuarterBridgeII;
                            break;
                        default:
                            aIStrain = AIStrainGageConfiguration.FullBridgeI;
                            break;
                    }
                    switch (excBox.SelectedIndex)
                    {
                        case 0:
                            aIExcitation = AIExcitationSource.External;
                            break;
                        case 1:
                            aIExcitation = AIExcitationSource.Internal;
                            break;
                        case 2:
                            aIExcitation = AIExcitationSource.None;
                            break;
                        default:
                            aIExcitation = AIExcitationSource.Internal;
                            break;
                    }
                    myTask.AIChannels.CreateStrainGageChannel(physicalChannelComboBox.Text, "", Convert.ToDouble(minimumValueNumeric.Value),
                        Convert.ToDouble(maximumValueNumeric.Value), aIStrain, aIExcitation, Convert.ToDouble(gageExc.Value) / 1000,
                        Convert.ToDouble(gageFac.Value)/1000000, Convert.ToDouble(gageVol.Value) / 1000, Convert.ToDouble(gageNom.Value),
                        Convert.ToDouble(gagePoisson.Text), Convert.ToDouble(gageWires.Value) / 1000, AIStrainUnits.Strain);
                }
                myTask.Timing.ConfigureSampleClock("", Convert.ToDouble(rateNumeric.Value),
                    SampleClockActiveEdge.Rising, SampleQuantityMode.ContinuousSamples, 1000);

                try
                {
                    myTask.Control(TaskAction.Verify);
                    analogInReader = new AnalogMultiChannelReader(myTask.Stream);
                    runningTask = myTask;
                    analogInReader.SynchronizeCallbacks = true;
                    analogInReader.BeginReadWaveform(samplePerchannel, myAsyncCallback, runningTask);
                    if (rateNumeric.Value <= 10000)
                    {
                        timer2.Interval = 100;
                    }
                    else
                    {
                        timer2.Interval = 100;
                    }
                    datas.Insert(0, 0);
                    timer2.Start();
                }
                catch (DaqException exception)
                {
                    MessageBox.Show(exception.Message);
                    runningTask = null;
                    startButton.Enabled = true;
                    stopButton.Enabled = false;
                    executeCal.Enabled = true;
                    deney = -1;
                    myTask.Dispose();
                }
            }
            catch (DaqException exception)
            {
                timer2.Stop();
                MessageBox.Show(exception.Message);
                startButton.Enabled = true;
                stopButton.Enabled = false;
                executeCal.Enabled = true;
                deney = -1;
                runningTask = null;
                myTask.Dispose();
            }
        }
        private void stopButton_Click(object sender, EventArgs e)
        {
            timer2.Stop();
            runningTask = null;
            myTask.Dispose();
            stopButton.Enabled = false;
            startButton.Enabled = true;
            deney = -1;
            executeCal.Enabled = true;
        }
        int tim2i = 0;
        double tim2tim = 0;
        int showpst = 0;
        int chartptcnt = 0;
        private void timer2_Tick(object sender, EventArgs e)
        {
            tim2tim = tim2i * timer2.Interval/1000.0;
            if (chartptcnt > 200)
            {
                for (int i = 0; i < chart1.Series.Count; i++)
                {
                    chart1.Series[i].Points.Clear();
                }
                chartptcnt = 0;
            }
            if (deney == 0 && datas != null)
            {
                TestBox.Text=datas[0].ToString();
            }
            else if (deney != -1 && datas != null)
            {
                for (int i = 0; i < datas.Count(); i++)
                {
                    if (temperChan == phychannel[i])
                    {
                        tempshow.Text = datas[i].ToString();
                        if (maxtemp < datas[i])
                        {
                            maxtemp = datas[i];
                        }
                    }
                    else if (voltagechannel == phychannel[i])
                    {
                        gageShow.Text = datas[i].ToString();
                    }
                    if (plotstop == 0)
                    {
                        chart1.Series[kan + 1 - datas.Count() + i].Points.AddXY(tim2tim,datas[i]);
                        chartptcnt++;
                    }
                }
            }
            if (deney != -1 && deney!=0 && loadd.Count > 0)
            {
                try
                {

                    if (plotstop == 0 && chtype != null && (loadExt || loadLive.Checked))
                    {
                        if(loadd.Last()<33000000 && loadd.Last() > -33000000)
                        {
                            chart1.Series[chart1.Series.Count-2].Points.AddXY(tim2tim, loadd.Last());
                            chartptcnt++;
                        }
                    }
                    pressShow.Text = loadd.Last().ToString();
                }
                catch
                {

                }
            }
            if (deney != -1 && deney!=0)
            {
                if (motorDrive.Checked)
                {
                    try
                    {
                        if (plotstop == 0 && showpst == 1)
                        {

                            chart1.Series.Last().Points.AddXY(tim2tim,
(motpos - motpos0) / 1000);
                            chartptcnt++;
                        }
                    }
                    catch { }
                }
                else if (motorApp.Checked && approaching == 2)
                {
                    try
                    {
                        if (plotstop == 0 && showpst == 1)
                        {

                            chart1.Series.Last().Points.AddXY(tim2tim,
(motpos - motpos0) / 1000);
                            chartptcnt++;

                        }
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        if (plotstop == 0 && showpst == 1)
                        {

                            chart1.Series.Last().Points.AddXY(tim2tim,
                                    (zpos - zpos0));
                            chartptcnt++;

                        }
                    }
                    catch { }
                }
            }
            tim2i++;
            if (emCounter % 2 == 0)
            {
                timer2.Stop();
            }
        }
        List<int> phychannel = new List<int>();
        List<double> maxval = new List<double>();
        List<double> minval = new List<double>();
        List<int> cjctype = new List<int>();
        List<double> cjcvalue = new List<double>();
        List<string> cjcchannel = new List<string>();
        List<int> tctype = new List<int>();
        List<int> chtype = new List<int>();
        List<int> rtdtype = new List<int>();
        List<double> rtdexc = new List<double>();
        List<double> rtdres = new List<double>();
        List<double> gagexc = new List<double>();
        List<double> gageres = new List<double>();
        List<double> gagefac = new List<double>();
        List<double> gageinit = new List<double>();
        List<double> gagewire = new List<double>();
        List<double> gagepois = new List<double>();
        List<int> gagetype = new List<int>();
        List<int> exctype = new List<int>();
        List<int> rtdexctype = new List<int>();
        List<bool> isFeed = new List<bool>();
        private void button1_Click(object sender, EventArgs e)
        {
            int sel = tcNo.SelectedIndex;
            if (phychannel.Contains(physicalChannelComboBox.SelectedIndex))
            {
                MessageBox.Show("This channel has been already set!", "DAQmx Channel Settings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (measureType.SelectedIndex == 0)
            {
                phychannel.Insert(tcNo.SelectedIndex, physicalChannelComboBox.SelectedIndex);
                maxval.Insert(tcNo.SelectedIndex, Convert.ToDouble(maximumValueNumeric.Value));
                minval.Insert(tcNo.SelectedIndex, Convert.ToDouble(minimumValueNumeric.Value));
                cjctype.Insert(tcNo.SelectedIndex, 0);
                tctype.Insert(tcNo.SelectedIndex, 0);
                cjcvalue.Insert(tcNo.SelectedIndex, 0);
                cjcchannel.Insert(tcNo.SelectedIndex, "-");
                rtdtype.Insert(tcNo.SelectedIndex, 0);
                rtdexc.Insert(tcNo.SelectedIndex, 0);
                rtdres.Insert(tcNo.SelectedIndex, 0);
                gagexc.Insert(tcNo.SelectedIndex, 0);
                gageres.Insert(tcNo.SelectedIndex, 0);
                gagefac.Insert(tcNo.SelectedIndex, 0);
                gageinit.Insert(tcNo.SelectedIndex, 0);
                gagewire.Insert(tcNo.SelectedIndex, 0);
                gagepois.Insert(tcNo.SelectedIndex, 0);
                gagetype.Insert(tcNo.SelectedIndex, 0);
                exctype.Insert(tcNo.SelectedIndex, 0);
                rtdexctype.Insert(tcNo.SelectedIndex, 0);
                chtype.Insert(tcNo.SelectedIndex, measureType.SelectedIndex);
                isFeed.Insert(tcNo.SelectedIndex, isHeaterFeed.Checked);
                if (phychannel.Count >= tcNo.SelectedIndex && tcNo.Items.Count <= tcNo.SelectedIndex + 1)
                {
                    tcNo.Items.Add(tcNo.SelectedIndex + 2);
                }
                button4.Enabled = true;
                if (phychannel.Count > tcNo.SelectedIndex + 1)
                {
                    phychannel.RemoveAt(tcNo.SelectedIndex + 1);
                    maxval.RemoveAt(tcNo.SelectedIndex + 1);
                    minval.RemoveAt(tcNo.SelectedIndex + 1);
                    cjcchannel.RemoveAt(tcNo.SelectedIndex + 1);
                    cjctype.RemoveAt(tcNo.SelectedIndex + 1);
                    cjcvalue.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdtype.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdres.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdexc.RemoveAt(tcNo.SelectedIndex + 1);
                    gagetype.RemoveAt(tcNo.SelectedIndex + 1);
                    gageres.RemoveAt(tcNo.SelectedIndex + 1);
                    gagexc.RemoveAt(tcNo.SelectedIndex + 1);
                    gagewire.RemoveAt(tcNo.SelectedIndex + 1);
                    gageinit.RemoveAt(tcNo.SelectedIndex + 1);
                    gagefac.RemoveAt(tcNo.SelectedIndex + 1);
                    gagepois.RemoveAt(tcNo.SelectedIndex + 1);
                    exctype.RemoveAt(tcNo.SelectedIndex + 1);
                    chtype.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdexctype.RemoveAt(tcNo.SelectedIndex + 1);
                    isFeed.RemoveAt(tcNo.SelectedIndex + 1);
                }
                tcNo.SelectedIndex++;
            }
            else if (measureType.SelectedIndex == 1)
            {
                phychannel.Insert(tcNo.SelectedIndex, physicalChannelComboBox.SelectedIndex);
                maxval.Insert(tcNo.SelectedIndex, Convert.ToDouble(maximumValueNumeric.Value));
                minval.Insert(tcNo.SelectedIndex, Convert.ToDouble(minimumValueNumeric.Value));
                cjctype.Insert(tcNo.SelectedIndex, cjcSourceComboBox.SelectedIndex);
                tctype.Insert(tcNo.SelectedIndex, thermocoupleTypeComboBox.SelectedIndex);
                cjcvalue.Insert(tcNo.SelectedIndex, Convert.ToDouble(cjcValueNumeric.Value));
                cjcchannel.Insert(tcNo.SelectedIndex, cjcChannelTextBox.Text);
                rtdtype.Insert(tcNo.SelectedIndex, 0);
                rtdexc.Insert(tcNo.SelectedIndex, 0);
                rtdres.Insert(tcNo.SelectedIndex, 0);
                gagexc.Insert(tcNo.SelectedIndex, 0);
                gageres.Insert(tcNo.SelectedIndex, 0);
                gagefac.Insert(tcNo.SelectedIndex, 0);
                gageinit.Insert(tcNo.SelectedIndex, 0);
                gagewire.Insert(tcNo.SelectedIndex, 0);
                gagepois.Insert(tcNo.SelectedIndex, 0);
                gagetype.Insert(tcNo.SelectedIndex, 0);
                exctype.Insert(tcNo.SelectedIndex, 0);
                rtdexctype.Insert(tcNo.SelectedIndex, 0);
                chtype.Insert(tcNo.SelectedIndex, measureType.SelectedIndex);
                isFeed.Insert(tcNo.SelectedIndex, isHeaterFeed.Checked);
                if (phychannel.Count >= tcNo.SelectedIndex && tcNo.Items.Count <= tcNo.SelectedIndex + 1)
                {
                    tcNo.Items.Add(tcNo.SelectedIndex + 2);
                }
                button4.Enabled = true;
                if (phychannel.Count > tcNo.SelectedIndex + 1)
                {
                    phychannel.RemoveAt(tcNo.SelectedIndex + 1);
                    maxval.RemoveAt(tcNo.SelectedIndex + 1);
                    minval.RemoveAt(tcNo.SelectedIndex + 1);
                    cjcchannel.RemoveAt(tcNo.SelectedIndex + 1);
                    cjctype.RemoveAt(tcNo.SelectedIndex + 1);
                    cjcvalue.RemoveAt(tcNo.SelectedIndex + 1);
                    tctype.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdtype.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdres.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdexc.RemoveAt(tcNo.SelectedIndex + 1);
                    gagetype.RemoveAt(tcNo.SelectedIndex + 1);
                    gageres.RemoveAt(tcNo.SelectedIndex + 1);
                    gagexc.RemoveAt(tcNo.SelectedIndex + 1);
                    gagewire.RemoveAt(tcNo.SelectedIndex + 1);
                    gageinit.RemoveAt(tcNo.SelectedIndex + 1);
                    gagefac.RemoveAt(tcNo.SelectedIndex + 1);
                    gagepois.RemoveAt(tcNo.SelectedIndex + 1);
                    chtype.RemoveAt(tcNo.SelectedIndex + 1);
                    exctype.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdexctype.RemoveAt(tcNo.SelectedIndex + 1);
                    isFeed.RemoveAt(tcNo.SelectedIndex + 1);
                }
                tcNo.SelectedIndex++;
            }
            else if (measureType.SelectedIndex == 2)
            {
                phychannel.Insert(tcNo.SelectedIndex, physicalChannelComboBox.SelectedIndex);
                maxval.Insert(tcNo.SelectedIndex, Convert.ToDouble(maximumValueNumeric.Value));
                minval.Insert(tcNo.SelectedIndex, Convert.ToDouble(minimumValueNumeric.Value));
                cjctype.Insert(tcNo.SelectedIndex, 0);
                tctype.Insert(tcNo.SelectedIndex, 0);
                cjcvalue.Insert(tcNo.SelectedIndex, 0);
                cjcchannel.Insert(tcNo.SelectedIndex, "-");
                rtdtype.Insert(tcNo.SelectedIndex, rtdBox.SelectedIndex);
                rtdexc.Insert(tcNo.SelectedIndex, Convert.ToDouble(rtdExc.Text));
                rtdres.Insert(tcNo.SelectedIndex, Convert.ToDouble(rtdRes.Text));
                rtdexctype.Insert(tcNo.SelectedIndex, rtdExcType.SelectedIndex);
                gagexc.Insert(tcNo.SelectedIndex, 0);
                gageres.Insert(tcNo.SelectedIndex, 0);
                gagefac.Insert(tcNo.SelectedIndex, 0);
                gageinit.Insert(tcNo.SelectedIndex, 0);
                gagewire.Insert(tcNo.SelectedIndex, 0);
                gagepois.Insert(tcNo.SelectedIndex, 0);
                gagetype.Insert(tcNo.SelectedIndex, 0);
                exctype.Insert(tcNo.SelectedIndex, 0);
                chtype.Insert(tcNo.SelectedIndex, measureType.SelectedIndex);
                isFeed.Insert(tcNo.SelectedIndex, isHeaterFeed.Checked);
                if (phychannel.Count >= tcNo.SelectedIndex && tcNo.Items.Count <= tcNo.SelectedIndex + 1)
                {
                    tcNo.Items.Add(tcNo.SelectedIndex + 2);
                }
                button4.Enabled = true;
                if (phychannel.Count > tcNo.SelectedIndex + 1)
                {
                    phychannel.RemoveAt(tcNo.SelectedIndex + 1);
                    maxval.RemoveAt(tcNo.SelectedIndex + 1);
                    minval.RemoveAt(tcNo.SelectedIndex + 1);
                    cjcchannel.RemoveAt(tcNo.SelectedIndex + 1);
                    cjctype.RemoveAt(tcNo.SelectedIndex + 1);
                    cjcvalue.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdtype.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdres.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdexc.RemoveAt(tcNo.SelectedIndex + 1);
                    gagetype.RemoveAt(tcNo.SelectedIndex + 1);
                    gageres.RemoveAt(tcNo.SelectedIndex + 1);
                    gagexc.RemoveAt(tcNo.SelectedIndex + 1);
                    gagewire.RemoveAt(tcNo.SelectedIndex + 1);
                    gageinit.RemoveAt(tcNo.SelectedIndex + 1);
                    gagefac.RemoveAt(tcNo.SelectedIndex + 1);
                    gagepois.RemoveAt(tcNo.SelectedIndex + 1);
                    chtype.RemoveAt(tcNo.SelectedIndex + 1);
                    exctype.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdexctype.RemoveAt(tcNo.SelectedIndex + 1);
                    isFeed.RemoveAt(tcNo.SelectedIndex + 1);
                }
                tcNo.SelectedIndex++;
            }
            else if (measureType.SelectedIndex == 3)
            {
                phychannel.Insert(tcNo.SelectedIndex, physicalChannelComboBox.SelectedIndex);
                maxval.Insert(tcNo.SelectedIndex, Convert.ToDouble(maximumValueNumeric.Value));
                minval.Insert(tcNo.SelectedIndex, Convert.ToDouble(minimumValueNumeric.Value));
                cjctype.Insert(tcNo.SelectedIndex, 0);
                tctype.Insert(tcNo.SelectedIndex, 0);
                cjcvalue.Insert(tcNo.SelectedIndex, 0);
                cjcchannel.Insert(tcNo.SelectedIndex, "-");
                rtdtype.Insert(tcNo.SelectedIndex, 0);
                rtdexc.Insert(tcNo.SelectedIndex, 0);
                rtdexctype.Insert(tcNo.SelectedIndex, 0);
                rtdres.Insert(tcNo.SelectedIndex, 0);
                gagexc.Insert(tcNo.SelectedIndex, Convert.ToDouble(gageExc.Value));
                gageres.Insert(tcNo.SelectedIndex, Convert.ToDouble(gageNom.Value));
                gagefac.Insert(tcNo.SelectedIndex, Convert.ToDouble(gageFac.Value));
                gageinit.Insert(tcNo.SelectedIndex, Convert.ToDouble(gageVol.Value));
                gagewire.Insert(tcNo.SelectedIndex, Convert.ToDouble(gageWires.Value));
                gagepois.Insert(tcNo.SelectedIndex, Convert.ToDouble(gagePoisson.Text));
                gagetype.Insert(tcNo.SelectedIndex, bridgeBox.SelectedIndex);
                exctype.Insert(tcNo.SelectedIndex, excBox.SelectedIndex);
                chtype.Insert(tcNo.SelectedIndex, measureType.SelectedIndex);
                isFeed.Insert(tcNo.SelectedIndex,isHeaterFeed.Checked);
                if (phychannel.Count >= tcNo.SelectedIndex && tcNo.Items.Count <= tcNo.SelectedIndex + 1)
                {
                    tcNo.Items.Add(tcNo.SelectedIndex + 2);
                }
                button4.Enabled = true;
                if (phychannel.Count > tcNo.SelectedIndex + 1)
                {
                    phychannel.RemoveAt(tcNo.SelectedIndex + 1);
                    maxval.RemoveAt(tcNo.SelectedIndex + 1);
                    minval.RemoveAt(tcNo.SelectedIndex + 1);
                    cjcchannel.RemoveAt(tcNo.SelectedIndex + 1);
                    cjctype.RemoveAt(tcNo.SelectedIndex + 1);
                    cjcvalue.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdtype.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdres.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdexc.RemoveAt(tcNo.SelectedIndex + 1);
                    gagetype.RemoveAt(tcNo.SelectedIndex + 1);
                    gageres.RemoveAt(tcNo.SelectedIndex + 1);
                    gagexc.RemoveAt(tcNo.SelectedIndex + 1);
                    gagewire.RemoveAt(tcNo.SelectedIndex + 1);
                    gageinit.RemoveAt(tcNo.SelectedIndex + 1);
                    gagefac.RemoveAt(tcNo.SelectedIndex + 1);
                    gagepois.RemoveAt(tcNo.SelectedIndex + 1);
                    chtype.RemoveAt(tcNo.SelectedIndex + 1);
                    exctype.RemoveAt(tcNo.SelectedIndex + 1);
                    rtdexctype.RemoveAt(tcNo.SelectedIndex + 1);
                    isFeed.RemoveAt(tcNo.SelectedIndex + 1);
                }
                tcNo.SelectedIndex++;
            }
            try
            {
                physicalChannelComboBox.SelectedIndex++;
            }
            catch (Exception ex)
            {
                label11.Text = ex.Message;
                for (int i = 0; i < physicalChannelComboBox.Items.Count; i++)
                {
                    if (!phychannel.Contains(i))
                    {
                        physicalChannelComboBox.SelectedIndex = i;
                        break;
                    }
                    else if(i== physicalChannelComboBox.Items.Count - 1)
                    {
                        MessageBox.Show("All Channels of NI PCI-e have been set!"
                            + Environment.NewLine + "Please use External Com.");
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int q = tcNo.SelectedIndex;
            if (phychannel.Count > q)
            {
                phychannel.RemoveAt(q);
                maxval.RemoveAt(q);
                minval.RemoveAt(q);
                cjcvalue.RemoveAt(q);
                tctype.RemoveAt(q);
                cjctype.RemoveAt(q);
                cjcchannel.RemoveAt(q);
                rtdtype.RemoveAt(q);
                rtdres.RemoveAt(q);
                rtdexc.RemoveAt(q);
                gagetype.RemoveAt(q);
                gageres.RemoveAt(q);
                gagexc.RemoveAt(q);
                gagewire.RemoveAt(q);
                gageinit.RemoveAt(q);
                gagefac.RemoveAt(q);
                gagepois.RemoveAt(q);
                chtype.RemoveAt(q);
                exctype.RemoveAt(q);
                rtdexctype.RemoveAt(q);
                isFeed.RemoveAt(q);
            }
            if (tcNo.Items.Count > 1)
            {
                tcNo.Items.RemoveAt(q);
                for (int i = 0; i < tcNo.Items.Count; i++)
                {
                    tcNo.Items[i] = i + 1;
                    if (tcNo.SelectedIndex > 0)
                    {
                        tcNo.SelectedIndex = q - 1;
                    }
                    else
                    {
                        tcNo.SelectedIndex = 0;
                    }

                }
            }
            if (phychannel.Count == 0)
            {
                button4.Enabled = false;
                btn9 = 0;
            }
        }
        bool tcnochange;
        private void TcNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            tcnochange = true;
            if (phychannel.Count > tcNo.SelectedIndex)
            {
                measureType.SelectedIndex = chtype[tcNo.SelectedIndex];
                physicalChannelComboBox.SelectedIndex = phychannel[tcNo.SelectedIndex];
                maximumValueNumeric.Value = Convert.ToDecimal(maxval[tcNo.SelectedIndex]);
                minimumValueNumeric.Value = Convert.ToDecimal(minval[tcNo.SelectedIndex]);
                cjcValueNumeric.Value = Convert.ToDecimal(cjcvalue[tcNo.SelectedIndex]);
                thermocoupleTypeComboBox.SelectedIndex = tctype[tcNo.SelectedIndex];
                cjcSourceComboBox.SelectedIndex = cjctype[tcNo.SelectedIndex];
                cjcChannelTextBox.Text = cjcchannel[tcNo.SelectedIndex];
                gagePoisson.Text = gagepois[tcNo.SelectedIndex].ToString();
                rtdExc.Text = rtdexc[tcNo.SelectedIndex].ToString();
                rtdRes.Text = rtdres[tcNo.SelectedIndex].ToString();
                gageExc.Value = Convert.ToDecimal(gagexc[tcNo.SelectedIndex]);
                gageNom.Value = Convert.ToDecimal(gageres[tcNo.SelectedIndex]);
                gageFac.Value = Convert.ToDecimal(gagefac[tcNo.SelectedIndex]);
                gageWires.Value = Convert.ToDecimal(gagewire[tcNo.SelectedIndex]);
                gageVol.Value = Convert.ToDecimal(gageinit[tcNo.SelectedIndex]);
                bridgeBox.SelectedIndex = gagetype[tcNo.SelectedIndex];
                excBox.SelectedIndex = exctype[tcNo.SelectedIndex];
                rtdExcType.SelectedIndex = rtdexctype[tcNo.SelectedIndex];
                groupBox11.Enabled = false;
                measureType.Enabled = false;
                gageBox.Enabled = false;
                isHeaterFeed.Checked = isFeed[tcNo.SelectedIndex];
            }
            else
            {
                gageBox.Enabled = true;
                measureType.Enabled = true;
                groupBox11.Enabled = true;
                measureType.SelectedIndex = -1;
            }
            tcnochange = false;
        }
        
        int btn9 = 0;
        int voltagechannel = -1;
        private void button9_Click(object sender, EventArgs e)
        {
            int sersay = chart1.Series.Count - 1;
            int feeder = 0;
            for (int i = sersay; i > -1; i--)
            {
                chart1.Series[i].Points.Clear();
                chart1.Series.RemoveAt(i);
                graphcount = 0;
                chartptcnt = 0;
            }
            int chansay = gageChan.Items.Count - 1;
            for (int i = chansay; i > 0; i--)
            {
                gageChan.Items.RemoveAt(i);
            }
            chansay = tempChan.Items.Count - 1;
            for (int i = chansay; i > 0; i--)
            {
                tempChan.Items.RemoveAt(i);
            }
            if (isFeed.Contains(true))
            {
                timer1.Stop();
                //for (int i = 0; i < 3; i++)
                //{
                //    if (activeHeat[i]&&!feedMode[i])
                //    {
                //        activeHeat[i] = false;
                //        Send($"9{i}000", "");
                //    }
                //}
                heaterFeedList.Items.Clear();
            }
            List<AIChannel> aIChannels = new List<AIChannel>();
            myTask = new NationalInstruments.DAQmx.Task();
            for (int i = 0; i < chtype.Count; i++)
            {
                try
                {
                    if (chtype[i] == 0)
                    {
                        aIChannels.Add(myTask.AIChannels.CreateVoltageChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
"Voltage" + phychannel[i].ToString(), (AITerminalConfiguration)(-1), minval[i], maxval[i], AIVoltageUnits.Volts));
                        if (isFeed[i])
                        {
                            heaterFeedList.Items.Add(aIChannels[i].VirtualName);
                            feedChannel[feeder] = i;
                            feeder++;
                        }
                        else
                        {
                            gageChan.Items.Add(aIChannels[i].VirtualName);
                        }

                    }
                    else if (chtype[i] == 1)
                    {
                        AIThermocoupleType thermocoupleType;
                        AIAutoZeroMode autoZeroMode;

                        switch (tctype[i])
                        {
                            case 0:
                                thermocoupleType = AIThermocoupleType.B;
                                break;
                            case 1:
                                thermocoupleType = AIThermocoupleType.E;
                                break;
                            case 2:
                                thermocoupleType = AIThermocoupleType.J;
                                break;
                            case 3:
                                thermocoupleType = AIThermocoupleType.K;
                                break;
                            case 4:
                                thermocoupleType = AIThermocoupleType.N;
                                break;
                            case 5:
                                thermocoupleType = AIThermocoupleType.R;
                                break;
                            case 6:
                                thermocoupleType = AIThermocoupleType.S;
                                break;
                            default:
                                thermocoupleType = AIThermocoupleType.T;
                                break;
                        }

                        switch (cjctype[i])
                        {
                            case 0: // Channel
                                aIChannels.Add(myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
"Temperature" + phychannel[i].ToString(), minval[i], maxval[i],
thermocoupleType, AITemperatureUnits.DegreesC, cjcChannelTextBox.Text));
                                if (isFeed[i])
                                {
                                    heaterFeedList.Items.Add(aIChannels[i].VirtualName);
                                    feedChannel[feeder] = i;
                                    feeder++;
                                }
                                else
                                {
                                    tempChan.Items.Add(aIChannels[i].VirtualName);
                                }
                                break;
                            case 1: // Constant
                                aIChannels.Add(myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
"Temperature" + phychannel[i].ToString(), minval[i], maxval[i],
thermocoupleType, AITemperatureUnits.DegreesC, Convert.ToDouble(cjcValueNumeric.Value)));
                                if (isFeed[i])
                                {
                                    heaterFeedList.Items.Add(aIChannels[i].VirtualName);
                                    feedChannel[feeder] = i;
                                    feeder++;
                                }
                                else
                                {
                                    tempChan.Items.Add(aIChannels[i].VirtualName);
                                }
                                break;

                            case 2: // Internal
                                aIChannels.Add(myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
     "Temperature" + phychannel[i].ToString(), minval[i], maxval[i],
     thermocoupleType, AITemperatureUnits.DegreesC, Convert.ToDouble(cjcValueNumeric.Value)));
                                if (isFeed[i])
                                {
                                    heaterFeedList.Items.Add(aIChannels[i].VirtualName);
                                    feedChannel[feeder] = i;
                                    feeder++;
                                }
                                else
                                {
                                    tempChan.Items.Add(aIChannels[i].VirtualName);
                                }
                                break;
                            default:
                                aIChannels.Add(myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
"Temperature" + phychannel[i].ToString(), minval[i], maxval[i],
thermocoupleType, AITemperatureUnits.DegreesC));
                                if (isFeed[i])
                                {
                                    heaterFeedList.Items.Add(aIChannels[i].VirtualName);
                                    feedChannel[feeder] = i;
                                    feeder++;
                                }
                                else
                                {
                                    tempChan.Items.Add(aIChannels[i].VirtualName);
                                }
                                break;
                        }
                        autoZeroMode = AIAutoZeroMode.None;
                        aIChannels[i].AutoZeroMode = autoZeroMode;
                    }
                    else if (chtype[i] == 2)//rtd temp measurement
                    {
                        AIRtdType aIRtd;
                        AIExcitationSource aIExc;
                        switch (rtdtype[i])
                        {
                            case 0:
                                aIRtd = AIRtdType.Pt3750;
                                break;
                            case 1:
                                aIRtd = AIRtdType.Pt3851;
                                break;
                            case 2:
                                aIRtd = AIRtdType.Pt3911;
                                break;
                            case 3:
                                aIRtd = AIRtdType.Pt3916;
                                break;
                            case 4:
                                aIRtd = AIRtdType.Pt3920;
                                break;
                            case 5:
                                aIRtd = AIRtdType.Pt3928;
                                break;
                            default:
                                aIRtd = AIRtdType.Pt3851;
                                break;
                        }
                        switch (rtdexctype[i])
                        {
                            case 0:
                                aIExc = AIExcitationSource.External;
                                break;
                            case 1:
                                aIExc = AIExcitationSource.Internal;
                                break;
                            default:
                                aIExc = AIExcitationSource.External;
                                break;
                        }
                        aIChannels.Add(myTask.AIChannels.CreateRtdChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
        "Temperature" + phychannel[i].ToString(), minval[i], maxval[i], AITemperatureUnits.DegreesC, aIRtd,
        AIResistanceConfiguration.TwoWire, aIExc, rtdexc[i], rtdres[i]));
                        if (isFeed[i])
                        {
                            heaterFeedList.Items.Add(aIChannels[i].VirtualName);
                            feedChannel[feeder] = i;
                            feeder++;
                        }
                        else
                        {
                            tempChan.Items.Add(aIChannels[i].VirtualName);
                        }
                    }
                    else if (chtype[i] == 3)//strain gage measurement
                    {
                        AIStrainGageConfiguration aIStrain;
                        AIExcitationSource aIExcitation;
                        switch (gagetype[i])
                        {
                            case 0:
                                aIStrain = AIStrainGageConfiguration.FullBridgeI;
                                break;
                            case 1:
                                aIStrain = AIStrainGageConfiguration.FullBridgeII;
                                break;
                            case 2:
                                aIStrain = AIStrainGageConfiguration.FullBridgeIII;
                                break;
                            case 3:
                                aIStrain = AIStrainGageConfiguration.HalfBridgeI;
                                break;
                            case 4:
                                aIStrain = AIStrainGageConfiguration.HalfBridgeII;
                                break;
                            case 5:
                                aIStrain = AIStrainGageConfiguration.QuarterBridgeI;
                                break;
                            case 6:
                                aIStrain = AIStrainGageConfiguration.QuarterBridgeII;
                                break;
                            default:
                                aIStrain = AIStrainGageConfiguration.FullBridgeI;
                                break;
                        }
                        switch (exctype[i])
                        {
                            case 0:
                                aIExcitation = AIExcitationSource.External;
                                break;
                            case 1:
                                aIExcitation = AIExcitationSource.Internal;
                                break;
                            case 2:
                                aIExcitation = AIExcitationSource.None;
                                break;
                            default:
                                aIExcitation = AIExcitationSource.Internal;
                                break;
                        }
                        aIChannels.Add(myTask.AIChannels.CreateStrainGageChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
"Strain" + phychannel[i].ToString(), minval[i], maxval[i], aIStrain, aIExcitation, gagexc[i] / 1000.0,
gagefac[i] / 1000000, gageinit[i] / 1000.0, gageres[i], gagepois[i], gagewire[i] / 1000.0, AIStrainUnits.Strain));
                        if (isFeed[i])
                        {
                            heaterFeedList.Items.Add(aIChannels[i].VirtualName);
                            feedChannel[feeder] = i;
                            feeder++;
                        }
                        else
                        {
                            gageChan.Items.Add(aIChannels[i].VirtualName);
                        }
                    }
                    btn9 = 1;
                }
                catch (DaqException exception)
                {
                    timer2.Stop();
                    MessageBox.Show(exception.Message);
                    runningTask = null;
                    myTask.Dispose();
                    if (isFeed.Contains(true))
                    {
                        heaterFeedList.Items.Clear();
                    }
                    aIChannels.Clear();
                    btn9 = 0;
                }
            }
            if (btn9 == 1)
            {
                myTask.Timing.ConfigureSampleClock("", Convert.ToDouble(rateNumeric.Value),
                    SampleClockActiveEdge.Rising, SampleQuantityMode.ContinuousSamples, 1000);
                try
                {
                    myTask.Control(TaskAction.Verify);
                    if (isFeed.Contains(false))
                    {
                        string channelsofexp = "";
                        for (int i = 0; i < chtype.Count; i++)
                        {
                            channelsofexp += aIChannels[i].VirtualName + Environment.NewLine;
                        }
                        MessageBox.Show($"{chtype.Count} Channels Ready for Experiment! :" + Environment.NewLine +
                            channelsofexp);
                    }
                    if (isFeed.Contains(true))
                    {
                        string channelsofheat = "";
                        for (int i = 0; i < heaterFeedList.Items.Count; i++)
                        {
                            channelsofheat += heaterFeedList.Items[i]+ Environment.NewLine;
                        }
                        MessageBox.Show($"{heaterFeedList.Items.Count} Channels Ready for Feedback! :" + Environment.NewLine +
                            channelsofheat);
                    }
                    myTask.Dispose();
                    runningTask = null;
                }
                catch (DaqException exception)
                {
                    MessageBox.Show(exception.Message);
                    runningTask = null;
                    myTask.Dispose();
                    if (isFeed.Contains(true))
                    {
                        heaterFeedList.Items.Clear();
                    }
                    aIChannels.Clear();
                }
            }
            else
            {
                runningTask = null;
                myTask.Dispose();
                if (isFeed.Contains(true))
                {
                    heaterFeedList.Items.Clear();
                }
                aIChannels.Clear();
            }
            if (chtype.Count == 0)
            {
                MessageBox.Show("There is not any channel set to be used in experiment!", "DAQmx Channel Settings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        int exptab = 0;
        int caltab = 0;
        int pretab = 0;
        int savecount = 0;
        int cursorcount = 0;
        void CreateGraphPage(string name, int graph)
        {
            if ((deney == 1 && graphexp == 0) || (deney == 3 && graphpre == 0) || (deney == 2 && graphcal == 0))
            {
                TabPage tab1 = new TabPage(name);
                tabControl1.TabPages.Add(tab1);
                tabControl1.SelectTab(tabControl1.TabCount - 1);
                if (deney == 1)
                {
                    graphexp = 1;
                    exptab = tabControl1.SelectedIndex;
                }
                else if (deney == 2)
                {
                    graphcal = 1;
                    caltab = tabControl1.SelectedIndex;
                }
                else if (deney == 3)
                {
                    graphpre = 1;
                    pretab = tabControl1.SelectedIndex;
                }
                Button buton = new Button();
                Button buton2 = new Button();
                Button buton3 = new Button();
                Button buton4 = new Button();
                CheckBox checkBoxg = new CheckBox();
                chart1.Location = new Point(5, -5);
                chart1.Size = new Size(750, 410);
                chart1.ChartAreas[0].AlignmentStyle = AreaAlignmentStyles.All;
                chart1.ChartAreas[0].CursorY.AxisType = AxisType.Primary;
                chart1.ChartAreas[0].CursorX.AxisType = AxisType.Primary;
                chart1.ChartAreas[0].CursorY.IsUserEnabled = true;
                chart1.ChartAreas[0].CursorX.IsUserEnabled = true;
                chart1.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
                chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
                chart1.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Dot;
                chart1.ChartAreas[0].AxisX2.MinorTickMark.LineDashStyle = ChartDashStyle.Dot;
                chart1.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Dot;
                chart1.ChartAreas[0].AxisY2.MinorTickMark.LineDashStyle = ChartDashStyle.Dot;
                chart1.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
                chart1.ChartAreas[0].AxisX2.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
                chart1.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
                chart1.ChartAreas[0].AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
                chart1.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisX2.MinorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY2.MinorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.FixedCount;
                chart1.ChartAreas[0].AxisX2.IntervalAutoMode = IntervalAutoMode.FixedCount;
                chart1.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount;
                chart1.ChartAreas[0].AxisY2.IntervalAutoMode = IntervalAutoMode.FixedCount;
                chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Number;
                chart1.ChartAreas[0].AxisX2.IntervalType = DateTimeIntervalType.Number;
                chart1.ChartAreas[0].AxisY.IntervalType = DateTimeIntervalType.Number;
                chart1.ChartAreas[0].AxisY2.IntervalType = DateTimeIntervalType.Number;
                chart1.ChartAreas[0].AxisX.IsLabelAutoFit = true;
                chart1.ChartAreas[0].AxisX2.IsLabelAutoFit = true;
                chart1.ChartAreas[0].AxisY.IsLabelAutoFit = true;
                chart1.ChartAreas[0].AxisY2.IsLabelAutoFit = true;
                chart1.Legends[0].LegendStyle = LegendStyle.Row;
                chart1.Legends[0].TableStyle = LegendTableStyle.Wide;
                chart1.Legends[0].TextWrapThreshold = 15;
                chart1.Legends[0].Docking = Docking.Top;
                chart1.Legends[0].IsDockedInsideChartArea = false;
                buton.Text = "Save Graph";
                buton.Location = new Point(100, 405);
                buton2.Location = new Point(180, 405);
                buton3.Location = new Point(260, 405);
                buton4.Location = new Point(340, 405);
                checkBoxg.Location = new Point(420, 405);
                checkBoxg.Text = "Show Position";
                buton2.Text = "Clear Graph";
                buton3.Text = "Stop Plotting";
                buton4.Text = "Temperature Cursor";
                chart1.Visible = true;
                tab1.Controls.Add(chart1);
                tab1.Controls.Add(buton);
                tab1.Controls.Add(buton2);
                tab1.Controls.Add(buton3);
                tab1.Controls.Add(buton4);
                tab1.Controls.Add(checkBoxg);
                checkBoxg.CheckedChanged += (sender, e) =>
                {
                    savecount++;
                    if (checkBoxg.Checked)
                    {
                        showpst = 1;
                    }
                    else
                    {
                        showpst = 0;
                    }
                };
                if (showpst == 1)
                {
                    checkBoxg.Checked = true;
                }
                buton2.Click += async (sender, e) =>
                  {
                      savecount++;
                      if (timer2.Enabled == true)
                      {
                          timer2.Stop();
                          await Task.Delay(timer2.Interval);
                      }
                      tabControl1.Cursor = Cursors.WaitCursor;
                      int say = chart1.Series.Count - 1;
                      for (int i = say; i > -1; i--)
                      {
                          chart1.Series[i].Points.Clear();
                          chartptcnt++;
                      }
                      if (deney == -1)
                      {
                          chart1.Series.Clear();
                          chartptcnt = 0;
                      }
                      graphcount = 0;
                      tabControl1.Cursor = DefaultCursor;
                      if (plotstop == 0 && deney!=-1)
                      {
                          timer2.Start();
                      }
                  };
                buton.Click += (sender, e) =>
                {
                    try
                    {
                        chart1.SaveImage(fileopen + "Graph" + savecount
                            + ".png", ChartImageFormat.Png);
                        tabControl1.Cursor = Cursors.WaitCursor;
                        if (savecount == 0)
                        {
                            for (int i = 0; i < chart1.Series.Count(); i++)///dataya sayac koyalım tekrar kayıt olmasın!
                            {
                                var datam = chart1.Series[i].Points.ToArray();
                                for (int j = 0; j < datam.Count(); j++)
                                {
                                    File.AppendAllText(fileopen + "Graph" + savecount +"_"+ chart1.Series[i].Name +
                                        ".txt", Convert.ToString(datam[j].XValue) +
                                        " " + Convert.ToString(datam[j].YValues[0]) + Environment.NewLine);
                                }
                            }
                        }
                        savecount++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    tabControl1.Cursor = DefaultCursor;

                };
                buton3.Click += (sender, e) =>
                {
                    savecount++;
                    if (plotstop == 0)
                    {
                        plotstop = 1;
                        buton3.Text = "Start Plotting";
                    }
                    else
                    {
                        plotstop = 0;
                        buton3.Text = "Stop Plotting";
                    }
                };
                buton4.Click += (ssender, e) =>
                {
                    chart1.Enabled = true;
                    if (cursorcount % 2 == 0)
                    {
                        chart1.ChartAreas[0].CursorY.AxisType = AxisType.Secondary;
                        buton4.Text = "Voltage";
                    }
                    else
                    {
                        chart1.ChartAreas[0].CursorY.AxisType = AxisType.Primary;
                        buton4.Text = "Temperature";
                    }
                    cursorcount++;
                };
                sizeadd = 1;
            }
            else if (deney == 1 && graphexp == 1)
            {
                tabControl1.SelectedIndex = exptab;
                chart1.Parent = tabControl1.TabPages[exptab];
                chart1.Visible = true;
            }
            else if (deney == 2 && graphcal == 1)
            {
                tabControl1.SelectedIndex = caltab;
                chart1.Parent = tabControl1.TabPages[caltab];
                chart1.Visible = true;
            }
            else if (deney == 3 && graphpre == 1)
            {
                tabControl1.SelectedIndex = pretab;
                chart1.Parent = tabControl1.TabPages[pretab];
                chart1.Visible = true;
            }
            if (deney == 3&&plotstop==0)
            {
                timer2.Start();
            }
        }
        int plotstop = 0;
        private async Task AutoConAsync(double threshold, double thresholdt, double thresholdg, bool direction)
        {
            double appsped = Convert.ToDouble(appSpeed.Text);
            double retractsped= 0;
            int avertime = 10;
            if (rateNumeric.Value < 10000)
            {
                avertime = Convert.ToInt32(1000 / rateNumeric.Value * 11);
            }
            temperChan = tempChan.SelectedIndex-1;
            double gradient = 0;
            double tempafter = 0;
            double tempnow = 0;
            int num = 0;
            int say = 0;
            double position = Convert.ToDouble(textBox1.Text);
            int stepno = Convert.ToInt16((zrange-zpos) * 100 / appsped);//*100= 0.1 s secure time interval * 1000(um to nm)--appsped in nm/s
            double voltageinc = (zrange-zpos) / (20 * kp * stepno);
            if (approaching == 1||approaching==3)
            {
                if (direction)
                {
                    Tdms_Saver("Approach");
                    plotstop = 1;
                    await Task.Delay(timer2.Interval);
                    timer3.Start();
                    timer2.Start();
                }
                say = 0;
            }
            if ((autoApp.Checked&&(useGagePress.Checked&&data!=null))||approaching==3||approaching==2)
            {
                double loadnow = 0;
                double sendcont = 0;
                plotstop = 0;
                if (direction)
                {
                    connection = "Approach Started!";
                    if (approaching != 3)
                    {
                        if (motorApp.Checked)
                        {
                            Send($"301{spedmode}0", "");
                            motsteps = 101;
                            if (approaching == 2)
                            {
                                motfreq = 1000000 * (stepinc / motspeed);
                            }
                            else
                            {
                                motfreq = 1000000 * (stepinc / appsped);
                            }
                            var motorsend = string.Format("{1:D10}{0:D10}", motsteps, Convert.ToInt32(motfreq));
                            texts = motorsend;
                            motorpass = 1;
                            motdir = 0;
                            await Task.Delay(100);
                            while (autopass == 0)
                            {
                                await Task.Delay(50);
                            }
                            autopass = 0;
                            if (approaching == 1)
                            {
                                loadnow = loadd.Last();
                                expfin = 0;
                                while (loadnow <= threshold && motpos + stepinc < motmax && emCounter % 2 != 0)
                                {
                                    try
                                    {
                                        motorPos.Text = Convert.ToString(motpos / 1000000);
                                    }
                                    catch (Exception ex)
                                    {
                                        connection = ex.Message;
                                    }
                                    await Task.Delay(10);
                                    loadnow = loadd.Last();
                                }
                                Properties.Settings.Default["motpos"] = motorPos.Text;
                                Properties.Settings.Default.Save();
                                ExperimentRangeController(depth, amplitude, retStep);
                                if (runningTask != null)
                                {
                                    timer2.Stop();
                                    runningTask = null;
                                    myTask.Dispose();
                                }
                                if (!doit)
                                {
                                    Send("Z0000", "");
                                    if (emCounter % 2 != 0)
                                    {
                                        MessageBox.Show("Z Range is not enough to make indent!",
                                       "Experiment Range Cotroller", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    emCounter = 0;
                                }
                                else
                                {
                                    Send("A0000", "");
                                    DialogResult result = DialogResult.Yes;
                                    if (loadnow < threshold)
                                    {
                                        result = MessageBox.Show("Surface could not be found by pressure set value! Start to Experiment?", "Auto Approach", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    }
                                    else
                                    {
                                        result = MessageBox.Show("Surface has been found! Start to Experiment?", "Auto Approach", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    }
                                    if (result == DialogResult.No)
                                    {
                                        emCounter = 0;
                                    }
                                    while (expfin == 0 && emCounter % 2 != 0)
                                    {
                                        await Task.Delay(20);
                                    }
                                    if (emCounter % 2 == 0)
                                    {
                                        await Task.Delay(500);
                                        Send("Z0000", "");
                                    }
                                }
                                approaching = 0;
                            }
                            else
                            {
                                while (!comread.Contains("CLOSe") && emCounter % 2 != 0)
                                {
                                    await Task.Delay(avertime);
                                }
                                if (emCounter % 2 != 0)
                                {
                                    MessageBox.Show("Landed to surface succesfully!" + Environment.NewLine +
          " Use auto approach for more controlled landing.", "Auto Land", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    micstepBox.SelectedIndex = 3;
                                    approaching = 0;
                                    try
                                    {
                                        abq = comread.Replace("CLOSe", "");
                                        abqi = Convert.ToInt16(abq);
                                        motpos = motpos + abqi * stepinc / 2;
                                        motorPos.Text = Convert.ToString(motpos / 1000000);
                                        // verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                                        comread = "";
                                        abqi = 0;
                                    }
                                    catch (Exception ex)
                                    {
                                        label11.Text = ex.Message;
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        motorPos.Text = Convert.ToString(motpos / 1000000);
                                        // verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                                    }
                                    catch (Exception ex)
                                    {
                                        label11.Text = ex.Message;
                                        //connection = comread;
                                    }
                                }
                                if (emCounter % 2 != 0)
                                {
                                    Send("STFIN", "");
                                }
                                else
                                {
                                    Send("Z0000", "");
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < stepno; i++)
                            {
                                if (emCounter % 2 != 0)
                                {
                                    loadnow = loadd.Last();
                                    if (loadnow <= threshold)
                                    {

                                        sendcont = position - (voltageinc * (i + 1));
                                        Send(TextFormat(sendcont.ToString(), "Vapp", -1, 7.5), "W");
                                        textBox1.Text = TextFormat(sendcont.ToString(), "Vapp", -1, 7.5);
                                        while (expresscon != 1 && emCounter % 2 != 0)
                                        {
                                            await Task.Delay(20);
                                            num++;
                                        }
                                        if (num < 5)
                                        {
                                            await Task.Delay(20 * (5 - num));
                                        }
                                        num = 0;
                                        expresscon = 0;
                                    }
                                    if (loadnow >= threshold)
                                    {
                                        if (runningTask != null)
                                        {
                                            timer2.Stop();
                                            runningTask = null;
                                            myTask.Dispose();
                                        }
                                        ExperimentRangeController(depth, amplitude, retStep);
                                        if (!doit)
                                        {
                                            Send("Z0000", "");
                                            if (emCounter % 2 != 0)
                                            {
                                                MessageBox.Show("Z Range is not enough to make indent!",
                                               "Experiment Range Cotroller", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            }
                                            emCounter = 0;
                                            break;
                                        }
                                        else
                                        {
                                            DialogResult result = MessageBox.Show("Surface has been found! Start to Experiment?", "Auto Approach", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                            if (result == DialogResult.No)
                                            {
                                                emCounter = 0;
                                                Send("Z0000", "");
                                            }
                                            else
                                            {
                                                Send("A0000", "");
                                                expfin = 0;
                                                while (expfin == 0 && emCounter % 2 != 0)
                                                {
                                                    await Task.Delay(10);
                                                }
                                                if (emCounter % 2 == 0)
                                                {
                                                    await Task.Delay(500);
                                                    Send("Z0000", "");
                                                    emCounter = 0;
                                                }
                                            }
                                        }
                                        break;
                                    }
                                    if (sendcont - voltageinc < -1)
                                    {
                                        MessageBox.Show("Surface could not be found by pressure set value!", "Auto Approach");
                                        Send("Z0000", "");
                                        emCounter = 0;
                                        break;
                                    }
                                }
                                else
                                {
                                    Send("Z0000", "");
                                    break;
                                }
                            }
                        }
                    }
                    else//internal mcu control
                    {
                        expfin = 0;
                        while (expfin == 0 && emCounter % 2 != 0)
                        {
                            try
                            {
                                motorPos.Text = Convert.ToString(motpos / 1000000);
                                textBox1.Text = receive;
                            }
                            catch (Exception ex)
                            {
                                connection = ex.Message;
                            }
                            await Task.Delay(25);
                        }
                        Properties.Settings.Default["motpos"] = motorPos.Text;
                        Properties.Settings.Default.Save();
                        ExperimentRangeController(depth, amplitude,retStep);
                        if (emCounter % 2 == 0||!doit)
                        {
                            emCounter = 0;
                            await Task.Delay(100);
                            Send("Z0000", "");
                        }
                        approaching = 0;
                    }
                }
                else
                {
                    connection = "Retraction Started!";
                    if (loadd.Count != 0)
                    {
                        holdposition = threshold+(maxload-threshold) * (100 - Convert.ToDouble(holdPercent.Value))/100;
                    }
                    stepno = Convert.ToInt16((Convert.ToDouble(initialpos) - position) * 20 * kp * 0.1 / retractsped);
                    voltageinc = (Convert.ToDouble(initialpos) / 20 - position) / stepno;
                    if (approaching==2)
                    { 
                        Send($"311{spedmode}0", "");
                        motsteps = motorTrack.Value;
                        if (approaching == 2)
                        {
                            motfreq = 1000000*(stepinc / motspeed);
                        }
                        else
                        {
                            motfreq = 1000000 * (stepinc / retractsped);
                        }
                        var motorsend = string.Format("{1:D10}{0:D10}", motsteps, Convert.ToInt32(motfreq));
                        texts = motorsend;
                        motorpass = 1;
                        motdir = 0;
                        while (autopass == 0 && emCounter%2 != 0)
                        {
                            await Task.Delay(10);
                        }
                        autopass = 0;
                        if (approaching==1)
                        {
                            say = 0;
                            loadnow = loadd.Last();
                            while (loadnow > 0 && motpos - stepinc > 0 && emCounter % 2 != 0)
                            {
                                if (comread.Contains("UPMOT"))
                                {
                                    motorPos.Text = "0.0";
                                    motpos = 0;
                                    //verticalProgressBar2.Value = 100;
                                    break;
                                }
                                try
                                {
                                    motorPos.Text = Convert.ToString(motpos / 1000000);
                                    //verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                                }
                                catch (Exception ex)
                                {
                                    label11.Text = ex.Message;
                                    connection = comread;
                                }
                                loadnow = loadd.Last();
                                if (loadnow <= holdposition)
                                    {
                                        Send("A1000", "");
                                        say = 0;
                                        await Task.Delay(holdtime);
                                        connection = "Holding at given position while retraction.";
                                        expfin = 0;
                                        while (expfin==0 && emCounter % 2 != 0)
                                        {
                                            await Task.Delay(10);
                                            //say++;
                                            //if (say == 10)
                                            //{
                                            //    say = 0;
                                            //    Send("A1000", "");
                                            //}
                                        }
                                        say = 1;
                                        await Task.Delay(holdtime);
                                        break;
                                    }
                                    connection = "Retracting!";
                                await Task.Delay(100);
                            }
                                Send("A1000", "");
                                say = 0;
                                while (!comread.Contains("Complete") && emCounter % 2 != 0)
                                {
                                    await Task.Delay(10);
                                    say++;
                                    if (say == 10)
                                    {
                                        Send("A1000", "");
                                        say = 0;
                                    }
                                }
                                say = 0;
                                if (loadnow > threshold)
                                {
                                    MessageBox.Show("Retraction set value can not be reached!", "Auto Approach");
                                }
                                else
                                {
                                    MessageBox.Show("Retraction is completed.", "Auto Approach");
                                }
                            if(!comread.Contains("UPMOT")&&say==1)
                            {
                                await Task.Delay(100);
                                Send($"311{spedmode}0", "");
                                motsteps = motorTrack.Value;
                                motfreq = 1000000 * (stepinc / retractsped);
                                motorsend = string.Format("{1:D10}{0:D10}", motsteps, Convert.ToInt32(motfreq));
                                texts = motorsend;
                                motorpass = 1;
                                motdir = 0;
                                while (autopass == 0 && emCounter % 2 != 0)
                                {
                                    await Task.Delay(10);
                                }
                                autopass = 0;
                                loadnow = loadd.Last();
                                    while (loadnow > 0 && motpos - stepinc > 0 && emCounter % 2 != 0)
                                    {
                                        if (comread.Contains("UPMOT"))
                                        {
                                            motorPos.Text = "0.0";
                                            motpos = 0;
                                            //verticalProgressBar2.Value = 100;
                                            break;
                                        }
                                        try
                                        {
                                            motorPos.Text = Convert.ToString(motpos / 1000000);
                                            // verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                                            comread = "";
                                        }
                                        catch (Exception ex)
                                        {
                                            connection = ex.Message;
                                        }
                                        loadnow = loadd.Last();
                                        await Task.Delay(10);
                                    }
                                    Send("A1000", "");
                                    say = 0;
                                    expfin = 0;
                                    while (expfin==0&& emCounter % 2 != 0)
                                    {
                                        await Task.Delay(10);
                                        //say++;
                                        //if (say == 10)
                                        //{
                                        //    Send("A1000", "");
                                        //    say = 0;
                                        //}
                                    }
                                    say = 0;
                                    if (loadnow > threshold)
                                    {
                                        MessageBox.Show("Retraction set value can not be reached!", "Auto Approach");
                                    }
                                    else
                                    {
                                        MessageBox.Show("Retraction is completed.", "Auto Approach");
                                    }
                                    while (!comread.Contains("UPMOT") && emCounter % 2 != 0)
                                    {
                                        try
                                        {
                                            motorPos.Text = Convert.ToString(motpos / 1000000);
                                            // verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                                        }
                                        catch (Exception ex)
                                        {
                                             connection= ex.Message;
                                        }
                                        await Task.Delay(10);
                                    }
                                    motpos = 0;
                                    Send("A1000", "");
                                    motorPos.Text = "0.0";
                                    expfin = 0;
                                    while (expfin==0&& emCounter % 2 != 0)
                                    {
                                        await Task.Delay(10);
                                        say++;
                                        if (say == 10)
                                        {
                                            Send("A1000", "");
                                            say = 0;
                                        }
                                    }
                            }
                            else
                            {
                                Send("A1000", "");
                                if (loadnow > threshold)
                                {
                                    MessageBox.Show("Set value can not be reached!", "Auto Approach");
                                }
                                else
                                {
                                    MessageBox.Show("Retraction is completed.", "Auto Approach");
                                }
                                expfin = 0;
                                while (expfin==0 && emCounter % 2 != 0)
                                {
                                    await Task.Delay(10);
                                    say++;
                                    if (say == 10)
                                    {
                                        Send("A1000", "");
                                        say = 0;
                                    }
                                }
                            }
                            
                        }
                        else
                        {
                            while (!comread.Contains("UPMOT") && emCounter % 2 != 0)
                            {
                                await Task.Delay(20);
                            }
                            if (emCounter % 2 != 0)
                            {
                                connection = "";
                                MessageBox.Show("Motor at Home Position!", "Auto Land", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                motorPos.Text = "0.0";
                                motpos = 0;
                                Send("STFIN", "");
                                //verticalProgressBar2.Value = 100;
                            }
                            else
                            {
                                Send("Z0000", "");
                                try
                                {
                                    motorPos.Text = Convert.ToString(motpos / 1000000);
                                    // verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                                }
                                catch (Exception ex)
                                {
                                    connection = ex.Message;
                                    //connection = comread;
                                }
                            }
                            
                        }
                        approaching = 0;
                    }
                    else
                    {
                        for (int i = 0; i < stepno; i++)
                        {
                            if (emCounter % 2 != 0)
                            {
                                say = 0;
                                loadnow = 0;
                                for (int j = 9; j > -1; j--)
                                {
                                    if (loadd.Count > j)
                                    {
                                        loadnow = loadnow + loadd[loadd.Count - 1 - j];
                                        say++;
                                    }
                                }
                                loadnow = loadnow / say;
                                say = 0;
                                if (loadnow > threshold)
                                {
                                    if (loadnow <= holdposition)
                                    {
                                        connection = "Holding at given position while retraction.";
                                        await Task.Delay(holdtime);
                                    }
                                    connection = "Retracting!";
                                    sendcont = position + (voltageinc * (i + 1));
                                    Send(TextFormat(sendcont.ToString(), "Vapp", -1, 7.5), "W");
                                    textBox1.Text = TextFormat(sendcont.ToString(), "Vapp", -1, 7.5);
                                    while (expresscon != 1 && emCounter % 2 != 0)
                                    {
                                        await Task.Delay(20);
                                        say++;
                                    }
                                    if (say < 5)
                                    {
                                        await Task.Delay(20 * (5 - say));
                                    }
                                    expresscon = 0;
                                }
                                if (loadnow <= threshold)
                                {
                                    break;
                                }
                                if (sendcont + voltageinc >= 7.5)
                                {
                                    break;
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        Send("A1000", "");
                        expfin = 0;
                        while (expfin == 0 && emCounter % 2 != 0)
                        {
                            await Task.Delay(10);
                            say++;
                            if (say == 10)
                            {
                                Send("A1000", "");
                                say = 0;
                            }
                        }
                    }
                    approaching = 0;
                }
            }
            groupBox19.Enabled = true;
        }
        private async Task InitialCommandsAsync(DialogResult result)
        {
            initialize = 0;
            if (result == DialogResult.OK)
            {
                Send(Properties.Settings.Default["Vapp"].ToString(), "W");
                while (directpass == 1)
                {
                    await Task.Delay(25);
                }
                await Task.Delay(50);
                Send(Properties.Settings.Default["Kp"].ToString(), "P");
                while (directpass == 1)
                {
                    await Task.Delay(25);
                }
                await Task.Delay(50);
                Send(Properties.Settings.Default["Ki"].ToString(), "I");
                while (directpass == 1)
                {
                    await Task.Delay(25);
                }
                await Task.Delay(50);
                Send(Properties.Settings.Default["Kd"].ToString(), "D");
                while (directpass == 1)
                {
                    await Task.Delay(25);
                }
                await Task.Delay(50);
                Send(Properties.Settings.Default["filterType"].ToString(), "C");
                while (directpass == 1)
                {
                    await Task.Delay(25);
                }
                await Task.Delay(50);
                Send(Properties.Settings.Default["cutOff1"].ToString(), "F");
                while (directpass == 1)
                {
                    await Task.Delay(25);
                }
                await Task.Delay(50);
                Send(Properties.Settings.Default["cutOff2"].ToString(), "S");
                while (directpass == 1)
                {
                    await Task.Delay(25);
                }
                await Task.Delay(50);
                Send(Properties.Settings.Default["Vmin"].ToString(), "N");
                while (directpass == 1)
                {
                    await Task.Delay(25);
                }
                await Task.Delay(50);
                Send(Properties.Settings.Default["Vmax"].ToString(), "M");
                while (directpass == 1)
                {
                    await Task.Delay(25);
                }
                await Task.Delay(50);
                if (Properties.Settings.Default["inverse"].ToString() == "0")
                {
                    Send(Properties.Settings.Default["deltaS"].ToString(), "G");
                }
                else
                {
                    Send(Properties.Settings.Default["deltaS"].ToString(), "G-");
                }
                await Task.Delay(500);
            }
            else
            {
            }
        }

        private void pictureBox1_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(pictureBox1, "Emergency (Electrical Protection): Applied Voltage will be set to zero immediately after click");
        }

        private void pictureBox2_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(pictureBox2, "Emergency (Tip&Surface Protection): Applied Voltage will be set to maximum (150V) immediately after click");
        }

        private void StopExp_Click(object sender, EventArgs e)
        {
            stopExp.Enabled = false;
            groupBox5.Enabled = true;
            executeExp.Enabled = false;
            emCounter = 0;
            texts = "";
            textexp.Clear();
            Send("Z0000", "");
        }

        private void button7_Click_2(object sender, EventArgs e)
        {

        }
        int toolax = 0;
        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
        }

        private void chart1_MouseWheel(object sender, MouseEventArgs e)
        {
            double locx = e.Location.X;
            double locy = e.Location.Y;
            chart1 = (Chart)sender;
            var xAxis = chart1.ChartAreas[0].AxisX;
            var yAxis = chart1.ChartAreas[0].AxisY;
            var xAxis2 = chart1.ChartAreas[0].AxisX2;
            var yAxis2 = chart1.ChartAreas[0].AxisY2;

            try
            {
                if (e.Delta < 0) // Scrolled down.
                {
                    xAxis.ScaleView.ZoomReset();
                    yAxis.ScaleView.ZoomReset();
                    //xAxis2.ScaleView.ZoomReset();
                    yAxis2.ScaleView.ZoomReset();
                }
                else if (e.Delta > 0) // Scrolled up.
                {
                    var xMin = xAxis.ScaleView.ViewMinimum;
                    var xMax = xAxis.ScaleView.ViewMaximum;
                    var yMin = yAxis.ScaleView.ViewMinimum;
                    var yMax = yAxis.ScaleView.ViewMaximum;
                    //var xMin2 = xAxis2.ScaleView.ViewMinimum;
                    //var xMax2 = xAxis2.ScaleView.ViewMaximum;
                    var yMin2 = yAxis2.ScaleView.ViewMinimum;
                    var yMax2 = yAxis2.ScaleView.ViewMaximum;
                    var posXStart = xAxis.PixelPositionToValue(locx) - (xMax - xMin) / 2;
                    var posXFinish = xAxis.PixelPositionToValue(locx) + (xMax - xMin) / 2;
                    var posYStart = yAxis.PixelPositionToValue(locy) - (yMax - yMin) / 2;
                    var posYFinish = yAxis.PixelPositionToValue(locy) + (yMax - yMin) / 2;
                    //var posXStart2 = xAxis2.PixelPositionToValue(locx) - (xMax2 - xMin2) / 2;
                    //var posXFinish2 = xAxis2.PixelPositionToValue(locx) + (xMax2 - xMin2) / 2;
                    var posYStart2 = yAxis2.PixelPositionToValue(locy) - (yMax2 - yMin2) / 2;
                    var posYFinish2 = yAxis2.PixelPositionToValue(locy) + (yMax2 - yMin2) / 2;
                    xAxis.ScaleView.Zoom(posXStart, posXFinish);
                    yAxis.ScaleView.Zoom(posYStart, posYFinish);
                    //xAxis2.ScaleView.Zoom(posXStart2, posXFinish2);
                    yAxis2.ScaleView.Zoom(posYStart2, posYFinish2);
                }
            }
            catch (Exception ex)
            {
                label11.Text = ex.Message;
            }
        }

        int sizeadd = 0;
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (sizeadd == 0)
            {
                if (tabControl1.SelectedIndex == 2)
                {
                    this.Width = 420 + sizeadd;
                }
                else if (tabControl1.SelectedIndex == 1 && defCount % 2 == 0)
                {
                    this.Width = 370 + sizeadd;
                }
                else if (tabControl1.SelectedIndex == 4 && serialAct.Checked)
                {
                    this.Width = 560;
                }
                else if (tabControl1.SelectedIndex == 4 && !serialAct.Checked)
                {
                    this.Width = 420 + sizeadd;
                }
                else if (tabControl1.SelectedIndex == 3 && chartprop == 0)
                {
                    this.Width = 420 + sizeadd;
                }
                else if (tabControl1.SelectedIndex == 0)
                {
                    this.Width = 769 + sizeadd;
                }
                else
                {
                    this.Width = 769;
                }
            }
            else
            {
                this.Width = 769;
            }

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (serialPort1.IsOpen)
                serialPort1.Close();
            if (serialPort2.IsOpen)
                serialPort2.Close();
            if (myTask != null)
            {
                runningTask = null;
                myTask.Dispose();
            }
        }
        int znull = 0;
        private void button10_Click(object sender, EventArgs e)
        {
            if (znull % 2 == 0)
            {
                vnull = Convert.ToDouble(vapplied);
                zrange = kp * 20*(vnull - vmin);
                button10.Text = "Null";
                textBox10.Text = TextFormat(((vnull - vol) *20* kp).ToString(), "Z", 0, (vnull - vmin) *20* kp);
                znull++;
            }
            else
            {
                button10.Text = "Set Max";
                vnull = vmax;
                zrange = kp *20* (vnull - vmin);
                textBox10.Text = TextFormat(((vnull - vol) *20* kp).ToString(), "Z", (vmin - (vmax - vnull) + 1) *20* kp, (vnull - vmin)*20 * kp);
                znull++;
            }

        }
        bool doit = false;
        private void ExperimentRangeController(List<string> indent, List<string> amp, List<int> retract)
        {
            zexp = 0;
            if (autoApp.Checked)
            {
                zexp = 10;//10um range for search
            }
            if (indent.Count != 0 && (exptype==0||exptype==2))
            {
                for (int i = 0; i < indent.Count; i++)
                {
                    zexp = zexp + Convert.ToDouble(indent[i]);
                    if (zexp + Convert.ToDouble(amp[i]) > zrange - zpos&&!motorDrive.Checked)
                    {
                        doit = false;
                        break;
                    }
                    else if (zexp + Convert.ToDouble(amp[i]) >motmax&& motorDrive.Checked)
                    {
                        doit = false;
                        break;
                    }
                    if ((retract[i] == 1))
                    {
                        zexp = 10;
                    }
                    else if (indent.Count-1 > i)
                    {
                        if(Convert.ToDouble(Xpos[i + 1]) != 0 || Convert.ToDouble(Xpos[i + 1]) != 0)
                        {
                            zexp = 10;
                        }
                    }
                    doit = true;
                }
            }
            else if (indent.Count != 0 && exptype == 1)
            {
                doit = true;
            }
            if(emCounter %2 == 0)
            {
                doit = false;
            }
        }

        double pressure = 0;
        private void textBox25_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                pressure = Convert.ToDouble(loadThres.Text);
                Properties.Settings.Default["PressThres"] = loadThres.Text;
                Properties.Settings.Default.Save();
            }
        }
        private void button11_Click(object sender, EventArgs e)
        {
            string zaman = @"\" + DateTime.Today.ToString("dd-MM-yyyy");
            folderBrowserDialog1.ShowDialog();
            pathsave = folderBrowserDialog1.SelectedPath;
            Directory.CreateDirectory(pathsave + zaman);
            pathsave = pathsave + zaman;
            Properties.Settings.Default["PathtoSave"] = pathsave;
            Properties.Settings.Default.Save();
        }
        public string pathtemp = "";
        private void directorset(string now)
        {
            pathtemp = pathsave + now;
            fileopen = pathsave + now;
            fileopentdms = pathsave + now + "\\NI_tdms.tdms";
            Directory.CreateDirectory(pathsave + now);
        }
        string pathsave = Application.StartupPath + DateTime.Now.ToString("dd-MM-yyyy");

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }
        int timer4i = 0;
        List<string> tb1 = new List<string>();
        List<string> tb10 = new List<string>();
        private void timer4_Tick(object sender, EventArgs e)
        {
            try
            {
                if (saveExt.Checked)
                {
                    extdata = Convert.ToDouble(response);
                    extResp.Text = response;
                }
                else
                {
                    extResp.Text = response;
                }
            }
            catch (Exception ex)
            {
                label11.Text = ex.Message;
                extResp.Text = response;
            }
        }
        double tempthres = 0;
        double gradthres = 0;
        List<double> dif = new List<double>();
        List<double> volcal1 = new List<double>();
        List<double> volcal2 = new List<double>();
        List<double> volvs1 = new List<double>();
        List<double> volvs2 = new List<double>();
        List<double> dv = new List<double>();
        private void PressureExtracter(List<double> experiment, List<double> cal, int start, int stop)
        {
            //double sum = 0;
            for (int i = start; i < stop; i++)
            {
                if (cal.Count - 1 > i)
                {
                    dv.Add(gageco * (cal[i] - experiment[i]));
                    gageShow.Text = (cal[i] - experiment[i]).ToString();
                }
                else
                {
                    dv.Add(0);
                }
            }
            /*for(int i=start; i < stop; i++)
            {
                sum = sum + dv[i];
            }
            sum = sum / (stop - start);
            textBox29.Text = sum.ToString();*/
        }
        bool cal = true;
        int timer5i = 0;
        private void timer5_Tick(object sender, EventArgs e)
        {
            if (!cal && data != null && voltagechannel != -1)
            {
                for (int i = 0; i < data[voltagechannel].SampleCount; i++)
                {
                    //volvs1.Add(Convert.ToDouble(tb1[timer4i - (10 - i)]) * 20);
                    volcal1.Add(data[voltagechannel].Samples[i].Value);
                }
                PressureExtracter(volcal1, volcal2, timer5i, timer5i + 10);
                timer5i = timer5i + 10;
            }
            if (cal && data != null && voltagechannel != -1)
            {
                for (int i = 0; i < data[voltagechannel].SampleCount; i++)
                {
                    //volvs2.Add(Convert.ToDouble(tb1[timer4i - (10 - i)]) * 20);
                    volcal2.Add(data[voltagechannel].Samples[i].Value);
                }
            }

        }
        double gageco = 1;
        double gageco0 = 0;
        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox13.Checked)
            {
                cal = true;
            }
            else
            {
                cal = false;
            }

        }
        int plug = 0;

        private void button12_Click(object sender, EventArgs e)
        {
            if (plug % 2 == 0)
            {
                nocon = 0;
                emCounter = 1;
                button12.Text = "Secure Unplug";
                label11.Text = "Searching Connection..";
                timer6.Start();
            }
            else
            {
                groupBox5.Enabled = false;
                groupBox8.Enabled = false;
                groupBox19.Enabled = false;
                timer6.Stop();
                isConnected = false;
                serialPort2.Close();
                label11.Text = "Ready for unplug!";
                button12.Text = "Secure Plug-in";
                //emergency call or closing chars
            }
            plug++;
        }
        double forcecoef = 1;
        string connection = "";
        double maxload = 0;
        double maxtemp = 0;
        double forceControl = 0;
        string forceSend = "";
        private void timer3_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!useGagePress.Checked)
                {
                    pressShow.Text = loadd.Last().ToString();
                    loadcellData.Text = loadd.Last().ToString();
                    loadExtraShow.Text = loadd.Last().ToString();
                }
                else if (useGagePress.Checked)
                {
                    if (saveExt.Checked && useGagePress.Checked)
                    {
                        pressShow.Text = extdata.ToString();
                        loadd.Add(extdata);
                    }
                    else if (data != null && voltagechannel != -1)
                    {
                        loadd.Add((data[voltagechannel].Samples.Last().Value *gageco+gageco0)*forcecoef);
                        //pressShow.Text = data[voltagechannel].Samples.Last().Value.ToString();
                    }
                    //pressShow.Text = (Convert.ToDouble(extraSG.Text) * forcecoef).ToString();
                    if (maxload < loadd.Last())
                    {
                        maxload = loadd.Last();
                    }
                    if (loadd.Last() < forceControl)
                    {
                        forceSend = "LP+++";
                    }
                    else if (loadd.Last() > forceControl)
                    {
                        forceSend = "LP---";
                    }
                    if (heat_com == 0 && autopass != 1 && feed != 1 && hxcom != 1 && motorpass != 1 && directpass != 1 && !sending
    && calpass != 1 && exppass == 0 && xymotor != 1 && rhcom != 1 && balancing != 1 && datacomp == 0)
                    {
                        Send(forceSend,"");
                    }
                }
            }
            catch (Exception ex)
            {
                label11.Text = ex.Message;

            }

        }

        private void maximumValueNumeric_ValueChanged(object sender, EventArgs e)
        {
            if (maximumValueNumeric.Value <= minimumValueNumeric.Value)
            {
                maximumValueNumeric.Value = minimumValueNumeric.Value + 1;
            }
        }

        private void minimumValueNumeric_ValueChanged(object sender, EventArgs e)
        {
            if (maximumValueNumeric.Value <= minimumValueNumeric.Value)
            {
                minimumValueNumeric.Value = maximumValueNumeric.Value - 1;
            }
        }

        double stpos = 0;
        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {
        }

        private void textBox11_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void textBox12_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void textBox13_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void textBox14_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void textBox15_MouseClick(object sender, MouseEventArgs e)
        {
        }

        private void textBox16_MouseClick(object sender, MouseEventArgs e)
        {
        }

        private void textBox20_MouseClick(object sender, MouseEventArgs e)
        {
        }

        private void textBox22_MouseClick(object sender, MouseEventArgs e)
        {
        }
        double divideresult = 0;

        private void VappNumeric_ValueChanged(object sender, EventArgs e)
        {
            if (clickinc != 0)
            {
                if (zCount % 2 != 0)
                {
                    divideresult = (Convert.ToDouble(VappNumeric.Value) / 1000.0);
                    textBox10.Text = TextFormat(divideresult.ToString(), "Z", 0, (vnull - vmin) *20* kp);
                }
                else
                {
                    divideresult = (Convert.ToDouble(VappNumeric.Value) / 1000000.0);
                    textBox1.Text = TextFormat(divideresult.ToString(), "Vapp", vmin, vnull);
                }
                labelVapp.Text = textBox1.Text;
                Send(labelVapp.Text, "W");
                labelZ.Text = textBox10.Text;
                Properties.Settings.Default["Vapp"] = textBox1.Text;
                Properties.Settings.Default["Z"] = textBox10.Text;
                Properties.Settings.Default.Save();
                //VappNumeric.Enabled = false;
                //verticalProgressBar1.Value = Convert.ToInt16(200 * (Convert.ToDouble(textBox1.Text) + 1) / 17);
            }
            clickinc = 1;
        }

        private void textBox32_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                Properties.Settings.Default["VoltInc"] = textBox32.Text;
                Properties.Settings.Default.Save();
                if (zCount % 2 == 0)
                {
                    VappNumeric.Increment = (decimal)(Convert.ToDouble(textBox32.Text) * 1000);
                }
            }
        }

        private void textBox33_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                Properties.Settings.Default["PosInc"] = textBox33.Text;
                Properties.Settings.Default.Save();
                if (zCount % 2 != 0)
                {
                    VappNumeric.Increment = (decimal)Convert.ToDouble(textBox33.Text);
                }
            }
        }
        int clickinc = 0;

        private void textBox10_MouseClick(object sender, MouseEventArgs e)
        {
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (!combochange)
            {
                if (!checkBox3.Checked)
                {
                    textBox14.Enabled = true;
                    textBox15.Enabled = true;
                    textBox16.Enabled = true;
                    checkBox11.Enabled = true;
                    checkBox5.Enabled = true;
                    checkBox3.Enabled = false;
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                }
                else
                {
                    textBox14.Enabled = false;
                    textBox15.Enabled = false;
                    textBox16.Enabled = false;
                    checkBox11.Enabled = false;
                    checkBox5.Enabled = false;
                    checkBox11.Checked = false;
                    checkBox5.Checked = false;
                    box3.Insert(comboBox2.SelectedIndex - 1, false);
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                }
                if (box3.Count > comboBox2.SelectedIndex)
                {
                    box3.RemoveAt(comboBox2.SelectedIndex);
                    box5.RemoveAt(comboBox2.SelectedIndex);
                    box11.RemoveAt(comboBox2.SelectedIndex);
                }
            }


        }

        private void textBox26_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                Properties.Settings.Default["AppSpeed"] = appSpeed.Text;
                Properties.Settings.Default.Save();
            }
        }
        int graphcount = 0;
        private void Tdms_Saver(string group)
        {
            if (btn9 == 1)
            {
                datas.RemoveRange(0, datas.Count());
                myTask = new NationalInstruments.DAQmx.Task();
                List<AIChannel> aIChannels = new List<AIChannel>();
                for (int i = 0; i < chtype.Count; i++)
                {
                    if (chtype[i] == 0)
                    {
                        aIChannels.Add(myTask.AIChannels.CreateVoltageChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
                            "Voltage" + phychannel[i].ToString(), (AITerminalConfiguration)(-1), minval[i], maxval[i], AIVoltageUnits.Volts));
                    }
                    else if (chtype[i] == 1)
                    {
                        AIThermocoupleType thermocoupleType;
                        AIAutoZeroMode autoZeroMode;

                        switch (tctype[i])
                        {
                            case 0:
                                thermocoupleType = AIThermocoupleType.B;
                                break;
                            case 1:
                                thermocoupleType = AIThermocoupleType.E;
                                break;
                            case 2:
                                thermocoupleType = AIThermocoupleType.J;
                                break;
                            case 3:
                                thermocoupleType = AIThermocoupleType.K;
                                break;
                            case 4:
                                thermocoupleType = AIThermocoupleType.N;
                                break;
                            case 5:
                                thermocoupleType = AIThermocoupleType.R;
                                break;
                            case 6:
                                thermocoupleType = AIThermocoupleType.S;
                                break;
                            default:
                                thermocoupleType = AIThermocoupleType.K;
                                break;
                        }

                        switch (cjctype[i])
                        {
                            case 0: // Channel
                                aIChannels.Add(myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
                                   "Temperature" + phychannel[i].ToString(), minval[i], maxval[i],
                                   thermocoupleType, AITemperatureUnits.DegreesC, cjcChannelTextBox.Text));
                                break;
                            case 1: // Constant
                                aIChannels.Add(myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
                                    "Temperature" + phychannel[i].ToString(), minval[i], maxval[i],
                                    thermocoupleType, AITemperatureUnits.DegreesC, Convert.ToDouble(cjcValueNumeric.Value)));
                                break;
                            case 2:
                                aIChannels.Add(myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
     "Temperature" + phychannel[i].ToString(), minval[i], maxval[i],
     thermocoupleType, AITemperatureUnits.DegreesC, Convert.ToDouble(cjcValueNumeric.Value)));
                                break;// Internal
                            default:
                                aIChannels.Add(myTask.AIChannels.CreateThermocoupleChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
                                    "Temperature" + phychannel[i].ToString(), minval[i], maxval[i],
                                    thermocoupleType, AITemperatureUnits.DegreesC));
                                break;
                        }
                        autoZeroMode = AIAutoZeroMode.None;
                        aIChannels[i].AutoZeroMode = autoZeroMode;
                    }
                    else if (chtype[i] == 2)//rtd temp measurement
                    {
                        AIRtdType aIRtd;
                        AIExcitationSource aIExc;
                        switch (rtdtype[i])
                        {
                            case 0:
                                aIRtd = AIRtdType.Pt3750;
                                break;
                            case 1:
                                aIRtd = AIRtdType.Pt3851;
                                break;
                            case 2:
                                aIRtd = AIRtdType.Pt3911;
                                break;
                            case 3:
                                aIRtd = AIRtdType.Pt3916;
                                break;
                            case 4:
                                aIRtd = AIRtdType.Pt3920;
                                break;
                            case 5:
                                aIRtd = AIRtdType.Pt3928;
                                break;
                            default:
                                aIRtd = AIRtdType.Pt3851;
                                break;
                        }
                        switch (rtdexctype[i])
                        {
                            case 0:
                                aIExc = AIExcitationSource.External;
                                break;
                            case 1:
                                aIExc = AIExcitationSource.Internal;
                                break;
                            default:
                                aIExc = AIExcitationSource.External;
                                break;
                        }
                        aIChannels.Add(myTask.AIChannels.CreateRtdChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
                                    "Temperature" + phychannel[i].ToString(), minval[i], maxval[i], AITemperatureUnits.DegreesC, aIRtd,
                                    AIResistanceConfiguration.TwoWire, aIExc, rtdexc[i], rtdres[i]));
                    }
                    else if (chtype[i] == 3)//strain gage measurement
                    {
                        AIStrainGageConfiguration aIStrain;
                        AIExcitationSource aIExcitation;
                        switch (gagetype[i])
                        {
                            case 0:
                                aIStrain = AIStrainGageConfiguration.FullBridgeI;
                                break;
                            case 1:
                                aIStrain = AIStrainGageConfiguration.FullBridgeII;
                                break;
                            case 2:
                                aIStrain = AIStrainGageConfiguration.FullBridgeIII;
                                break;
                            case 3:
                                aIStrain = AIStrainGageConfiguration.HalfBridgeI;
                                break;
                            case 4:
                                aIStrain = AIStrainGageConfiguration.HalfBridgeII;
                                break;
                            case 5:
                                aIStrain = AIStrainGageConfiguration.QuarterBridgeI;
                                break;
                            case 6:
                                aIStrain = AIStrainGageConfiguration.QuarterBridgeII;
                                break;
                            default:
                                aIStrain = AIStrainGageConfiguration.FullBridgeI;
                                break;
                        }
                        switch (exctype[i])
                        {
                            case 0:
                                aIExcitation = AIExcitationSource.External;
                                break;
                            case 1:
                                aIExcitation = AIExcitationSource.Internal;
                                break;
                            case 2:
                                aIExcitation = AIExcitationSource.None;
                                break;
                            default:
                                aIExcitation = AIExcitationSource.Internal;
                                break;
                        }
                        aIChannels.Add(myTask.AIChannels.CreateStrainGageChannel(physicalChannelComboBox.Items[phychannel[i]].ToString(),
                            "Strain" + phychannel[i].ToString(), minval[i], maxval[i], aIStrain, aIExcitation, gagexc[i] / 1000.0,
                            gagefac[i] / 1000000, gageinit[i] / 1000.0, gageres[i], gagepois[i], gagewire[i] / 1000.0, AIStrainUnits.Strain));
                    }
                    btn9 = 1;
                }
                myTask.ConfigureLogging(fileopentdms, TdmsLoggingOperation.OpenOrCreate, LoggingMode.LogAndRead, group);
                myTask.Timing.ConfigureSampleClock("", Convert.ToDouble(rateNumeric.Value),
SampleClockActiveEdge.Rising, SampleQuantityMode.ContinuousSamples, 1000);
                myTask.Control(TaskAction.Verify);
                runningTask = myTask;
                for (int j = 0; j < aIChannels.Count; j++)
                {
                    kan = 0;
                    string kanal = "";
                    kanal = aIChannels[j].VirtualName + "_" + group + "_" + graphcount;
                    chart1.Series.Add(kanal);
                    kan = chart1.Series.IndexOf(kanal);
                    chart1.Series[kan].ChartType = SeriesChartType.FastPoint;
                    datas.Insert(j,0);
                    if (chtype[j] == 1 || chtype[j] == 2)
                    {
                        chart1.Series[kan].YAxisType = AxisType.Secondary;
                        chart1.Series.Last().Color = Color.FromArgb(255-Convert.ToInt16(255*j/aIChannels.Count()),20, Convert.ToInt16(255 * j / aIChannels.Count()));
                        chart1.Series.Last().MarkerStyle = MarkerStyle.Diamond;
                        //chart1.ChartAreas[0].AxisY2.Maximum = maxval[j];
                        //chart1.ChartAreas[0].AxisY2.Minimum = minval[j];
                    }
                    else
                    {
                        chart1.Series.Last().Color = Color.FromArgb(20, Convert.ToInt16(255 * j / aIChannels.Count()), 255- Convert.ToInt16(255 * j / aIChannels.Count()));
                        chart1.Series.Last().MarkerStyle = MarkerStyle.Circle;
                    }
                    //else
                    //{
                    //    chart1.ChartAreas[0].AxisY.Maximum=
                    //    chart1.ChartAreas[0].AxisY.Maximum = maxval[j];
                    //    chart1.ChartAreas[0].AxisY.Minimum = minval[j];
                    //}
                }
                analogInReader = new AnalogMultiChannelReader(myTask.Stream);
                analogInReader.SynchronizeCallbacks = true;
                analogInReader.BeginReadWaveform(10, myAsyncCallback, runningTask);

            }
            savecount = 0;
            if (loadExt||loadLive.Checked)
            {
                chart1.Series.Add("Load"+ group + "_" + graphcount);
                chart1.Series.Last().ChartType = SeriesChartType.FastPoint;
                chart1.Series.Last().Color = Color.Lime;
                chart1.Series.Last().MarkerStyle = MarkerStyle.Square;
            }
            if (saveExt.Checked)
            {
                //myTask.AddGlobalChannel("External_Com");
            }
            chart1.Series.Add("Position"+ group +"_"+graphcount);
            chart1.Series.Last().ChartType = SeriesChartType.FastPoint;
            chart1.Series.Last().Color = Color.Black;
            chart1.Series.Last().MarkerStyle = MarkerStyle.Cross;
            if (deney == 2)
            {
                CreateGraphPage("Experiment", 1);
            }
            else
            {
                CreateGraphPage(group, 1);
            }
        }
        int graphpre = 0;
        int graphcal = 0;
        int graphexp = 0;
        int kan = 0;
        private void serialAct_CheckedChanged(object sender, EventArgs e)
        {
            if (serialAct.Checked)
            {
                this.Width = 560 + sizeadd;
                portList.Items.AddRange(SerialPort.GetPortNames());
            }
            else
            {
                this.Width = 420 + sizeadd;
                try
                {
                    if (serialPort1.IsOpen)
                    {
                        serialPort1.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
        }


        private void gageChan_SelectedIndexChanged(object sender, EventArgs e)
        {
            voltagechannel = gageChan.SelectedIndex - 1;
            Properties.Settings.Default["GageChan"] = gageChan.SelectedIndex.ToString();
            Properties.Settings.Default.Save();
        }

        private void chart1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //Point mousePoint = new Point(e.X, e.Y);
            //double valy1 = chart1.ChartAreas[0].AxisY.PixelPositionToValue(mousePoint.Y);
            //double valx1 = chart1.ChartAreas[0].AxisX.PixelPositionToValue(mousePoint.Y);
            //double valy2 = chart1.ChartAreas[0].AxisY2.PixelPositionToValue(mousePoint.Y);
            //double valx2 = chart1.ChartAreas[0].AxisX2.PixelPositionToValue(mousePoint.Y);
            //DataPoint point1 = new DataPoint(valx1, valy1);
            //DataPoint point2 = new DataPoint(valx2, valy2);

            //for (int i = 0; i < chart1.Series.Count(); i++)
            //{
            //    foreach (DataPoint dataPoint in chart1.Series[i].Points)
            //    {
            //        if (dataPoint.YValues[0] < valy1 + 0.01 && dataPoint.YValues[0] > valy1 - 0.01 &&
            //            dataPoint.XValue < valx1 + 0.01 && dataPoint.XValue > valx1 - 0.01)
            //        {
            //            toolax = 0;
            //            chart1.ChartAreas[0].CursorX.AxisType = AxisType.Primary;
            //            chart1.ChartAreas[0].CursorY.AxisType = AxisType.Primary;
            //            break;
            //        }
            //        else if (dataPoint.YValues[0] < valy2 + 0.01 && dataPoint.YValues[0] > valy2 - 0.01 &&
            //            dataPoint.XValue < valx2 + 0.01 && dataPoint.XValue > valx2 - 0.01)
            //        {
            //            toolax = 1;
            //            chart1.ChartAreas[0].CursorX.AxisType = AxisType.Secondary;
            //            chart1.ChartAreas[0].CursorY.AxisType = AxisType.Secondary;
            //            break;
            //        }
            //    }
            //    if (chart1.Series[i].Points.Contains(point1))
            //    {
            //        toolax = 0;
            //        chart1.ChartAreas[0].CursorX.AxisType = AxisType.Primary;
            //        chart1.ChartAreas[0].CursorY.AxisType = AxisType.Primary;
            //        break;
            //    }
            //    else if (chart1.Series[i].Points.Contains(point2))
            //    {
            //        toolax = 1;
            //        chart1.ChartAreas[0].CursorX.AxisType = AxisType.Secondary;
            //        chart1.ChartAreas[0].CursorY.AxisType = AxisType.Secondary;
            //        break;
            //    }
            //}
        }
        double motpos0 = 0;
        private void upArrow_Click(object sender, EventArgs e)
        {
            emCounter = 1;
            motsteps = motorTrack.Value;
            motfreq = 1000000 * stepinc / motspeed;
            var motorsend = string.Format("{1:D10}{0:D10}", motsteps, Convert.ToInt32(motfreq));
            Send($"310{spedmode}0", "");
            texts = motorsend;
            motorpass = 1;
            motdir = 1;
            stopMot.Enabled = true;
            groupBox19.Enabled = false;
        }

        private void downArrow_Click(object sender, EventArgs e)
        {
            emCounter = 1;
            if (motpos + motsteps * stepinc <= motmax)
            {
                motsteps = motorTrack.Value;
                motfreq = 1000000*(stepinc / motspeed);
                var motorsend = string.Format("{1:D10}{0:D10}", motsteps, Convert.ToInt32(motfreq));
                Send($"300{spedmode}0", "");
                texts = motorsend;
                motorpass = 1;
                motdir = 0;
                groupBox19.Enabled = false;
                stopMot.Enabled = true;
            }
            else
            {
                MessageBox.Show("Motor range is not enough!", "Motor Range Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void threadPitch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                threadPitch.Text = TextFormat(threadPitch.Text, "pitch", 0, motmax);
                pitch = Convert.ToDouble(threadPitch.Text) * 1000000;
                Properties.Settings.Default["pitch"] = threadPitch.Text;
                Properties.Settings.Default.Save();
                stepIncbox.Text = Convert.ToString(pitch * stepang / (360 * micstep));
                stepinc = pitch * stepang / (360 * micstep);
            }
        }

        private void stepAng_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                stepAng.Text = TextFormat(stepAng.Text, "ang", 0, 360);
                stepang = Convert.ToDouble(stepAng.Text);
                Properties.Settings.Default["ang"] = stepAng.Text;
                Properties.Settings.Default.Save();
                stepIncbox.Text = Convert.ToString(pitch * stepang / (360 * micstep));
                stepinc = pitch * stepang / (360 * micstep);
            }
        }
        double stepinc = 0;
        private void stepSpeed_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                stepSpeed.Text = TextFormat(stepSpeed.Text, "motorspeed", 0, 1000000);
                motspeed = Convert.ToDouble(stepSpeed.Text);
                Properties.Settings.Default["motorspeed"] = stepSpeed.Text;
                Properties.Settings.Default.Save();
                motfreq = 1000000* stepinc / motspeed;
            }
        }
        double motpos = 0;
        double motmax = 0;
        double motspeed = 0;
        double stepang = 0;
        double pitch = 0;
        int micstep = 0;
        double motfreq = 0;
        long motsteps = 0;
        int motdir = 0;
        private void motorMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                motorMax.Text = TextFormat(motorMax.Text, "motormax", 0, 200);
                motmax = Convert.ToDouble(motorMax.Text) * 1000000;
                Properties.Settings.Default["motormax"] = motorMax.Text;
                Properties.Settings.Default.Save();
                //verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
            }
        }


        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void stopMot_Click(object sender, EventArgs e)
        {
            texts = "";
            automot = 0;
            approaching = 0;
            stopMot.Enabled = false;
            groupBox19.Enabled = true;
            loadcell.Enabled = true;
            loadTare.Enabled = true;
            loadLive.Enabled = true;
            Send("Z0000", "");
        }
        int motorcon = 0;
        private void motorDrive_CheckedChanged(object sender, EventArgs e)
        {
            if (motorDrive.Checked)
            {
                motorcon = 1;
                motorApp.Checked = true;
                textBox21.Enabled = false;
            }
            else
            {
                textBox21.Enabled = true;
                motorcon = 0;
            }
        }

        private void portList_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void measureType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (measureType.SelectedIndex == 0)
            {
                maximumValueNumeric.Maximum = 10;
                maximumValueNumeric.Minimum = -10;
                minimumValueNumeric.Maximum = 10;
                minimumValueNumeric.Minimum = -10;
                maximumValueNumeric.Value = 10;
                minimumValueNumeric.Value = -10;
                groupBox18.Enabled = false;
                groupBox16.Enabled = false;
                gageBox.Enabled = false;
            }
            else if (measureType.SelectedIndex == 1)
            {
                maximumValueNumeric.Maximum = 500;
                maximumValueNumeric.Minimum = -2147483648;
                minimumValueNumeric.Maximum = 500;
                minimumValueNumeric.Minimum = -2147483648;
                maximumValueNumeric.Value = 100;
                minimumValueNumeric.Value = 0;
                cjcSourceComboBox.SelectedIndex = 1;
                cjcValueNumeric.Value = 25;
                groupBox18.Enabled = true;
                groupBox16.Enabled = false;
                gageBox.Enabled = false;
            }
            else if (measureType.SelectedIndex == 2)
            {
                maximumValueNumeric.Maximum = 500;
                maximumValueNumeric.Minimum = -2147483648;
                minimumValueNumeric.Maximum = 500;
                minimumValueNumeric.Minimum = -2147483648;
                maximumValueNumeric.Value = 100;
                minimumValueNumeric.Value = 0;
                rtdBox.SelectedIndex = 1;
                rtdExcType.SelectedIndex = 0;
                groupBox18.Enabled = false;
                groupBox16.Enabled = true;
                gageBox.Enabled = false;
            }
            else if (measureType.SelectedIndex == 3)
            {
                maximumValueNumeric.Maximum = 10;
                maximumValueNumeric.Minimum = -10;
                minimumValueNumeric.Maximum = 10;
                minimumValueNumeric.Minimum = -10;
                maximumValueNumeric.Value = 10;
                minimumValueNumeric.Value = -10;
                bridgeBox.SelectedIndex = 4;
                excBox.SelectedIndex = 0;
                gageNom.Value = 1000;
                groupBox18.Enabled = false;
                groupBox16.Enabled = false;
                gageBox.Enabled = true;
            }
        }
        int extsay = 0;
        private void tryExtCom_Click(object sender, EventArgs e)
        {
            try
            {
                if (extsay == 0)
                {
                    serialPort1.BaudRate = Convert.ToUInt16(baudBox.Items[baudBox.SelectedIndex].ToString());
                    serialPort1.PortName = portList.SelectedItem.ToString();
                    serialPort1.Open();
                    tryExtCom.Text = "Close" + Environment.NewLine + "Com";
                    extsay = 1;
                    timer4.Start();
                }
                else
                {
                    serialPort1.Close();
                    extsay = 0;
                    tryExtCom.Text = "Try" + Environment.NewLine + "Com";
                    timer4.Stop();
                }
            }
            catch (Exception ex)
            {
                extsay = 0;
                tryExtCom.Text = "Try" + Environment.NewLine + "Com";
                serialPort1.Close();
                MessageBox.Show(ex.Message, "COM Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                timer4.Stop();
            }
        }

        private void gagePoisson_TextChanged(object sender, EventArgs e)
        {

        }

        private void gagePoisson_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {

            }
        }

        private void rtdExc_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {

            }
        }

        private void rtdRes_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {

            }
        }

        private void startChars_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                try
                {
                    if (startChars.TextLength != 0)
                    {
                        serialPort1.Write(startChars.Text);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "COM Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }
        int chartprop = 0;
        private void ChartProps_Click(object sender, EventArgs e)
        {
            if (chartprop == 0)
            {
                chartprop = 1;
                ChartProps.Text = "Chart Properties <<";
                this.Width = 769;
            }
            else
            {
                chartprop = 0;
                ChartProps.Text = "Chart Properties >>";
                this.Width = 420 + sizeadd;
            }
        }
        int showbut = 0;
        private void showData_Click(object sender, EventArgs e)
        {
            int sersay = chart1.Series.Count - 1;
            tim2i = 0;
            graphcount = 0;
            for (int i = sersay; i > -1; i--)
            {
                chart1.Series[i].Points.Clear();
                chart1.Series.RemoveAt(i);
                chartptcnt = 0;
            }
            if (showbut == 0)
            {
                if (btn9 == 1||loadLive.Checked)
                {
                    string zaman = "\\" + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss");
                    directorset(zaman);
                    fileopen = pathtemp + "\\Preview";
                    deney = 3;
                    Tdms_Saver("Preview");
                    showData.Text = "Stop";
                    showbut = 1;
                }
                else
                {
                    MessageBox.Show("Please set channels for experiment first!", "Preview Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                timer2.Stop();
                //loadLive.Checked = false;
                stopButton.Enabled = false;
                startButton.Enabled = true;
                deney = -1;
                executeCal.Enabled = true;
                showData.Text = "Show Preview";
                showbut = 0;
                tabControl1.SelectedIndex = 4;
                Thread.Sleep(timer2.Interval);
                if (btn9 == 1)
                {
                    runningTask = null;
                    myTask.Dispose();
                }
            }

        }

        private void gageShow_TextChanged(object sender, EventArgs e)
        {
            
        }
        int temperChan = 0;
        private void tempChan_SelectedIndexChanged(object sender, EventArgs e)
        {
            temperChan = tempChan.SelectedIndex - 1;
            Properties.Settings.Default["TempCon"] = tempChan.SelectedIndex.ToString();
            Properties.Settings.Default.Save();
        }

        private void goHome_Click(object sender, EventArgs e)
        {
            emCounter = 1;
            stopMot.Enabled = true;
            approaching = 2;
            connection = "Motor is Moving to Home";
            groupBox19.Enabled = false;
            var motorsend = string.Format("{1:D10}{0:D10}", motsteps, Convert.ToInt32(motfreq));
            texts = motorsend;
            motorpass = 1;
            motdir = 0;
            loadLive.Checked = true;
            Send($"311{spedmode}0", "");
            loadcell.Enabled = false;
            loadTare.Enabled = false;
            loadLive.Enabled = false;
            timer3.Start();
            //AutoConAsync(-10, 0, 0, false);
        }
        int automot = 0;
        private void coarsePst_Click(object sender, EventArgs e)
        {
            emCounter = 1;
            stopMot.Enabled = true;
            approaching = 2;
            connection = "Motor is Moving to Sample";
            groupBox19.Enabled = false;
            var motorsend = string.Format("{1:D10}{0:D10}", motsteps, Convert.ToInt32(motfreq));
            texts = motorsend;
            motorpass = 1;
            motdir = 0;
            loadLive.Checked = true;
            Send($"301{spedmode}0", "");
            loadcell.Enabled = false;
            loadTare.Enabled = false;
            loadLive.Enabled = false;
            timer3.Start();
            //AutoConAsync(pressure, 0, 0, true);
        }

        private void g1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                Properties.Settings.Default["GageCo"] = g1.Text;
                Properties.Settings.Default.Save();
                gageco = Convert.ToDouble(g1.Text);
            }
        }

        private void g0_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                Properties.Settings.Default["GageCo0"] = g0.Text;
                Properties.Settings.Default.Save();
                gageco0 = Convert.ToDouble(g0.Text);
            }
        }

        private void motorPos_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                if (motpos > Convert.ToDouble(motorPos.Text)*1000000 && (Convert.ToDouble(motorPos.Text) >= 0))
                {
                    motsteps = Math.Abs(Convert.ToInt32((motpos - Convert.ToDouble(motorPos.Text) * 1000000) / stepinc));
                    motfreq = 1000000*stepinc / motspeed; 
                    var motorsend = string.Format("{1:D10}{0:D10}", motsteps, Convert.ToInt32(motfreq));
                    Send($"310{spedmode}0", "");
                    texts = motorsend;
                    motorpass = 1;
                    motdir = 1;
                    groupBox19.Enabled = false;
                }
                else if (motpos < Convert.ToDouble(motorPos.Text)*1000000 && (Convert.ToDouble(motorPos.Text) <= motmax))
                {
                    motsteps = Math.Abs(Convert.ToInt32(((Convert.ToDouble(motorPos.Text) * 1000000 - motpos) / stepinc)));
                    motfreq = 1000000*stepinc / motspeed;
                    var motorsend = string.Format("{1:D10}{0:D10}", motsteps, Convert.ToInt32(motfreq));
                    Send($"300{spedmode}0", "");
                    texts = motorsend;
                    motorpass = 1;
                    motdir = 0;
                    groupBox19.Enabled = false;
                }
                stopMot.Enabled = true;
            }
        }

        private void helpNi_Click(object sender, EventArgs e)
        {
            string filePath = Directory.GetCurrentDirectory() + "\\helper\\Device_Pinouts.chm";
            Help.ShowHelp(this, filePath);
        }

        private void tryUc45_Click(object sender, EventArgs e)
        {
            Send("60000", "");
        }
        private void forceCo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            if (e.KeyChar == (char)13)
            {
                forcecoef = Convert.ToDouble(forceCo.Text);
                Properties.Settings.Default["forceCo"] = forceCo.Text;
                Properties.Settings.Default.Save();
            }
        }
        private void stopCal_Click(object sender, EventArgs e)
        {
            texts = "";
            emCounter = 0;
            Send("Z0000", "");
            groupBox8.Enabled = true;
            stopCal.Enabled = false;
        }
        private void tabControl1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)64)
            {
                MessageBox.Show("This is the end of the program!");
            }
        }
        int spedmode = 0;
        private void micstepBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            micstep = Convert.ToUInt16(micstepBox.SelectedItem);
            Properties.Settings.Default["micros"] = micstepBox.SelectedIndex.ToString();
            Properties.Settings.Default.Save();
            stepIncbox.Text = Convert.ToString(pitch * stepang / (360 * micstep));
            stepinc = pitch * stepang / (360 * micstep);
            spedmode = micstepBox.SelectedIndex;
        }
        int clicksay = 0;
        private void chart1_MouseClick(object sender, MouseEventArgs e)
        {
            Point mousePoint = new Point(e.X, e.Y);
            HitTestResult result = chart1.HitTest(mousePoint.X, mousePoint.Y);
            chart1.ChartAreas[0].CursorX.SetCursorPixelPosition(mousePoint, false);
            chart1.ChartAreas[0].CursorY.SetCursorPixelPosition(mousePoint, false);
            if (result.ChartElementType == ChartElementType.DataPoint)
            {
                    toolTip1.SetToolTip(chart1, "X:" + chart1.Series[result.Series.Name].Points[result.PointIndex].XValue.ToString("0.000") + "\n" + "Y:" +
        chart1.Series[result.Series.Name].Points[result.PointIndex].YValues.Last().ToString("0.000"));
            }
            else if (result.ChartElementType == ChartElementType.LegendArea)
            {
                if (clicksay == 0)
                {
                    chart1.Legends[0].LegendStyle = LegendStyle.Table;
                    clicksay = 1;
                }
                else if (clicksay == 1)
                {
                    chart1.Legends[0].LegendStyle = LegendStyle.Row;
                    clicksay = 2;
                }
                else
                {
                    chart1.Legends[0].LegendStyle = LegendStyle.Column;
                    clicksay = 0;
                }
            }
            else
            {
                try
                {
                    toolTip1.SetToolTip(chart1, "X:" + chart1.ChartAreas[0].AxisX.PixelPositionToValue(mousePoint.X).ToString("0.000") + "\n" + "Y:" +
        chart1.ChartAreas[0].AxisY.PixelPositionToValue(mousePoint.Y).ToString("0.000"));

                }
                catch (Exception ex)
                {
                    label11.Text = ex.Message;

                }
            }

        }
        int feed = 0;
        private void motorFeed_Click(object sender, EventArgs e)
        {
            Send("MFEED", "");
            feed = 1;
        }
        int xymotor = 0;
        private void rightArrow_Click(object sender, EventArgs e)
        {
            Send($"V00{Convert.ToInt16(xEncode.Value):D1}{Convert.ToInt16(xStep.Value):D1}","");
            xymotor = 1;
            groupBox19.Enabled = false;
        }

        private void leftArrow_Click(object sender, EventArgs e)
        {
            Send($"V01{Convert.ToInt16(xEncode.Value):D1}{Convert.ToInt16(xStep.Value):D1}", "");
            xymotor = 1;
            groupBox19.Enabled = false;
        }

        private void forwardArrow_Click(object sender, EventArgs e)
        {
            Send($"V10{Convert.ToInt16(yEncode.Value):D1}{Convert.ToInt16(yStep.Value):D1}", "");
            xymotor = 1;
            groupBox19.Enabled = false;
        }

        private void backArrow_Click(object sender, EventArgs e)
        {
            Send($"V11{Convert.ToInt16(yEncode.Value):D1}{Convert.ToInt16(yStep.Value):D1}", "");
            xymotor = 1;
            groupBox19.Enabled = false;
        }
        int senscom = 0;
        private void digiCon_CheckedChanged(object sender, EventArgs e)
        {
            if (digiCon.Checked)
            {
                Send($"7{Convert.ToInt16(sensorCon.Value):D4}", "");
                senscom = 1;
            }
        }
        private void sensorCon_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                Send($"7{Convert.ToInt16(sensorCon.Value):D4}", "");
                senscom = 1;
            }
        }
        int readbalup = 0;
        int readbalmov = 0;
        int readbalbas = 0;
        int balancing = 0;
        private void upBal_Click(object sender, EventArgs e)
        {
            Send("81000", "");
            readbalup = 1;
            balancing = 1;
        }

        private void movBal_Click(object sender, EventArgs e)
        {
            Send("82000", "");
            readbalmov = 1;
            balancing = 1;
        }

        private void basBal_Click(object sender, EventArgs e)
        {
            Send("83000", "");
            readbalbas = 1;
            balancing = 1;
        }

        private void touchLevel_Click(object sender, EventArgs e)
        {
            Send($"80{slopeType.SelectedIndex}00", "");
        }

        private void vibCheck_Click(object sender, EventArgs e)
        {
            Send("84000", ""); //connect to above commands!! 81'000' especially with varying DigitalLowPassFilter(DLPF) of mpu6050
        }

        private void motorApp_CheckedChanged(object sender, EventArgs e)
        {
            if (!motorApp.Checked)
            {
            }
        }
        List<string> Xpos = new List<string>();
        List<string> Ypos = new List<string>();
        private void indXpst_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                Xpos.Insert(comboBox2.SelectedIndex - 1, indXpst.Text);
                if (Xpos.Count > comboBox2.SelectedIndex)
                {
                    Xpos.RemoveAt(comboBox2.SelectedIndex);
                }
                if (Ypos.Count < Xpos.Count)
                {
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (duration.Count < Xpos.Count)
                {
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (depth.Count < Xpos.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (amplitude.Count < Xpos.Count)
                {
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                    holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                    retHold.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (Xpos.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }

            }
        }

        private void indYpst_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                Ypos.Insert(comboBox2.SelectedIndex - 1, indYpst.Text);
                if (Ypos.Count > comboBox2.SelectedIndex)
                {
                    Ypos.RemoveAt(comboBox2.SelectedIndex);
                }
                if (Xpos.Count < Ypos.Count)
                {
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (depth.Count < Ypos.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (duration.Count < Ypos.Count)
                {
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (amplitude.Count < Ypos.Count)
                {
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                    holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                    retHold.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (Ypos.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }

            }
        }

        private void getHumid_Click(object sender, EventArgs e)
        {
            rhcom = 1;
            Send("84000","");
        }
        string hxsend = "";
        bool loadExt = false;
        private void controlExt_CheckedChanged(object sender, EventArgs e)
        {
            if (controlExt.Checked)
            {
                isMcuAdc.Checked = false;
                Send("85310", "");
                hxsend = string.Format("T{0:00000}", hxSet.Value);
            }
            if (controlExt.Checked || isMcuAdc.Checked)
            {
                digiCon.Checked = false;
                useGagePress.Checked = false;
                loadExt = true;
            }
            else
            {
                Send("85300", "");
            }
        }

        private void loadTare_Click(object sender, EventArgs e)
        {
            hxcom = 1;
            Send("85200", "");
        }

        private void loadLive_CheckedChanged(object sender, EventArgs e)
        {
            if (loadLive.Checked)
            {
                loadcell.Enabled = false;
                loadTare.Enabled = false;
                if (motorpass != 1)
                {
                    Send("85100", "");
                }
                timer3.Start();
            }
            else
            {
                loadcell.Enabled = true;
                loadTare.Enabled = true;
                Send("HXFIN", "");
                timer3.Stop();
            }
        }
        string[,] definexp;
        private void saveSet_Click(object sender, EventArgs e)
        {
            try
            {
                definexp = new string[depth.Count, 19];
                for (int i = 0; i < depth.Count; i++)
                {
                    definexp[i, 0] = depth[i];
                    definexp[i, 1] = Xpos[i];
                    definexp[i, 2] = Ypos[i];
                    definexp[i, 3] = speed[i];
                    definexp[i, 4] = duration[i];
                    definexp[i, 5] = box5[i].ToString();
                    definexp[i, 6] = box11[i].ToString();
                    definexp[i, 7] = amplitude[i];
                    definexp[i, 8] = frequency[i];
                    definexp[i, 9] = interval[i];
                    definexp[i, 10] = dTemp[i];
                    definexp[i, 11] = dTamp[i];
                    definexp[i, 12] = dTfreq[i];
                    definexp[i, 13] = dTlag[i];
                    definexp[i, 14] = dTspeed[i];
                    definexp[i, 15] = dTwhen[i].ToString();
                    definexp[i, 16] = retStep[i].ToString();
                    definexp[i, 17] = holdAt[i].ToString();
                    definexp[i, 18] = retHold[i];
                }
                saveFileDialog1.CreatePrompt = true;
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.InitialDirectory = @"D:\Indenter";
                saveFileDialog1.Title = "Save Experiment Settings";
                saveFileDialog1.DefaultExt = "indenter";
                saveFileDialog1.Filter = "INDENTER Files (*.indenter)|*.indenter|All Files(*.*)|*.*";
                saveFileDialog1.ShowDialog();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                    StreamWriter writer = new StreamWriter(saveFileDialog1.FileName,false);
                    for (int j = 0; j < depth.Count(); j++)
                    {
                        for (int i = 0; i < 19; i++)
                        {
                            writer.WriteLine(definexp[j, i]);
                        }
                    }
                writer.Close();
            }
            catch (Exception ex){ MessageBox.Show(ex.Message);
            }

        }

        private void loadSet_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.InitialDirectory = @"D:\Indenter";
                openFileDialog1.Title = "Load Experiment Settings";
                openFileDialog1.DefaultExt = "indenter";
                openFileDialog1.Filter = "INDENTER Files (*.indenter)|*.indenter|All Files(*.*)|*.*";
                openFileDialog1.ShowDialog();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
}

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                depth.RemoveRange(0, depth.Count);
                speed.RemoveRange(0, speed.Count);
                duration.RemoveRange(0, duration.Count);
                amplitude.RemoveRange(0, amplitude.Count);
                frequency.RemoveRange(0, frequency.Count);
                interval.RemoveRange(0, interval.Count);
                box5.RemoveRange(0, box5.Count);
                box11.RemoveRange(0, box11.Count);
                box3.RemoveRange(0, box3.Count);
                Xpos.RemoveRange(0, Xpos.Count);
                Ypos.RemoveRange(0, Ypos.Count);
                dTemp.RemoveRange(0, dTemp.Count);
                dTamp.RemoveRange(0, dTamp.Count);
                dTfreq.RemoveRange(0, dTfreq.Count);
                dTspeed.RemoveRange(0, dTspeed.Count);
                dTlag.RemoveRange(0, dTlag.Count);
                dTwhen.RemoveRange(0, dTwhen.Count);
                retStep.RemoveRange(0, retStep.Count);
                holdAt.RemoveRange(0, holdAt.Count);
                retHold.RemoveRange(0, retHold.Count);
                for (int j = comboBox2.Items.Count - 1; j > 0; j--)
                {
                    comboBox2.Items.RemoveAt(j);
                }
                comboBox2.SelectedIndex = 0;
                string cizgi = "";
                StreamReader reader = new StreamReader(openFileDialog1.FileName);
                int i = 0;
                while (cizgi != null)
                {
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        depth.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        Xpos.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        Ypos.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        speed.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        duration.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        box5.Insert(i, Convert.ToBoolean(cizgi));
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        box11.Insert(i, Convert.ToBoolean(cizgi));
                        box3.Insert(i, !(box11[i] || box5[i]));
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        amplitude.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        frequency.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        interval.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        dTemp.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        dTamp.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        dTfreq.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        dTspeed.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        dTlag.Insert(i, cizgi);
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        dTwhen.Insert(i, Convert.ToInt16(cizgi));
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        retStep.Insert(i, Convert.ToInt16(cizgi));
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        holdAt.Insert(i, Convert.ToDecimal(cizgi));
                    }
                    cizgi = reader.ReadLine();
                    if (cizgi != null)
                    {
                        retHold.Insert(i,cizgi);
                    }
                    comboBox2.Items.Add(i + 1);
                    i++;
                }
                if (comboBox2.Items.Count < 2)
                {
                    comboBox2.Items.Add(1);
                }
                comboBox2.SelectedIndex = i;
                reader.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
}
        int jspass = 0;
        private void timer6_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (isConnected)
            {
                groupBox5.Enabled = true;
                groupBox8.Enabled = true;
                groupBox19.Enabled = true;
                label11.Text = "Connected MCU via " + serialPort2.PortName + Environment.NewLine + connection;
                tim1say++;
                if (initialize == -1)
                {

                    initialize = 1;
                    tim1say = 0;
                    timer6.Stop();
                }
                else if (initialize == 1 && comread.Contains("nected"))
                {
                    timer6.Stop();
                    initialize = 0;
                    DialogResult result = MessageBox.Show("Connected to UC45, last (or default) values will be sent when click OK.",
        "Connection to Device", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    InitialCommandsAsync(result);
                }
                else if (jspass == 1)
                {
                    timer6.Stop();
                    jspass = 0;
                    micstepBox.SelectedIndex = spedmode;
                }
                else if(approaching==2&& pass == 1)
                {
                    pass = 0;
                    if (receive == "UPMOT")
                    {
                        Send("STFIN", "");
                        connection = "";
                        receive = "";
                        MessageBox.Show("Motor at Home Position!", "Auto Land", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        motorPos.Text = "0.0";
                        motpos = 0;
                        approaching = 0;
                    }
                    else if (receive.Contains("CLOSe"))
                    {
                        MessageBox.Show("Surface has been found succesfully!" + Environment.NewLine +
          " Use auto approach for more controlled landing.", "Auto Land", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        micstepBox.SelectedIndex = 5;
                        approaching = 0;
                        Send("STFIN", "");
                        loadcell.Enabled = true;
                        loadTare.Enabled = true;
                        loadLive.Enabled = true;
                        try
                        {
                            abq = receive.Replace("CLOSe", "");
                            abqi = Convert.ToInt16(abq);
                            motpos = motpos + abqi * stepinc / 2;
                            motorPos.Text = Convert.ToString(motpos / 1000000);
                            comread = "";
                            abqi = 0;
                            receive = "";
                        }
                        catch (Exception ex)
                        {
                            label11.Text = ex.Message;
                        }
                    }
                    timer6.Stop();
                }
                else if (pass == 1 && autopass == 1 && comread != "")
                {
                    pass = 0;
                    connection = "Motor Moving!";
                    stopMot.Enabled = true;
                    if (approaching == 0)
                    {
                        motorPos.Text = Convert.ToString(motpos / 1000000);
                        Properties.Settings.Default["motpos"] = motorPos.Text;
                        Properties.Settings.Default.Save();
                        //verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                        timer6.Stop();
                        groupBox19.Enabled = true;
                        Send("STFIN", "");
                        autopass = 0;
                        if (comread.Contains("UPMOT"))
                        {
                            motpos = 0;
                            motorPos.Text = "0.0";
                            Properties.Settings.Default["motpos"] = motorPos.Text;
                            Properties.Settings.Default.Save();
                            //verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                            timer6.Stop();
                            groupBox19.Enabled = true;
                            autopass = 0;
                        }
                    }
                    else
                    {
                        autopass = 0;
                        comread = "";
                        connection = "";
                        timer6.Stop();
                    }

                }
                else if (pass == 1 && exppass == -1 && plug % 2 == 1)
                {
                    repsay = 0;
                    exppass = -2;
                    executeExp.Enabled = true;
                    timer6.Stop();
                    label11.Text = comread;
                    MessageBox.Show("Ready to Go! Click 'Execute' when you ready.");
                }
                else if (pass == 1 && feed == 1)
                {
                    timer6.Stop();
                    pass = 0;
                    motpos = abqi * stepinc*(Math.Pow(2,spedmode-5));
                    motorPos.Text = Convert.ToString(motpos / 1000000);
                    abqi = 0;
                    Properties.Settings.Default["motpos"] = motorPos.Text;
                    Properties.Settings.Default.Save();
                    //verticalProgressBar2.Value = 100 - Convert.ToInt16(motpos * 100 / motmax);
                    feed = 0;
                    Send("FBFIN", "");
                }
                else if (pass == 3)
                {
                    pass = 0;
                    leftArrow.Enabled = xylimits[0];
                    rightArrow.Enabled = xylimits[1];
                    upArrow.Enabled = xylimits[2];
                    backArrow.Enabled = xylimits[3];
                    timer6.Stop();
                }
                else if (pass == 1 && (xymotor == 1||jsBox.Checked || deney==1))
                {
                    pass = 0;
                    string reccon="";
                    for(int i=0; i < receive.Count(); i++)
                    {
                        if (receive[i] == '+')
                        {
                            if (Convert.ToDouble(reccon) - 32767 < 0)
                            {
                                label11.Text = "XY Direction Error!";
                            }
                            break;
                        }
                        else if (receive[i] == '-')
                        {
                            if (Convert.ToDouble(reccon) - 32767 > 0)
                            {
                                label11.Text = "XY Direction Error!";
                            }
                            break;
                        }
                        reccon = reccon+receive[i];
                    }
                    if (receive.Contains("X"))
                    {
                        xkon = xkon + ((xinc*(Convert.ToDouble(reccon) - 32767)))/1000;
                        xPosition.Text = xkon.ToString();
                    }
                    else if (receive.Contains("Y"))
                    {
                        ykon = ykon + ((yinc * (Convert.ToDouble(reccon) - 32767)))/1000;
                        yPosition.Text = ykon.ToString();
                    }
                    timer6.Stop();
                    xymotor = 0;
                }
                else if (pass == 1 && senscom == 1)
                {
                    pass = 0;
                    pressShow.Text = receives;
                    Send("SCFIN", "");
                    senscom = 0;
                    timer6.Stop();
                }
                else if (pass == 1 && hxcom == 1)
                {
                    pass = 0;
                    pressShow.Text = receives;
                    loadcellData.Text = receives;
                    Send("HXFIN", "");
                    hxcom = 0;
                    timer6.Stop();
                }
                else if (pass == 1 && rhcom == 1)
                {
                    pass = 0;
                    relHum.Text = Convert.ToString(humid);//ikili okuma olacak!! tek mesajda DHT icin duzenle
                    humTemp.Text = Convert.ToString(tmed);
                    Send("RHFIN", "");
                    rhcom = 0;
                    timer6.Stop();
                }
                else if (pass == 2)
                {
                    pass = 0;
                    try
                    {
                        double xr = Convert.ToDouble(recx) / 16384.0;
                        double yr = Convert.ToDouble(recy) / 16384.0;
                        double zr = Convert.ToDouble(recz) / 16384.0;
                        double theta = Math.Acos(zr / Math.Sqrt(xr * xr + yr * yr + zr * zr)) * 180 / Math.PI;
                        double phi = Math.Atan(yr / xr) * 180 / Math.PI;
                        if (readbalbas == 1)
                        {
                            balance[2, 0] = phi;
                            balance[2, 1] = theta;
                            groundX.Text = Convert.ToString(balance[2, 0]);
                            groundY.Text = Convert.ToString(balance[2, 1]);
                            readbalbas = 0;
                        }
                        else if (readbalmov == 1)
                        {
                            balance[1, 0] = phi;
                            balance[1, 1] = theta;
                            indentX.Text = Convert.ToString(balance[1, 0]);
                            indentY.Text = Convert.ToString(balance[1, 1]);
                            readbalmov = 0;
                        }
                        else if (readbalup == 1)
                        {
                            balance[0, 0] = phi;
                            balance[0, 1] = theta;
                            headX.Text = Convert.ToString(balance[0, 0]);
                            headY.Text = Convert.ToString(balance[0, 1]);
                            readbalup = 0;
                        }
                        balancing = 0;
                    }
                    catch
                    {
                        balancing = 1;
                        if (readbalbas == 1)
                        {
                            Send("83000", "");
                        }
                        else if (readbalmov == 1)
                        {
                            Send("82000", "");
                        }
                        else if (readbalup == 1)
                        {
                            Send("81000", "");
                        }
                    }
                    timer6.Stop();
                }
                else if (pass == 1)
                {
                    tim1say = 0;
                    timer6.Stop();
                }
                else if (receive == "PROCESS")
                {
                    groupBox8.Enabled = true;
                    stopCal.Enabled = false;
                    stopExp.Enabled = false;
                    groupBox5.Enabled = true;
                    executeExp.Enabled = false;
                    button3.Enabled = true;
                    timer6.Stop();
                }
                else if (directpass == 1 || exppass == 1 || calpass==1 || motorpass == 1)
                {
                    microsayac++;
                    if (microsayac == 12 && rep > 5)
                    {
                        exppass = 0;
                        directpass = 0;
                        calpass = 0;
                        expcounter = 0;
                        motorpass = 0;
                        approaching = 0;
                        label11.Text = "MCU CONNECTION MAY BE LOST!";
                        rep = 0;
                        microsayac = 0;
                        texts = "";
                        textexp.Clear();
                        tim1say = 0;
                        timer6.Stop();
                    }
                    else if (microsayac == 12 && plug % 2 == 1)
                    {
                        //Send("Repet", "");
                        microsayac = 0;
                        rep++;
                    }
                }
            }
            else
            {
                var ports = SerialPort.GetPortNames();
                if (tim1say < ports.Length && !serialPort2.IsOpen)
                {
                    try
                    {
                        serialPort2.PortName = ports[tim1say];
                        serialPort2.Open();
                        serialPort2.Write("I");
                    }
                    catch (Exception ex)
                    {
                        label11.Text = ex.Message;
                        isConnected = false;
                        //MessageBox.Show(ex.Message);
                    }
                }
                else if (tim1say > 10)
                {
                    tim1say = 0;
                    try
                    {
                        nocon++;
                        serialPort2.Close();
                    }
                    catch (Exception ex)
                    {
                        label11.Text = ex.Message;
                    }
                }
                else if (tim1say >= ports.Length && !serialPort2.IsOpen)
                {
                    tim1say = 0;
                }
                if (nocon > ports.Count() * 100)
                {
                    label11.Text = "No device detected.";
                    button12.Text = "Secure Plug-in";
                    plug++;
                    tim1say = 0;
                    isConnected = false;
                    groupBox5.Enabled = false;
                    groupBox8.Enabled = false;
                    groupBox19.Enabled = false;
                    serialPort2.Close();
                    timer6.Stop();
                }
                tim1say++;
            }
        }

        private void motorTrack_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(motorTrack,motorTrack.Value.ToString());
        }

        private void yEncode_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(yEncode, yEncode.Value.ToString());
        }

        private void xEncode_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(xEncode, xEncode.Value.ToString());
        }
        string heatSender = "";
        double[] setTem = new double[3];
        private void sampleTSet_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!isComboheat)
            {
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
                {

                    e.Handled = true;

                }
                else if (e.KeyChar == (char)13)
                {
                    setTemper[selectedHeater.SelectedIndex] = sampleTSet.Text;
                    setTem[selectedHeater.SelectedIndex] = Convert.ToDouble(setTemper[selectedHeater.SelectedIndex]);
                    if (heatAct.Checked)
                    {
                        dutyNominal[selectedHeater.SelectedIndex] = (setTem[selectedHeater.SelectedIndex] - cnstHeat[selectedHeater.SelectedIndex]) / slopeHeat[selectedHeater.SelectedIndex];
                        var sendtemp = Convert.ToInt32( setTem[selectedHeater.SelectedIndex]* 100);
                        heatSender = $"{sendtemp:D6}|0" + "000000000000";
                        heat_com = 1;
                        Send($"9{selectedHeater.SelectedIndex}10{Convert.ToInt16(feedInternal.Checked)}", "");
                    }
                }
            }
        }
        int heat_com = 0;
        private void heatAct_CheckedChanged(object sender, EventArgs e)
        {
            if (!isComboheat)
            {
                activeHeat[selectedHeater.SelectedIndex] = heatAct.Checked;
                if (heatAct.Checked && ((!feedMode[selectedHeater.SelectedIndex] && heatChannels[selectedHeater.SelectedIndex] != -1)||feedMode[selectedHeater.SelectedIndex]))
                {
                    var sendtemp = Convert.ToInt32(Convert.ToDouble(setTemper[selectedHeater.SelectedIndex]) * 100);
                    heatSender = $"{sendtemp:D6}|0" + Convert.ToInt16(isSampleheat[selectedHeater.SelectedIndex]).ToString() +"00000000000";
                    heat_com = 1;
                    Send($"9{selectedHeater.SelectedIndex}10{Convert.ToInt16(feedInternal.Checked)}", "");
                }
                else if (heatAct.Checked)
                {
                    isComboheat = true;
                    heatAct.Checked = false;
                    activeHeat[selectedHeater.SelectedIndex] = heatAct.Checked;
                    isComboheat = false;
                    MessageBox.Show("Please Set Feedback Channel First!");
                }
                else
                {
                    Send($"9{selectedHeater.SelectedIndex}000", "");
                }
            }
        }

        private void heatOnof_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void heatPid_CheckedChanged(object sender, EventArgs e)
        {
        }
        List<string> dTemp = new List<string>();
        List<string> dTspeed = new List<string>();
        List<int> dTwhen = new List<int>();
        List<string> dTamp = new List<string>();
        List<string> dTfreq = new List<string>();
        List<string> dTlag = new List<string>();
        List<int> retStep = new List<int>();
        private void dtRamp_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                dTemp.Insert(comboBox2.SelectedIndex - 1, dtRamp.Text);
                if (dTemp.Count > comboBox2.SelectedIndex)
                {
                    dTemp.RemoveAt(comboBox2.SelectedIndex);
                }
                if (depth.Count < dTemp.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                }
                if (dTemp.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }
            }
        }

        private void dtSpeed_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                dTspeed.Insert(comboBox2.SelectedIndex - 1, dtSpeed.Text);
                if (dTspeed.Count > comboBox2.SelectedIndex)
                {
                    dTspeed.RemoveAt(comboBox2.SelectedIndex);
                }
                if (depth.Count < dTspeed.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                }
                if (dTspeed.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }
            }
        }

        private void dtAmp_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                dTamp.Insert(comboBox2.SelectedIndex - 1, dtAmp.Text);
                if (dTamp.Count > comboBox2.SelectedIndex)
                {
                    dTamp.RemoveAt(comboBox2.SelectedIndex);
                }
                if (depth.Count < dTamp.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                    retHold.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                }
                if (dTamp.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }
            }
        }

        private void dtF_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                dTfreq.Insert(comboBox2.SelectedIndex - 1, dtF.Text);
                if (dTfreq.Count > comboBox2.SelectedIndex)
                {
                    dTfreq.RemoveAt(comboBox2.SelectedIndex);
                }
                if (depth.Count < dTfreq.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                    holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                    retHold.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (dTfreq.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }
            }
        }

        private void dtPhase_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                dTlag.Insert(comboBox2.SelectedIndex - 1, dtF.Text);
                if (dTlag.Count > comboBox2.SelectedIndex)
                {
                    dTlag.RemoveAt(comboBox2.SelectedIndex);
                }
                if (depth.Count < dTlag.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                    holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                    retHold.Insert(comboBox2.SelectedIndex - 1, "0");
                }
                if (dTlag.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }
            }
        }

        private void label61_Click(object sender, EventArgs e)
        {

        }

        private void retractStep_CheckedChanged(object sender, EventArgs e)
        {
            if (!combochange)
            {
                retStep.Insert(comboBox2.SelectedIndex - 1, Convert.ToInt16(retractStep.Checked));
                if (retStep.Count > comboBox2.SelectedIndex)
                {
                    retStep.RemoveAt(comboBox2.SelectedIndex);
                }
                if (depth.Count < retStep.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                }
                if (retStep.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }
            }
        }
        int btn7say = 0;
        int exptype = 0;
        private void button7_Click(object sender, EventArgs e)
        {
            btn7say++;
            if (btn7say > 2)
            {
                btn7say = 0;
            }
            if (btn7say == 0)
            {
                Send("43000","");
                button7.Text = "Control: Piezo Voltage";
                label14.Text = "Depth (um)";
                label15.Text = "Speed (um/s)";
                exptype = 0;
            }
            else if (btn7say == 1)
            {
                button7.Text = "Control: Force";
                label14.Text = "Max Force (mN)";
                label15.Text = "Speed (mN/s)";
                if (useGagePress.Checked)//Use external USART or NI feed
                {
                    Send("43120", "");
                }
                else if (controlExt.Checked)
                {//Use HX711 or similar..
                    Send("43100", "");
                }
                else if (isMcuAdc.Checked)//use ADC --strain gage via MCU
                {
                    Send("43110", "");
                }
                exptype = 1;
            }
            else if (btn7say == 2)
            {
                button7.Text = "Control: Displacement";
                if (isMcuAdc.Checked)//use ADC --strain gage via MCU
                {
                    Send("43210", "");
                }
                else
                {
                    Send("43220", "");
                }
                label14.Text = "Depth (um)";
                label15.Text = "Speed (um/s)";
                exptype = 2;
            }
        }
        double xkon, ykon = 0;
        double xinc = 1;
        double yinc = 1;
        private void xPosition_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void yPosition_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
        bool [] activeHeat  = new bool [3];
        bool [] isSampleheat = new bool [3];
        string [] setTemper = new string[3];
        bool[] feedMode = new bool[3];
        bool isComboheat = false;
        bool[] isBand = new bool[3];
        double[] bandInter = new double[3];
        private void selectedHeater_SelectedIndexChanged(object sender, EventArgs e)
        {
            isComboheat = true;
            sampleTSet.Text = setTemper[selectedHeater.SelectedIndex];
            heatAct.Checked = activeHeat[selectedHeater.SelectedIndex];
            isSample.Checked = isSampleheat[selectedHeater.SelectedIndex];
            feedInternal.Checked = feedMode[selectedHeater.SelectedIndex];
            feedNi.Checked=!feedMode[selectedHeater.SelectedIndex];
            heaterTimeConst.Text = heaterTime[selectedHeater.SelectedIndex].ToString();
            heatSensDev.Text = heaterSensorDev[selectedHeater.SelectedIndex].ToString();
            bandInterval.Text = bandInter[selectedHeater.SelectedIndex].ToString();
            bandControl.Checked = isBand[selectedHeater.SelectedIndex];
            heatSlopeTb.Text = slopeHeat[selectedHeater.SelectedIndex].ToString();
            heatConstTb.Text = cnstHeat[selectedHeater.SelectedIndex].ToString();
            heaterDuty.Value = Convert.ToDecimal(duty[selectedHeater.SelectedIndex]);
            proGain.Value = Convert.ToDecimal(proGains[selectedHeater.SelectedIndex]);

            if (isFeed.Contains(true)&&heatChannels[selectedHeater.SelectedIndex]!=-1)
            {
                heaterFeedList.SelectedIndex = heatChannels[selectedHeater.SelectedIndex];
            }
            isComboheat = false;
        }

        private void isSample_CheckedChanged(object sender, EventArgs e)
        {
            if (!isComboheat)
            {
                isSampleheat[selectedHeater.SelectedIndex] = isSample.Checked;
            }
        }

        private void feedNi_CheckedChanged(object sender, EventArgs e)
        {
            if (!isComboheat)
            {
                isComboheat = true;
                feedMode[selectedHeater.SelectedIndex] = !feedNi.Checked;
                feedInternal.Checked = !feedNi.Checked;
                isComboheat = false;
            }
        }

        private void feedInternal_CheckedChanged(object sender, EventArgs e)
        {
            if (!isComboheat)
            {
                isComboheat = true;
                feedMode[selectedHeater.SelectedIndex] = feedInternal.Checked;
                feedNi.Checked = !feedInternal.Checked;
                isComboheat = false;
                if (!feedNi.Checked && heatChannels[selectedHeater.SelectedIndex]!=-1)
                {
                    heatChannels[selectedHeater.SelectedIndex] = -1;
                }
            }
        }

        private void xTour_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                try
                {
                    xinc = Convert.ToDouble(xTour.Text);
                    Properties.Settings.Default["xenc"] = xTour.Text;
                    Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void yTour_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                try
                {
                    yinc = Convert.ToDouble(yTour.Text);
                    Properties.Settings.Default["yenc"] = yTour.Text;
                    Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void isHeaterFeed_CheckedChanged(object sender, EventArgs e)
        {
            if (!tcnochange)
            {

            }
        }
        double measuredT = 0;
        double[] measuredTmcu = new double[3];
        int[] feedChannel = new int[5];
        int []duty = new int[3] ;
        double error = 0;
        bool err_chan = false;
        int []heatTime = new int [3];
        int thresholdDuty = 0;
        bool []integral_control = new bool [3];
        double []T0 = new double[3];
        double []T1 = new double[3];
        double []dutyNominal = new double[3];
        double []slopeHeat = new double [3];
        double []cnstHeat = new double [3];
        double[] proGains = new double[3];
        private void timer1_Elapsed(object sender, ElapsedEventArgs e)
        {
            for(int i = 0; i < 3; i++)
            {
                if (runningTask != null && heatChannels[i]!=-1 && data!=null)
                {
                    measuredT = data[feedChannel[heatChannels[i]]].GetScaledData().Average();
                    if (selectedHeater.SelectedIndex == i)
                    {
                        sampleTMeasure.Text = measuredT.ToString();
                    }
                    if (activeHeat[i]&&heat_com==0 && autopass != 1 && feed != 1 && hxcom != 1 && motorpass != 1 && directpass != 1 && !sending
                        && calpass != 1 && exppass == 0 && xymotor != 1 && rhcom != 1 && balancing != 1 && datacomp == 0)
                    {
                        /*if (setTem[i] - 0.05 > measuredT)
                        {
                            Send("9"+i+"+++", "");
                        }
                        else if (setTem[i] + 0.05 < measuredT)
                        {
                            Send("9" + i + "---", "");
                        }*/
                        if (!integral_control[i])
                        {
                            var tempduty = duty[i];
                            if (error + heaterSensorDev[i] < setTem[i] - measuredT || error - heaterSensorDev[i] > setTem[i] - measuredT)
                            {
                                err_chan = true;
                                error = setTem[i] - measuredT;
                                if((isBand[i]&& bandInter[i]>Math.Abs(error))||!isBand[i])
                                {
                                    duty[i] = Convert.ToInt32(dutyNominal[i] + (proGains[i] * error / 100));
                                }
                                else
                                {
                                    if(bandInter[i] + setTem[i] < measuredT)
                                    {
                                        duty[i] = 5;
                                    }
                                    else if(setTem[i] - bandInter[i] > measuredT)
                                    {
                                        duty[i] = 40;
                                    }
                                }
                                if (duty[i] > 50)
                                {
                                    duty[i] = 50;
                                    label11.Text = "Duty is at Max!";
                                }
                                else if (duty[i] < 1)
                                {
                                    duty[i] = 1;
                                    label11.Text = "Duty is at Min!";
                                }
                                if (duty[i] != tempduty)
                                {
                                    Send($"9{i}D{duty[i]:D2}", "");
                                    heaterDuty.Value = duty[i];
                                }
                                err_chan = false;
                            }
                        }
                        else
                        {
                            heatTime[i]++;
                            if (heatTime[i] == 1)
                            {
                                duty[i] = 10;
                                Send($"9{i}D{ duty[i]:D2}", "");
                                heaterDuty.Value = duty[i];
                                heaterDuty.Value = duty[i];

                            }
                            else if (heatTime[i] == heaterTime[i]+1)
                            {
                                T0[i] = measuredT;
                                duty[i] = 20;
                                Send($"9{i}D{duty[i]:D2}", "");
                                heaterDuty.Value = duty[i];
                                heaterDuty.Value = duty[i];

                            }
                            else if(heatTime[i] == 2*heaterTime[i]+1)
                            {
                                T1[i] = measuredT;
                                integral_control[i] = false;
                                slopeHeat[i] = (T1[i] - T0[i]) / (10);
                                cnstHeat[i] = T0[i] - slopeHeat[i] * 10;
                                dutyNominal[i] = (setTem[i] - cnstHeat[i]) / slopeHeat[i];
                                if (selectedHeater.SelectedIndex == i)
                                {
                                    heatSlopeTb.Text = Convert.ToString(slopeHeat[i]);
                                    heatConstTb.Text = Convert.ToString(cnstHeat[i]);
                                }
                                heatTime[i] = 0;
                            }
                        }
                    }
                }
                else if (feedMode[i])
                {
                    measuredT = measuredTmcu[i];
                    if (selectedHeater.SelectedIndex == i)
                    {
                        sampleTMeasure.Text = measuredT.ToString();
                    }
                }
            }

        }
        int[] heatChannels = new int[3];
        private void heaterFeedList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!isComboheat)
            {
                heatChannels[selectedHeater.SelectedIndex] = heaterFeedList.SelectedIndex;
                if (feedNi.Checked&&deney==-1)
                {
                    DialogResult dialogResult = MessageBox.Show("Starting to Data Acquisition ?", "Heater Feedback", MessageBoxButtons.OKCancel);
                    if (dialogResult == DialogResult.OK && runningTask==null)
                    {
                        deney = 3;
                        //Tdms_Saver("Preview_Heat");
                        showData.PerformClick();
                        timer1.Start();
                    }
                }
            }
        }

        private void stopFeed_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 3; i++)
            {
                if (!feedMode[i] && heatChannels[selectedHeater.SelectedIndex] != -1)
                {
                    heatChannels[i] = -1;
                    activeHeat[i] = false;
                    Send($"9{i}000", "");
                }
            }
            if (deney == 3)
            {
                showData.PerformClick();
            }
        }

        private void saveExt_CheckedChanged(object sender, EventArgs e)
        {
            if (saveExt.Checked)
            {
                gageChan.Items.Add("External_Com");
                tempChan.Items.Add("External_Com");
            }
            else
            {
                gageChan.Items.Remove("External_Com");
                tempChan.Items.Remove("External_Com");
            }
        }
        double extContPara = 0;
        private void extPara_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                extContPara = Convert.ToDouble(extPara.Text);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void heaterDuty_ValueChanged(object sender, EventArgs e)
        {
            if (!err_chan && !isComboheat)
            {
                duty[selectedHeater.SelectedIndex] = Convert.ToInt16(heaterDuty.Value);
                dutyNominal[selectedHeater.SelectedIndex] = Convert.ToDouble(heaterDuty.Value);
                Send($"9{selectedHeater.SelectedIndex}D{duty[selectedHeater.SelectedIndex]:D2}", "");
            }
        }
        double[] heaterTime = new double[3];
        double[] heaterSensorDev = new double[3];
        private void heaterTimeConst_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                try
                {
                    heaterTime[selectedHeater.SelectedIndex] = Convert.ToDouble(heaterTimeConst.Text);
                    //Properties.Settings.Default["yenc"] = yTour.Text;
                    //Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void heatSensDev_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                try
                {
                    heaterSensorDev[selectedHeater.SelectedIndex] = Convert.ToDouble(heatSensDev.Text);
                    //Properties.Settings.Default["yenc"] = yTour.Text;
                    //Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void findHeaterChar_Click(object sender, EventArgs e)
        {
            heatTime[selectedHeater.SelectedIndex] = 0;
            integral_control[selectedHeater.SelectedIndex] = true;
        }

        private void heatSlopeTb_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                try
                {
                    slopeHeat[selectedHeater.SelectedIndex] = Convert.ToDouble(heatSlopeTb.Text);
                    //Properties.Settings.Default["yenc"] = yTour.Text;
                    //Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void heatConstTb_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                try
                {
                    cnstHeat[selectedHeater.SelectedIndex] = Convert.ToDouble(heatConstTb.Text);
                    //Properties.Settings.Default["yenc"] = yTour.Text;
                    //Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void proGain_ValueChanged(object sender, EventArgs e)
        {
            proGains[selectedHeater.SelectedIndex] = Convert.ToDouble(proGain.Value);
        }

        private void bandControl_CheckedChanged(object sender, EventArgs e)
        {
            isBand[selectedHeater.SelectedIndex] = bandControl.Checked;
        }

        private void bandInterval_KeyPress(object sender, KeyPressEventArgs e)
        {
            bandInter[selectedHeater.SelectedIndex] = Convert.ToDouble(bandInterval.Text);
        }
        List<string> retHold = new List<string>();
        List<decimal> holdAt = new List<decimal>();
        double loadAppThres = 0;
        private void loadThres_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                loadAppThres = Convert.ToDouble(loadThres.Text);
            }
        }

        private void autoApp_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void holdDur_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                retHold.Insert(comboBox2.SelectedIndex - 1, holdDur.Text);
                if (retHold.Count > comboBox2.SelectedIndex)
                {
                    retHold.RemoveAt(comboBox2.SelectedIndex);
                }
                if (depth.Count < dTamp.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                }
                if (retHold.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }
            }
        }

        private void holdPercent_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                holdAt.Insert(comboBox2.SelectedIndex - 1, holdPercent.Value);
                if (holdAt.Count > comboBox2.SelectedIndex)
                {
                    holdAt.RemoveAt(comboBox2.SelectedIndex);
                }
                if (depth.Count < dTamp.Count)
                {
                    depth.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTwhen.Insert(comboBox2.SelectedIndex - 1, 0);
                    speed.Insert(comboBox2.SelectedIndex - 1, "0");
                    duration.Insert(comboBox2.SelectedIndex - 1, "0");
                    Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                    Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                    amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                    box5.Insert(comboBox2.SelectedIndex - 1, false);
                    box11.Insert(comboBox2.SelectedIndex - 1, false);
                    box3.Insert(comboBox2.SelectedIndex - 1, true);
                    frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                    interval.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                    dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                    retHold.Insert(comboBox2.SelectedIndex - 1, "0");
                    retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                }
                if (holdAt.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
                {
                    comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
                }
            }
        }
        double calHoldDur = 0;
        private void textBox18_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {

                e.Handled = true;

            }
            else if (e.KeyChar == (char)13)
            {
                calHoldDur = Convert.ToDouble(textBox18.Text);
            }
        }
        double speedApp = 0;
        private void appSpeed_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                speedApp = Convert.ToDouble(appSpeed.Text);
            }
        }

        private void hxSet_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
            else if (e.KeyChar == (char)13)
            {
                if (controlExt.Checked)
                {
                    Send("85310", "");
                }
                else if (isMcuAdc.Checked)
                {
                    Send("85320", "");
                }
                hxsend =string.Format("T{0:00000}", hxSet.Value);
            }
        }

        private void isMcuAdc_CheckedChanged(object sender, EventArgs e)
        {
            if (isMcuAdc.Checked)
            {
                loadExt = true;
                Send("85320", "");
                controlExt.Checked = false;
                hxsend = string.Format("T{0:00000}", hxSet.Value);
            }
            if (controlExt.Checked || isMcuAdc.Checked)
            {
                digiCon.Checked = false;
                useGagePress.Checked = false;
                loadExt = true;
            }
            else
            {
                Send("85300", "");
            }
        }

        private void actAppOnly_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void forwardArrow_EnabledChanged(object sender, EventArgs e)
        {
            if (!forwardArrow.Enabled && groupBox19.Enabled)
            {
                MessageBox.Show("There is no more range at forward direction, please turn backward.", "XY Stage Limits", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void backArrow_EnabledChanged(object sender, EventArgs e)
        {
            if (!backArrow.Enabled && groupBox19.Enabled)
            {
                MessageBox.Show("There is no more range at backward direction, please turn forward.", "XY Stage Limits", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void rightArrow_EnabledChanged(object sender, EventArgs e)
        {
            if (!rightArrow.Enabled && groupBox19.Enabled)
            {
                MessageBox.Show("There is no more range at rightward direction, please turn leftward.","XY Stage Limits",MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void leftArrow_EnabledChanged(object sender, EventArgs e)
        {
            if (!leftArrow.Enabled && groupBox19.Enabled)
            {
                MessageBox.Show("There is no more range at leftward direction, please turn rightward.", "XY Stage Limits", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void commentSave_Click(object sender, EventArgs e)
        {
            if (explogPath == "")
            {
                explogPath = pathsave+"\\Log.txt";
            }
            File.AppendAllText(explogPath, Environment.NewLine +"Comments: " + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss") + Environment.NewLine + commentsExp.Text);
        }

        private void tipName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                Properties.Settings.Default["tipName"] = tipName.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void sampleName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar == (char)13)
            {
                Properties.Settings.Default["sampleName"] = sampleName.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void procedureName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                Properties.Settings.Default["procedureName"] = procedureName.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void giveWhile_SelectedIndexChanged(object sender, EventArgs e)
        {
            dTwhen.Insert(comboBox2.SelectedIndex - 1, giveWhile.SelectedIndex);
            for (int i = 0; i < 3; i++)
            {
                if (i == dTwhen[comboBox2.SelectedIndex - 1])
                {
                    giveWhile.SetItemChecked(dTwhen[comboBox2.SelectedIndex - 1], true);
                }
                else
                {
                    giveWhile.SetItemCheckState(i, CheckState.Unchecked);
                }
            }
            if (dTwhen.Count > comboBox2.SelectedIndex)
            {
                dTwhen.RemoveAt(comboBox2.SelectedIndex);
            }
            if (depth.Count < dTwhen.Count)
            {
                depth.Insert(comboBox2.SelectedIndex - 1, "0");
                dTemp.Insert(comboBox2.SelectedIndex - 1, "0");
                dTlag.Insert(comboBox2.SelectedIndex - 1, "0");
                speed.Insert(comboBox2.SelectedIndex - 1, "0");
                duration.Insert(comboBox2.SelectedIndex - 1, "0");
                Xpos.Insert(comboBox2.SelectedIndex - 1, "0");
                Ypos.Insert(comboBox2.SelectedIndex - 1, "0");
                amplitude.Insert(comboBox2.SelectedIndex - 1, "0");
                box5.Insert(comboBox2.SelectedIndex - 1, false);
                box11.Insert(comboBox2.SelectedIndex - 1, false);
                box3.Insert(comboBox2.SelectedIndex - 1, true);
                frequency.Insert(comboBox2.SelectedIndex - 1, "0");
                interval.Insert(comboBox2.SelectedIndex - 1, "0");
                dTspeed.Insert(comboBox2.SelectedIndex - 1, "0");
                dTamp.Insert(comboBox2.SelectedIndex - 1, "0");
                dTfreq.Insert(comboBox2.SelectedIndex - 1, "0");
                retStep.Insert(comboBox2.SelectedIndex - 1, 0);
                holdAt.Insert(comboBox2.SelectedIndex - 1, 0);
                retHold.Insert(comboBox2.SelectedIndex - 1, "0");
            }
            if (dTwhen.Count <= comboBox2.SelectedIndex && comboBox2.SelectedIndex + 2 > comboBox2.Items.Count)
            {
                comboBox2.Items.Add(comboBox2.SelectedIndex + 1);
            }
        }

        private void loadcell_Click(object sender, EventArgs e)
        {
            hxcom = 1;
            Send("85000", "");
        }

        private void jsBox_CheckedChanged(object sender, EventArgs e)
        {
            if (jsBox.Checked)
            {
                Send("JSENN", "");
                groupBox19.Enabled = false;
            }
            else
            {
                Send("JSUNN", "");
                groupBox19.Enabled = true;

            }
        }
    }
    }
