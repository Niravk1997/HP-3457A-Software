using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Speech.Synthesis;
using System.Windows.Threading;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Reflection;

namespace HP_3457A
{
    public static class Serial_COM_Info
    {
        public static bool isConnected = false;

        //HP 3457A COM Device Info
        public static string COM_Port;
        public static int COM_BaudRate;
        public static int COM_Parity;
        public static int COM_StopBits;
        public static int COM_DataBits;
        public static int COM_Handshake;
        public static int COM_WriteTimeout;
        public static int COM_ReadTimeout;
        public static bool COM_RtsEnable;
        public static int GPIB_Address;

        public static string folder_Directory;
    }

    public partial class MainWindow : Window
    {
        //Reference to the graph window
        DateTime_Graph_Window HP3457A_DateTime_Graph_Window;

        //Reference to the graph window
        Graphing_Window HP3457A_Graph_Window;

        //Reference to the N graph Window
        N_Sample_Graph_Window HP3457A_N_Graph_Window;

        //Reference to Measurement Table
        Measurement_Data_Table HP3457A_Table;
        string Current_Measurement_Unit = "VDC";

        //HP3457A serial connection
        SerialPort HP3457A;

        //Which Measurement is currently selected
        int Measurement_Selected = 0;
        int Selected_Measurement_type = 0;
        //VDC = 0
        //VAC = 1
        //2Ohm = 2
        //4Ohm = 3
        //ADC = 4
        //AAC = 5
        //VDCVAC = 6
        //ADCAAC = 7
        //FREQ = 8
        //PER = 9

        double Measurement_Range = 0; //0 = Auto
        int NDigit = 7; //Setting the NPLC determines the Digits
        double NPLC = 10;
        bool isNDigit_7_Selected = false;
        bool isNPLC_10_Selected = false;
        bool show_7Digit_Info = false;
        int NPLC_100_Delay_Time = 1000;

        //All Serial Write Commands are stored in this queue
        BlockingCollection<string> SerialWriteQueue = new BlockingCollection<string>();

        //Clear Logs after this count
        int Auto_Clear_Output_Log_Count = 20;

        //Lets the function know the queue has data
        bool isUserSendCommand = false;
        bool isSamplingOnly = false;
        bool isUpdateSpeed_Changed = false;
        string SCPI_Command;

        //User decides whether to save data to text file or not
        //to save output log or not
        bool saveOutputLog = false;
        //to save measurements or not
        bool saveMeasurements = false;
        //to add data to table
        bool save_to_Table = false;
        //to add data to graphs
        bool save_to_Graph = false;
        bool Save_to_N_Graph = false;
        bool Save_to_DateTime_Graph = false;

        //Data is stored in these queues, waiting for it to be written to text files
        BlockingCollection<string> save_data_VDC = new BlockingCollection<string>();
        BlockingCollection<string> save_data_VAC = new BlockingCollection<string>();
        BlockingCollection<string> save_data_2Ohm = new BlockingCollection<string>();
        BlockingCollection<string> save_data_4Ohm = new BlockingCollection<string>();
        BlockingCollection<string> save_data_ADC = new BlockingCollection<string>();
        BlockingCollection<string> save_data_AAC = new BlockingCollection<string>();
        BlockingCollection<string> save_data_VDCVAC = new BlockingCollection<string>();
        BlockingCollection<string> save_data_ADCAAC = new BlockingCollection<string>();
        BlockingCollection<string> save_data_FREQ = new BlockingCollection<string>();
        BlockingCollection<string> save_data_PER = new BlockingCollection<string>();

        //Options for Speech Synthesizer
        SpeechSynthesizer Voice = new SpeechSynthesizer();
        int Speech_Value_Precision = 1;
        int isSpeechActive = 0;
        int isSpeechContinuous = 0;
        int isSpeechMIN = 0;
        int isSpeechMAX = 0;
        double Speech_Continuous_Voice_Value = 0;
        double Speech_min_value = 0;
        double Speech_max_value = 0;

        //Default border color for when a switch is selected or not
        SolidColorBrush Selected = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00CE30"));
        SolidColorBrush Deselected = new SolidColorBrush((Color)ColorConverter.ConvertFromString("White"));

        //Options for Measurement Data sampling speed
        double UpdateSpeed = 1000;

        //COM Select Window
        COM_Select_Window COM_Select;

        //Timer for getting data from multimeter at specified update speed.
        private System.Timers.Timer Speech_MIN_Max;
        private System.Timers.Timer Speech_Measurement_Interval;
        private System.Timers.Timer DataTimer;
        private DispatcherTimer runtime_Timer;
        private DispatcherTimer Process_Data;
        private System.Timers.Timer saveMeasurements_Timer;

        //Allow data timer to get data from multimeter or not
        bool DataSampling = false;

        //Data is stored here for display
        BlockingCollection<string> measurements = new BlockingCollection<string>();
        int Total_Samples = 0;
        int Invalid_Samples = 0;

        //Display Measurement as
        bool B_FSI_Double = true;
        bool B_RTZ_Display = false;

        //Calculate Runtime from this
        DateTime StartDateTime;

        //Min, Max, Avg values
        //Program will compare input values to these value
        //and update these values
        double min = 0;
        double max = 0;
        double avg = 0;
        int AVG_Calculate = 1;
        int avg_count = 0;
        int avg_factor = 1000;
        int avg_resolution = 5;
        int resetMinMaxAvg = 1;

        public MainWindow()
        {
            InitializeComponent();
            if (Thread.CurrentThread.CurrentCulture.Name != "en-US")
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                insert_Log("Culture set to english-US, decimal numbers will use dot as the seperator.", 0);
                insert_Log("Write decimal values with a dot as a seperator, not a comma.", 2);
            }
            Create_GetDataTimer();
            General_Timer();
            SetupSpeechSythesis();
            Check_Speech_MIN_MAX_Timer();
            Continuous_Voice_Measurement();
            Save_measurements_to_files_Timer();
            Load_Main_Window_Settings();
            insert_Log("Click the Config Menu then click Connect.", 5);
            insert_Log("AR488 GPIB Adapter and a HP 3457A are required to use this software.", 5);
        }

        private void Save_measurements_to_files_Timer()
        {
            saveMeasurements_Timer = new System.Timers.Timer();
            saveMeasurements_Timer.Interval = 60000; //Default is 1 minute;
            saveMeasurements_Timer.AutoReset = false;
            saveMeasurements_Timer.Enabled = false;
            saveMeasurements_Timer.Elapsed += Save_MeasurementData_to_files;
        }

        private void Save_MeasurementData_to_files(Object source, ElapsedEventArgs e)
        {
            string Date = DateTime.UtcNow.ToString("yyyy-MM-dd");
            int VDC_Count = save_data_VDC.Count;
            int VAC_Count = save_data_VAC.Count;
            int TwoOhm_Count = save_data_2Ohm.Count;
            int FourOhm_Count = save_data_4Ohm.Count;
            int ADC_Count = save_data_ADC.Count;
            int AAC_Count = save_data_AAC.Count;
            int VDCVAC_Count = save_data_VDCVAC.Count;
            int ADCAAC_Count = save_data_ADCAAC.Count;
            int FREQ_Count = save_data_FREQ.Count;
            int PER_Count = save_data_PER.Count;

            if (VDC_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "VDC" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_VDC.txt", true))
                    {
                        for (int i = 0; i < VDC_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_VDC.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save VDC measurements to text file.", 1);
                }
            }

            if (VAC_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "VAC" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_VAC.txt", true))
                    {
                        for (int i = 0; i < VAC_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_VAC.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save VAC measurements to text file.", 1);
                }
            }

            if (TwoOhm_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "2WireOhms" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_2WireOhms.txt", true))
                    {
                        for (int i = 0; i < TwoOhm_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_2Ohm.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save 2 Wire Ohms measurements to text file.", 1);
                }
            }

            if (FourOhm_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "4WireOhms" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_4WireOhms.txt", true))
                    {
                        for (int i = 0; i < FourOhm_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_4Ohm.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save 4 Wire Ohms measurements to text file.", 1);
                }
            }

            if (ADC_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "ADC" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_ADC.txt", true))
                    {
                        for (int i = 0; i < ADC_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_ADC.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save ADC measurements to text file.", 1);
                }
            }

            if (AAC_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "AAC" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_AAC.txt", true))
                    {
                        for (int i = 0; i < AAC_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_AAC.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save AAC measurements to text file.", 1);
                }
            }

            if (VDCVAC_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "VDCVAC" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_VDCVAC.txt", true))
                    {
                        for (int i = 0; i < VDCVAC_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_VDCVAC.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save VDCVAC measurements to text file.", 1);
                }
            }

            if (ADCAAC_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "ADCAAC" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_ADCAAC.txt", true))
                    {
                        for (int i = 0; i < ADCAAC_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_ADCAAC.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save ADCAAC measurements to text file.", 1);
                }
            }

            if (FREQ_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "FREQ" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_FREQ.txt", true))
                    {
                        for (int i = 0; i < FREQ_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_FREQ.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save FREQ measurements to text file.", 1);
                }
            }

            if (PER_Count > 0)
            {
                try
                {
                    using (TextWriter datatotxt = new StreamWriter(Serial_COM_Info.folder_Directory + @"\" + "PER" + @"\" + Date + "_" + Serial_COM_Info.COM_Port + "_PER.txt", true))
                    {
                        for (int i = 0; i < PER_Count; i++)
                        {
                            datatotxt.WriteLine(save_data_PER.Take());
                        }
                    }
                }
                catch (Exception)
                {
                    insert_Log("Cannot save PER measurements to text file.", 1);
                }
            }

            saveMeasurements_Timer.Enabled = true;
            if (saveMeasurements == false)
            {
                while (save_data_VDC.TryTake(out _)) { }
                while (save_data_VAC.TryTake(out _)) { }
                while (save_data_2Ohm.TryTake(out _)) { }
                while (save_data_4Ohm.TryTake(out _)) { }
                while (save_data_ADC.TryTake(out _)) { }
                while (save_data_AAC.TryTake(out _)) { }
                while (save_data_VDCVAC.TryTake(out _)) { }
                while (save_data_ADCAAC.TryTake(out _)) { }
                while (save_data_FREQ.TryTake(out _)) { }
                while (save_data_PER.TryTake(out _)) { }
                saveMeasurements_Timer.Enabled = false;
                saveMeasurements_Timer.Stop();
                insert_Log("Save Measurements Queues Cleared.", 0);
            }
        }

        private void SetupSpeechSythesis()
        {
            Voice.Volume = 100;
            Voice.SelectVoiceByHints(VoiceGender.Male);
            Voice.Rate = 1;
        }

        private void General_Timer()
        {
            runtime_Timer = new DispatcherTimer();
            runtime_Timer.Interval = TimeSpan.FromSeconds(1);
            runtime_Timer.Tick += runtime_Update;
            runtime_Timer.Start();
        }

        public void Serial_COM_Selected()
        {
            if (Serial_COM_Info.isConnected == true)
            {
                Connect.IsEnabled = false;
                unlockControls();
                Serial_Connect();
                this.Title = "HP 3457A " + Serial_COM_Info.COM_Port;
                DataSampling = true;
                saveOutputLog = true;
                saveMeasurements = true;
                Stop_Sampling.IsEnabled = true;
                DataTimer.Enabled = true;
                StartDateTime = DateTime.Now;
                Data_process();
                saveMeasurements_Timer.Enabled = true;
                Sampling_Only.IsEnabled = true;
                Local_Exit.IsEnabled = true;
                DataLogger.IsEnabled = true;
            }
        }

        private void runtime_Update(object sender, EventArgs e)
        {
            InvalidSamples_Total.Content = Invalid_Samples.ToString();
            Samples_Total.Content = Total_Samples.ToString();
            if (DataSampling == true)
            {
                Runtime_Timer.Content = GetTimeSpan();
            }
        }

        private void Continuous_Voice_Measurement()
        {
            Speech_Measurement_Interval = new System.Timers.Timer();
            Speech_Measurement_Interval.Interval = 60000; //Default is 1 minute;
            Speech_Measurement_Interval.AutoReset = false;
            Speech_Measurement_Interval.Enabled = false;
            Speech_Measurement_Interval.Elapsed += Check_Continuous_Voice_Measurement;
        }

        private void Check_Continuous_Voice_Measurement(Object source, ElapsedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                try
                {
                    if (isSpeechContinuous == 1)
                    {
                        if (Speech_Continuous_Voice_Value > 999999999)
                        {
                            Voice.Speak("Overload" + " " + MeasurementType_String());
                        }
                        else
                        {
                            Voice.Speak((decimal)Math.Round(Speech_Continuous_Voice_Value, Speech_Value_Precision) + " " + MeasurementType_String());
                        }
                        Speech_Measurement_Interval.Enabled = true;
                    }
                }
                catch (Exception)
                {
                    insert_Log("Speech Synthesizer Continuous Voice measurement feature failed.", 1);
                    insert_Log("Don't worry. Trying again.", 2);
                    Speech_Measurement_Interval.Enabled = true;
                }
            }
            else
            {
                Interlocked.Exchange(ref isSpeechContinuous, 0);
            }
        }

        private void Check_Speech_MIN_MAX_Timer()
        {
            Speech_MIN_Max = new System.Timers.Timer();
            Speech_MIN_Max.Interval = 1000;
            Speech_MIN_Max.AutoReset = false;
            Speech_MIN_Max.Enabled = false;
            Speech_MIN_Max.Elapsed += Check_Speech_MIN_MAX;
        }

        private void Check_Speech_MIN_MAX(Object source, ElapsedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                try
                {
                    if (isSpeechMAX == 1)
                    {
                        if (Speech_max_value <= max)
                        {
                            Voice.Speak("Warning, maximum value of " + (decimal)Math.Round(Speech_max_value, Speech_Value_Precision) + " " + MeasurementType_String() + " reached.");
                            if (max > 9999999999)
                            {
                                Voice.Speak("maximum value is " + "overload" + " " + MeasurementType_String());
                            }
                            else
                            {
                                Voice.Speak("maximum value is " + (decimal)Math.Round(max, Speech_Value_Precision) + " " + MeasurementType_String());
                            }
                        }
                    }
                    if (isSpeechMIN == 1)
                    {
                        if (Speech_min_value >= min)
                        {
                            Voice.Speak("Warning, minimum value of " + (decimal)Math.Round(Speech_min_value, Speech_Value_Precision) + " " + MeasurementType_String() + " reached.");
                            if (min < -9999999999)
                            {
                                Voice.Speak("minimum value is " + "overload" + " " + MeasurementType_String());
                            }
                            else
                            {
                                Voice.Speak("minimum value is " + (decimal)Math.Round(min, Speech_Value_Precision) + " " + MeasurementType_String());
                            }
                        }
                    }
                    Speech_MIN_Max.Enabled = true;
                }
                catch (Exception)
                {
                    insert_Log("Speech Synthesizer MIN and MAX feature failed.", 1);
                    insert_Log("Don't worry. Trying again.", 2);
                    Speech_MIN_Max.Enabled = true;
                }
            }
            if (isSpeechMAX == 0 & isSpeechMIN == 0)
            {
                Speech_MIN_Max.Enabled = false;
                Speech_MIN_Max.Stop();
            }
        }

        private string MeasurementType_String()
        {
            switch (Selected_Measurement_type)
            {
                case 0:
                    return "volts DC";
                case 1:
                    return "volts AC";
                case 2:
                    return "ohms";
                case 3:
                    return "ohms";
                case 4:
                    return "amps DC";
                case 5:
                    return "amps AC";
                case 6:
                    return "volts ACDC";
                case 7:
                    return "amps ACDC";
                case 8:
                    return "hertz";
                case 9:
                    return "seconds";
                default:
                    return "value";
            }
        }

        private (string, string) MeasurementUnit_String()
        {
            switch (Selected_Measurement_type)
            {
                case 0:
                    return ("VDC", "VDC Voltage");
                case 1:
                    return ("VAC", "VAC Voltage");
                case 2:
                    return ("Ω", "Ω 2Wire Ohms");
                case 3:
                    return ("Ω", "Ω 4Wire Ohms");
                case 4:
                    return ("ADC", "ADC Current");
                case 5:
                    return ("AAC", "AAC Current");
                case 6:
                    return ("VAC", "VAC Voltage");
                case 7:
                    return ("AAC", "AAC Current");
                case 8:
                    return ("Hz", "Hz Frequency");
                case 9:
                    return ("s", "T Period");
                default:
                    return ("Unk", "Unknown");
            }
        }

        private string GetTimeSpan()
        {
            TimeSpan span = (DateTime.Now - StartDateTime);
            return (String.Format("{0:00}:{1:00}:{2:00}", span.Hours, span.Minutes, span.Seconds));
        }

        private void unlockControls()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, new ThreadStart(delegate
            {
                Measurements.IsEnabled = true;
                Range.IsEnabled = true;
                Meter_Config.IsEnabled = true;
                UpdateSpeed_Box.IsEnabled = true;
            }));
        }

        private void lockControls()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, new ThreadStart(delegate
            {
                Measurements.IsEnabled = false;
                Range.IsEnabled = false;
                Meter_Config.IsEnabled = false;
                UpdateSpeed_Box.IsEnabled = false;
            }));
        }

        private void Speedup_Interval()
        {
            if (UpdateSpeed > 2000)
            {
                DataTimer.Interval = 0.01;
            }
        }

        private void Restore_Interval()
        {
            DataTimer.Interval = UpdateSpeed;
        }

        private void Serial_Connect()
        {
            HP3457A = new SerialPort(Serial_COM_Info.COM_Port, Serial_COM_Info.COM_BaudRate, (Parity)Serial_COM_Info.COM_Parity, Serial_COM_Info.COM_DataBits, (StopBits)Serial_COM_Info.COM_StopBits);
            HP3457A.WriteTimeout = Serial_COM_Info.COM_WriteTimeout;
            HP3457A.ReadTimeout = Serial_COM_Info.COM_ReadTimeout;
            HP3457A.RtsEnable = Serial_COM_Info.COM_RtsEnable;
            HP3457A.Handshake = (Handshake)Serial_COM_Info.COM_Handshake;
            HP3457A.Open();
        }

        private void Create_GetDataTimer()
        {
            DataTimer = new System.Timers.Timer();
            DataTimer.Interval = 1000;
            DataTimer.Elapsed += HP3457ACommunicateEvent;
            DataTimer.AutoReset = false;
        }

        private void Data_process()
        {
            Process_Data = new DispatcherTimer();
            Process_Data.Interval = TimeSpan.FromSeconds(0);
            Process_Data.Tick += DataProcessor;
            Process_Data.Start();
        }

        private void DataProcessor(object sender, EventArgs e)
        {
            while (measurements.Count > 0)
            {
                try
                {
                    string measurement = measurements.Take();
                    if (measurement == "+1.0000000E+38")
                    {
                        Measurement_Value.Content = "OVLD";
                        Measurement_Scale.Content = "";
                        Display_MIN_MAX_AVG(100000000000);
                    }
                    else if (measurement == "-1.0000000E+38")
                    {
                        Measurement_Value.Content = "OVLD";
                        Measurement_Scale.Content = "";
                        Display_MIN_MAX_AVG(-100000000000);
                    }
                    else
                    {
                        double value = double.Parse(measurement, System.Globalization.NumberStyles.Float);
                        DisplayData(measurement, value);
                        Display_MIN_MAX_AVG(value);
                        setContinuousVoiceMeasurement(value);
                    }
                }
                catch (Exception Ex)
                {
                    if (Show_Display_Error.IsChecked == true)
                    {
                        insert_Log(Ex.ToString(), 2);
                        insert_Log("Sample display process failed. Trying again.", 2);
                    }
                }
            }
            Process_Data.Stop();
        }

        private void setContinuousVoiceMeasurement(double value)
        {
            if (isSpeechContinuous == 1)
            {
                Interlocked.Exchange(ref Speech_Continuous_Voice_Value, value);
            }
        }

        private void DisplayData(string measurement, double value)
        {
            if (B_RTZ_Display == true)
            {
                int Significant_Digits;
                double Range;
                if (Measurement_Range == 0)
                {
                    Range = predict_range_display(value);
                }
                else
                {
                    Range = Measurement_Range;
                }

                Significant_Digits = Calculate_Significant_Digits(Math.Abs(value), Range) + 1;
                if (!(Significant_Digits >= 1 & Significant_Digits <= 15))
                {
                    Display_PSI_Double(measurement, value);
                }
                else
                {
                    string unit = measurement.Substring(measurement.IndexOf("E") + 1);

                    switch (unit)
                    {
                        case "-12":
                        case "-012":
                        case "-11":
                        case "-011":
                        case "-10":
                        case "-010":
                        case "-09":
                        case "-009":
                        case "-08":
                        case "-008":
                        case "-07":
                        case "-007":
                        case "-06":
                        case "-006":
                        case "-05":
                        case "-005":
                        case "-04":
                        case "-004":
                        case "-03":
                        case "-003":
                        case "-02":
                        case "-002":
                        case "-01":
                        case "-001":
                            Measurement_Scale.Content = "m"; //milli
                            Measurement_Value.Content = SignificantDigits.ToString((value * 1E3), Significant_Digits);
                            break;
                        case "+01":
                        case "+001":
                        case "+02":
                        case "+002":
                        case "+03":
                        case "+003":
                        case "+04":
                        case "+004":
                        case "+05":
                        case "+005":
                            Measurement_Scale.Content = "K"; //kilo
                            Measurement_Value.Content = SignificantDigits.ToString((value * 1E-3), Significant_Digits);
                            break;
                        case "+06":
                        case "+006":
                        case "+07":
                        case "+007":
                        case "+08":
                        case "+008":
                            Measurement_Scale.Content = "M"; //Mega
                            Measurement_Value.Content = SignificantDigits.ToString((value * 1E-6), Significant_Digits);
                            break;
                        case "+09":
                        case "+009":
                        case "+10":
                        case "+010":
                        case "+11":
                        case "+011":
                        case "+12":
                        case "+012":
                            Measurement_Scale.Content = "G"; //Giga
                            Measurement_Value.Content = SignificantDigits.ToString((value * 1E-9), Significant_Digits);
                            break;
                        default:
                            Measurement_Scale.Content = "";
                            Measurement_Value.Content = SignificantDigits.ToString((value), Significant_Digits);
                            break;
                    }
                }
            }
            else if (B_FSI_Double == true)
            {
                Display_FSI_Double(measurement, value);
            }
            else
            {
                Display_PSI_Double(measurement, value);
            }
        }

        private void Display_FSI_Double(string measurement, double value)
        {
            string unit = measurement.Substring(measurement.IndexOf("E") + 1);
            switch (unit)
            {
                case "-12":
                case "-012":
                    Measurement_Scale.Content = "p"; //pico official
                    Measurement_Value.Content = double.Parse(measurement.Substring(0, measurement.IndexOf("E")));
                    break;
                case "-11":
                case "-011":
                    Measurement_Value.Content = (decimal)(value * 1E9);
                    Measurement_Scale.Content = "n"; //nano
                    break;
                case "-10":
                case "-010":
                    Measurement_Value.Content = (decimal)(value * 1E9);
                    Measurement_Scale.Content = "n"; //nano
                    break;
                case "-09":
                case "-009":
                    Measurement_Scale.Content = "n"; //nano official
                    Measurement_Value.Content = double.Parse(measurement.Substring(0, measurement.IndexOf("E")));
                    break;
                case "-08":
                case "-008":
                    Measurement_Value.Content = (decimal)(value * 1E6);
                    Measurement_Scale.Content = "μ"; //micro
                    break;
                case "-07":
                case "-007":
                    Measurement_Value.Content = (decimal)(value * 1E6);
                    Measurement_Scale.Content = "μ"; //micro
                    break;
                case "-06":
                case "-006":
                    Measurement_Scale.Content = "μ"; //micro official
                    Measurement_Value.Content = double.Parse(measurement.Substring(0, measurement.IndexOf("E")));
                    break;
                case "-05":
                case "-005":
                    Measurement_Value.Content = (decimal)(value * 1E3);
                    Measurement_Scale.Content = "m"; //milli
                    break;
                case "-04":
                case "-004":
                    Measurement_Value.Content = (decimal)(value * 1E3);
                    Measurement_Scale.Content = "m"; //milli
                    break;
                case "-03":
                case "-003":
                    Measurement_Scale.Content = "m"; //milli official
                    Measurement_Value.Content = double.Parse(measurement.Substring(0, measurement.IndexOf("E")));
                    break;
                case "-02":
                case "-002":
                    Measurement_Value.Content = (decimal)(value * 1E3);
                    Measurement_Scale.Content = "m"; //milli
                    break;
                case "-01":
                case "-001":
                    Measurement_Value.Content = (decimal)(value * 1E3);
                    Measurement_Scale.Content = "m"; //milli
                    break;
                case "+01":
                case "+001":
                    Measurement_Value.Content = (decimal)(value * 1E-3);
                    Measurement_Scale.Content = "K"; //kilo
                    break;
                case "+02":
                case "+002":
                    Measurement_Value.Content = (decimal)(value * 1E-3);
                    Measurement_Scale.Content = "K"; //kilo
                    break;
                case "+03":
                case "+003":
                    Measurement_Scale.Content = "K"; //kilo official
                    Measurement_Value.Content = double.Parse(measurement.Substring(0, measurement.IndexOf("E")));
                    break;
                case "+04":
                case "+004":
                    Measurement_Value.Content = (decimal)(value * 1E-3);
                    Measurement_Scale.Content = "K"; //kilo
                    break;
                case "+05":
                case "+005":
                    Measurement_Value.Content = (decimal)(value * 1E-3);
                    Measurement_Scale.Content = "K"; //kilo
                    break;
                case "+06":
                case "+006":
                    Measurement_Scale.Content = "M"; //Mega official
                    Measurement_Value.Content = double.Parse(measurement.Substring(0, measurement.IndexOf("E")));
                    break;
                case "+07":
                case "+007":
                    Measurement_Value.Content = (decimal)(value * 1E-6);
                    Measurement_Scale.Content = "M"; //Mega
                    break;
                case "+08":
                case "+008":
                    Measurement_Value.Content = (decimal)(value * 1E-6);
                    Measurement_Scale.Content = "M"; //Mega
                    break;
                case "+09":
                case "+009":
                    Measurement_Scale.Content = "G"; //Giga official
                    Measurement_Value.Content = double.Parse(measurement.Substring(0, measurement.IndexOf("E")));
                    break;
                case "+10":
                case "+010":
                    Measurement_Value.Content = (decimal)(value * 1E-9);
                    Measurement_Scale.Content = "G"; //Giga
                    break;
                case "+11":
                case "+011":
                    Measurement_Value.Content = (decimal)(value * 1E-9);
                    Measurement_Scale.Content = "G"; //Giga
                    break;
                case "+12":
                case "+012":
                    Measurement_Value.Content = double.Parse(measurement.Substring(0, measurement.IndexOf("E")));
                    Measurement_Scale.Content = "T"; //Tera official
                    break;
                default:
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = double.Parse(measurement.Substring(0, measurement.IndexOf("E")));
                    break;
            }
        }

        private void Display_PSI_Double(string measurement, double value)
        {
            if (value == 0)
            {
                Measurement_Value.Content = (decimal)(value);
                Measurement_Scale.Content = "";
            }
            else if (value > -1 & value < 1)
            {
                Measurement_Value.Content = (decimal)(value * 1000);
                Measurement_Scale.Content = "m";
            }
            else if (value > 999999999)
            {
                Measurement_Value.Content = (decimal)(value / 1000000000);
                Measurement_Scale.Content = "G";
            }
            else if (value > 999999)
            {
                Measurement_Value.Content = (decimal)(value / 1000000);
                Measurement_Scale.Content = "M";
            }
            else if (value > 999)
            {
                Measurement_Value.Content = (decimal)(value / 1000);
                Measurement_Scale.Content = "k";
            }
            else
            {
                Measurement_Value.Content = (decimal)(value);
                Measurement_Scale.Content = "";
            }
        }

        private double predict_range_display(double Value)
        {
            double Range = 0;
            if (Value <= 300E-6 & Value >= -300E-6) //300u
            {
                Range = 300E-6;
            }
            else if (Value <= 0.003 & Value >= -0.003) //3m
            {
                Range = 0.003;
            }
            else if (Value <= 0.03 & Value >= -0.03) //30m
            {
                Range = 0.03;
            }
            else if (Value <= 0.3 & Value >= -0.3)  //300m
            {
                Range = 0.3;
            }
            else if (Value <= 3 & Value >= -3) //3
            {
                Range = 3;
            }
            else if (Value <= 30 & Value >= -30) //30
            {
                Range = 30;
            }
            else if (Value <= 300 & Value >= -300) //300
            {
                Range = 300;
            }
            else if (Value <= 3E3 & Value >= -3E3) //3K
            {
                Range = 3E3;
            }
            else if (Value <= 30E3 & Value >= -30E3) //30K
            {
                Range = 30E3;
            }
            else if (Value <= 300E3 & Value >= -300E3) //300K
            {
                Range = 300E3;
            }
            else if (Value <= 3E6 & Value >= -3E6) //3M
            {
                Range = 3E6;
            }
            else if (Value <= 30E6 & Value >= -30E6) //30M
            {
                Range = 30E6;
            }
            else if (Value <= 3E9 & Value >= -3E9) //3G
            {
                Range = 3E9;
            }
            return Range;
        }

        private void Display_MIN_MAX_AVG(double measurement)
        {
            if (resetMinMaxAvg == 1)
            {
                min = measurement;
                max = measurement;
                avg = 0;
                avg_count = 0;
                insert_Log("Reset MIN, MAX, AVG values.", 0);
                updateMIN(measurement);
                updateMAX(measurement);
                Interlocked.Exchange(ref resetMinMaxAvg, 0);
            }
            if (measurement < min)
            {
                updateMIN(measurement);
            }
            if (measurement > max)
            {
                updateMAX(measurement);
            }
            if (AVG_Calculate == 1)
            {
                updateAVG(measurement);
            }
        }

        private void Reset_Click_MIN_MAX_AVG(object sender, MouseButtonEventArgs e)
        {
            Interlocked.Exchange(ref resetMinMaxAvg, 1);
            insert_Log("Reset MIN, MAX, AVG command has been send.", 4);
        }

        private void updateAVG(double measurement)
        {
            avg_count += 1;
            avg = avg + (measurement - avg) / Math.Min(avg_count, avg_factor);
            if (avg == 0)
            {
                AVG_Value.Content = ((decimal)Math.Round((avg), avg_resolution)).ToString();
                AVG_Scale.Content = "";
            }
            else if (avg < 1E-6 & avg > -1E-6)
            {
                AVG_Value.Content = ((decimal)Math.Round((avg * 1E9), avg_resolution)).ToString();
                AVG_Scale.Content = "n";
            }
            else if (avg < 1 & avg > -1)
            {
                AVG_Value.Content = ((decimal)Math.Round((avg * 1000), avg_resolution)).ToString();
                AVG_Scale.Content = "m";
            }
            else if (avg < -99999999999 || avg > 99999999999)
            {
                AVG_Value.Content = "OVLD";
                AVG_Scale.Content = "";
            }
            else if (avg < -999999999 || avg > 999999999)
            {
                AVG_Value.Content = ((decimal)Math.Round((avg / 1000000000), avg_resolution)).ToString();
                AVG_Scale.Content = "G";
            }
            else if (avg < -999999 || avg > 999999)
            {
                AVG_Value.Content = ((decimal)Math.Round((avg / 1000000), avg_resolution)).ToString();
                AVG_Scale.Content = "M";
            }
            else if (avg < -999 || avg > 999)
            {
                AVG_Value.Content = ((decimal)Math.Round((avg / 1000), avg_resolution)).ToString();
                AVG_Scale.Content = "K";
            }
            else
            {
                AVG_Value.Content = ((decimal)Math.Round((avg), avg_resolution)).ToString();
                AVG_Scale.Content = "";
            }
        }

        private void updateMIN(double measurement)
        {
            min = measurement;
            if (min == 0)
            {
                MIN_Value.Content = ((decimal)(min)).ToString();
                MIN_Scale.Content = "";
            }
            else if (min < 1E-6 & min > -1E-6)
            {
                MIN_Value.Content = ((decimal)(min * 1E9)).ToString();
                MIN_Scale.Content = "n";
            }
            else if (min < 1 & min > -1)
            {
                MIN_Value.Content = ((decimal)(min * 1000)).ToString();
                MIN_Scale.Content = "m";
            }
            else if (min < -99999999999 || min > 99999999999)
            {
                MIN_Value.Content = "OVLD";
                MIN_Scale.Content = "";
            }
            else if (min < -999999999 || min > 999999999)
            {
                MIN_Value.Content = ((decimal)(min / 1000000000)).ToString();
                MIN_Scale.Content = "G";
            }
            else if (min < -999999 || min > 999999)
            {
                MIN_Value.Content = ((decimal)(min / 1000000)).ToString();
                MIN_Scale.Content = "M";
            }
            else if (min < -999 || min > 999)
            {
                MIN_Value.Content = ((decimal)(min / 1000)).ToString();
                MIN_Scale.Content = "K";
            }
            else
            {
                MIN_Value.Content = ((decimal)min).ToString();
                MIN_Scale.Content = "";
            }
        }

        private void updateMAX(double measurement)
        {
            max = measurement;
            if (max == 0)
            {
                MAX_Value.Content = ((decimal)(max)).ToString();
                MAX_Scale.Content = "";
            }
            else if (max < 1E-6 & max > -1E-6)
            {
                MAX_Value.Content = ((decimal)(max * 1E9)).ToString();
                MAX_Scale.Content = "n";
            }
            else if (max < 1 & max > -1)
            {
                MAX_Value.Content = ((decimal)(max * 1000)).ToString();
                MAX_Scale.Content = "m";
            }
            else if (max < -99999999999 || max > 99999999999)
            {
                MAX_Value.Content = "OVLD";
                MAX_Scale.Content = "";
            }
            else if (max < -999999999 || max > 999999999)
            {
                MAX_Value.Content = ((decimal)(max * 1000000000)).ToString();
                MAX_Scale.Content = "G";
            }
            else if (max < -999999 || max > 999999)
            {
                MAX_Value.Content = ((decimal)(max / 1000000)).ToString();
                MAX_Scale.Content = "M";
            }
            else if (max < -999 || max > 999)
            {
                MAX_Value.Content = ((decimal)(max / 1000)).ToString();
                MAX_Scale.Content = "K";
            }
            else
            {
                MAX_Value.Content = ((decimal)max).ToString();
                MAX_Scale.Content = "";
            }
        }

        private void HP3457ACommunicateEvent(Object source, ElapsedEventArgs e)
        {
            try
            {
                if (isUserSendCommand == true)
                {
                    Serial_WriteQueue();
                    Measurement_Type_Select();
                    unlockControls();
                    isUserSendCommand = false;
                    if (UpdateSpeed > 2000)
                    {
                        Restore_Interval();
                    }
                }

                if (DataSampling == true)
                {
                    do
                    {
                        Read_Measurement();
                        Process_Data.Start();
                    } while (isSamplingOnly == true & DataSampling == true);
                }
                if (isUpdateSpeed_Changed == true)
                {
                    isUpdateSpeed_Changed = false;
                    insert_Log("Update Speed has been set to " + (UpdateSpeed / 1000) + " seconds.", 0);
                    DataTimer.Interval = UpdateSpeed;
                }
                DataTimer.Enabled = true;

            }
            catch (Exception Ex)
            {
                this.Dispatcher.Invoke(DispatcherPriority.ContextIdle, new ThreadStart(delegate
                {
                    if (Show_COM_Error.IsChecked == true)
                    {
                        insert_Log(Ex.Message, 2);
                        insert_Log("Could not get a measurement reading.", 2);
                        insert_Log("Don't worry. Trying again.", 2);
                        insert_Log("Slow the Update Speed if warning persists.", 2);
                    }
                }));
                if (isUpdateSpeed_Changed == true)
                {
                    isUpdateSpeed_Changed = false;
                    insert_Log("Update Speed has been set to " + (UpdateSpeed / 1000) + " seconds.", 0);
                    DataTimer.Interval = UpdateSpeed;
                }
                DataTimer.Enabled = true;
            }
        }

        private void Read_Measurement()
        {
            if (isNDigit_7_Selected == false)
            {
                HP3457A.WriteLine("++read");
                string data = HP3457A.ReadLine().Trim();
                int Length = data.Length;
                if (Length <= 14 & Length >= 13)
                {
                    measurements.Add(data);
                    Total_Samples++;
                    if (saveMeasurements == true || save_to_Table == true || save_to_Graph == true || Save_to_N_Graph == true || Save_to_DateTime_Graph == true)
                    {
                        Process_Measurement_Data(data, "");
                    }
                }
                else
                {
                    Invalid_Samples++;
                }
            }
            else 
            {
                HP3457A.WriteLine("TRIG SGL");
                if (isNPLC_10_Selected == false) 
                {
                    Thread.Sleep(NPLC_100_Delay_Time); //NPLC = 100 is slow at sampling
                }
                HP3457A.WriteLine("++read");
                string Measurement_data = HP3457A.ReadLine().Trim();
                if ((Measurement_data != "+1.0000000E+38") & (Measurement_data != "-1.0000000E+38"))
                {
                    int Measurement_data_length = Measurement_data.Length;

                    HP3457A.WriteLine("RMATH HIRES");
                    HP3457A.WriteLine("++read");
                    string HIRES_data = HP3457A.ReadLine().Trim();
                    int HIRES_data_length = HIRES_data.Length;

                    HP3457A.WriteLine("RANGE?");
                    string Range_Query = HP3457A.ReadLine().Trim();
                    int Range_Query_length = Range_Query.Length;

                    if ((Measurement_data_length <= 14 & Measurement_data_length >= 13) & (HIRES_data_length <= 14 & HIRES_data_length >= 13) & (Range_Query_length <= 14 & Range_Query_length >= 13))
                    {
                        double Range = double.Parse(Range_Query);
                        double Measurement_value = ((double.Parse(Measurement_data, NumberStyles.Float)) + (double.Parse(HIRES_data, NumberStyles.Float)));

                        string data = Measurement_value.ToString("E8", CultureInfo.InvariantCulture); //Formates the data into Scientific notation

                        double partial_data = double.Parse(data.Substring(0, data.IndexOf("E"))); //Removes "E+00" Part

                        int Significant_Digit = Calculate_Significant_Digits(Math.Abs(Measurement_value), Range) + 1; //Calculates the number of significant digits

                        decimal Measurement = (decimal)SignificantDigits.Round(partial_data, Significant_Digit);

                        string Display_Data = (Measurement).ToString();
                        Display_Data = Display_Data + data.Substring(data.IndexOf("E"));
                        measurements.Add(Display_Data);
                        Total_Samples++;
                        if (show_7Digit_Info == true)
                        {
                            insert_Log(("Meas: " + Measurement_data + ", HIRES: " + HIRES_data + ", Sum: " + Measurement_value + ", Sig: " + Significant_Digit + ", Correct Sum: " + Display_Data + ", Range: " + Range + ", NDigit: " + NDigit), 5);
                        }
                        if (saveMeasurements == true || save_to_Table == true || save_to_Graph == true || Save_to_N_Graph == true || Save_to_DateTime_Graph == true)
                        {
                            string data_7th_digit = "," + Measurement_data + "," + HIRES_data + "," + Measurement_value;
                            Process_Measurement_Data(Display_Data, data_7th_digit);
                        }
                    }
                    else
                    {
                        Invalid_Samples++;
                    }
                }
                else
                {
                    measurements.Add(Measurement_data);
                }

            }
        }

        private void NPLC_100_Time_Button_Click(object sender, RoutedEventArgs e)
        {
            (bool isValid, double Value) = Text_Num(NPLC_100_Delay_Time_Text.Text, false, true);
            if (isValid == true)
            {
                if (Value <= 30 & Value > 0)
                {
                    try
                    {
                        int Value_milliSeconds = (int)(Value * 1E3);
                        Interlocked.Exchange(ref NPLC_100_Delay_Time, Value_milliSeconds);
                        insert_Log("NPLC 100 Time Delay set to " + NPLC_100_Delay_Time + " milliseconds.", 0);
                    }
                    catch (Exception) 
                    {
                        insert_Log("Failed to set NPLC 100 Time Delay, try again.", 1);
                    }
                }
                else
                {
                    insert_Log("NPLC 100 Time Delay must be > 0 and < 30 seconds.", 2);
                }
            }
            else 
            {
                insert_Log("NPLC 100 Time Delay must be a real number, only integers.", 2);
            }
        }

        private void Send_SCPI_Command_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string SCPI_Command = Send_SCPI_Command_Text.Text.Trim();
                SerialWriteQueue.Add(SCPI_Command);
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Write Command: " + SCPI_Command + " sent.", 3);
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("Something went wrong while sending write command.", 1);
            }
        }

        private void Query_SCPI_Command_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SCPI_Command = Send_SCPI_Command_Text.Text.Trim();
                SerialWriteQueue.Add("Query_Command");
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Query Command: " + SCPI_Command + " sent.", 3);
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("Something went wrong while sending query command.", 1);
            }
        }

        //converts a string into a number
        private (bool, double) Text_Num(string text, bool allowNegative, bool isInteger)
        {
            if (isInteger == true)
            {
                bool isValid = int.TryParse(text, out int value);
                if (isValid == true)
                {
                    if (allowNegative == false)
                    {
                        if (value < 0)
                        {
                            return (false, 0);
                        }
                        else
                        {
                            return (true, value);
                        }
                    }
                    else
                    {
                        return (true, value);
                    }
                }
                else
                {
                    return (false, 0);
                }
            }
            else
            {
                bool isValid = double.TryParse(text, out double value);
                if (isValid == true)
                {
                    if (allowNegative == false)
                    {
                        if (value < 0)
                        {
                            return (false, 0);
                        }
                        else
                        {
                            return (true, value);
                        }
                    }
                    else
                    {
                        return (true, value);
                    }
                }
                else
                {
                    return (false, 0);
                }
            }
        }

        //This code for calculating significant digits was taken from sigrok 3457A source code
        // https://sigrok.org/gitweb/?p=libsigrok.git;a=blob;f=src/hardware/hp-3457a/protocol.c;h=f0197290f3b6d0e4825194335309d8a1a0df8d42;hb=HEAD
        // Special thanks to Alexandru Gagniuc for this code
        private int Calculate_Significant_Digits(double measurement, double range)
        {
            int zero_digits;
            double min_full_res_reading;
            double log10_range;
            double full_res_ratio;

            log10_range = Math.Log10(range);
            min_full_res_reading = Math.Pow(10, (int)log10_range);

            if (measurement > min_full_res_reading)
            { zero_digits = 0; }
            else if (measurement == 0.0)
            {
                zero_digits = 0;
            }
            else
            {
                full_res_ratio = min_full_res_reading / measurement;
                zero_digits = (int)Math.Ceiling(Math.Log10(Math.Ceiling(full_res_ratio)));
            }
            return (NDigit - zero_digits);
        }

        private void Process_Measurement_Data(string data, string more_data)
        {
            string Date = DateTime.Now.ToString("yyyy-MM-dd h:mm:ss.fff tt");
            if (saveMeasurements == true)
            {
                switch (Selected_Measurement_type)
                {
                    case 0:
                        save_data_VDC.Add(Date + "," + data + more_data);
                        break;
                    case 1:
                        save_data_VAC.Add(Date + "," + data + more_data);
                        break;
                    case 2:
                        save_data_2Ohm.Add(Date + "," + data + more_data);
                        break;
                    case 3:
                        save_data_4Ohm.Add(Date + "," + data + more_data);
                        break;
                    case 4:
                        save_data_ADC.Add(Date + "," + data + more_data);
                        break;
                    case 5:
                        save_data_AAC.Add(Date + "," + data + more_data);
                        break;
                    case 6:
                        save_data_VDCVAC.Add(Date + "," + data + more_data);
                        break;
                    case 7:
                        save_data_ADCAAC.Add(Date + "," + data + more_data);
                        break;
                    case 8:
                        save_data_FREQ.Add(Date + "," + data);
                        break;
                    case 9:
                        save_data_PER.Add(Date + "," + data);
                        break;
                    default:
                        insert_Log("Data was not saved. Something went wrong.", 0);
                        break;
                }
            }

            if (save_to_Table == true)
            {
                try
                {
                    HP3457A_Table.Table_Data_Queue.Add(Date + "," + data + "," + Current_Measurement_Unit);
                }
                catch (Exception)
                {
                    insert_Log("Could not add data to Table Window.", 2);
                    insert_Log("This could happen if the table window was opened or closed recently.", 2);
                }
            }

            if (save_to_Graph == true)
            {
                try
                {
                    HP3457A_Graph_Window.Data_Queue.Add(Date + "," + data);
                }
                catch (Exception)
                {
                    insert_Log("Could not add data to Graph Window.", 2);
                    insert_Log("This could happen if the graph window was opened or closed recently.", 2);
                }
            }

            if (Save_to_N_Graph == true)
            {
                try
                {
                    HP3457A_N_Graph_Window.Data_Queue.Add(Date + "," + data);
                }
                catch (Exception)
                {
                    insert_Log("Could not add data to N Sample Graph Window.", 2);
                    insert_Log("This could happen if the N Sample Graph Window was opened or closed recently.", 2);
                }
            }

            if (Save_to_DateTime_Graph == true)
            {
                try
                {
                    HP3457A_DateTime_Graph_Window.Data_Queue.Add(Date + "," + data);
                }
                catch (Exception)
                {
                    insert_Log("Could not add data to DateTime Graph Window.", 2);
                    insert_Log("This could happen if the DateTime Graph Window was opened or closed recently.", 2);
                }
            }
        }

        private void Serial_WriteQueue()
        {
            while (SerialWriteQueue.Count != 0)
            {
                string WriteCommand = SerialWriteQueue.Take();
                Serial_Queue_Command_Process(WriteCommand);
            }
        }

        private void Serial_Queue_Command_Process(string Command) 
        {
            switch (Command)
            {
                case "DCV":
                case "ACV":
                case "OHM":
                case "ACDCV":
                case "DCI":
                case "ACI":
                case "ACDCI":
                case "OHMF":
                case "FREQ":
                case "PER":
                    Measurement_Range = 0;
                    HP3457A.WriteLine(Command);
                    Interlocked.Exchange(ref resetMinMaxAvg, 1);
                    break;
                case "NDIG 3":
                case "NDIG 4":
                case "NDIG 5":
                case "NDIG 6":
                    if (isNDigit_7_Selected == true)
                    {
                        isNDigit_7_Selected = false;
                        HP3457A.WriteLine("TRIG 1");
                        this.Dispatcher.Invoke(() =>
                        {
                            Trigger_Indicator(0);
                        });
                        NDigit = 7;
                        insert_Log("Trigger set to internal. 7½ resolution is disabled.", 0);
                    }
                    HP3457A.WriteLine(Command);
                    break;
                case "NDIG 7":
                    isNDigit_7_Selected = true;
                    HP3457A.WriteLine("TRIG 4");
                    HP3457A.WriteLine("NDIG 6");
                    HP3457A.WriteLine("NPLC 10");
                    NDigit = 8;
                    NPLC = 10;
                    this.Dispatcher.Invoke(() =>
                    {
                        NPLC_Indicator(4);
                        Trigger_Indicator(3);
                    });
                    isNPLC_10_Selected = true;
                    insert_Log("N Digit 7½, Trigger Hold, NPLC 10", 0);
                    break;
                case "NPLC 0.0005":
                    NDigit = 4;
                    NPLC = 0.0005;
                    HP3457A.WriteLine(Command);
                    break;
                case "NPLC 0.005":
                    NDigit = 5;
                    NPLC = 0.005;
                    HP3457A.WriteLine(Command);
                    break;
                case "NPLC 0.1":
                    NDigit = 6;
                    NPLC = 0.1;
                    HP3457A.WriteLine(Command);
                    break;
                case "NPLC 1":
                    NDigit = 7;
                    NPLC = 1;
                    HP3457A.WriteLine(Command);
                    break;
                case "NPLC 10":
                    if (isNDigit_7_Selected == true)
                    {
                        NDigit = 8;
                    }
                    else
                    {
                        NDigit = 7;
                    }
                    NPLC = 10;
                    isNPLC_10_Selected = true;
                    HP3457A.WriteLine(Command);
                    break;
                case "NPLC 100":
                    HP3457A.WriteTimeout = 20000;
                    HP3457A.ReadTimeout = 20000;
                    if (isNDigit_7_Selected == true)
                    {
                        NDigit = 8;
                    }
                    else
                    {
                        NDigit = 7;
                    }
                    NPLC = 10;
                    isNPLC_10_Selected = false;
                    HP3457A.WriteLine(Command);
                    insert_Log("When NPLC 100 is set, multimeter may ignore GPIB Write Commands.", 2);
                    insert_Log("Software will capture measurement data extremely slowly and timeout counter may go up.", 2);
                    insert_Log("NPLC 100 is not recommended.", 2);
                    break;
                case "RANGE AUTO":
                    Measurement_Range = 0;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 300E-6":
                    Measurement_Range = 0.0003;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 0.003":
                    Measurement_Range = 0.003;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 0.03":
                    Measurement_Range = 0.03;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 0.3":
                    Measurement_Range = 0.3;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 3":
                    Measurement_Range = 3;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 30":
                    Measurement_Range = 30;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 300":
                    Measurement_Range = 300;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 3E3":
                    Measurement_Range = 3000;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 30E3":
                    Measurement_Range = 30000;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 300E3":
                    Measurement_Range = 300000;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 3E6":
                    Measurement_Range = 3000000;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 30E6":
                    Measurement_Range = 30000000;
                    HP3457A.WriteLine(Command);
                    break;
                case "RANGE 3E9":
                    Measurement_Range = 3000000000;
                    HP3457A.WriteLine(Command);
                    break;
                case "NPLC?":
                case "TERM?":
                case "ID?":
                case "STB?":
                case "ERR?":
                case "AUXERR?":
                case "CALNUM?":
                case "REV?":
                case "FIXEDZ?":
                case "LFREQ?":
                case "LINE?":
                    HP3457A.WriteLine(Command);
                    Thread.Sleep(200);
                    string Message = HP3457A.ReadLine();
                    insert_Log(Message, 0);
                    break;
                case "Query_Command":
                    HP3457A.WriteLine(SCPI_Command);
                    Thread.Sleep(200);
                    string Query_Message = HP3457A.ReadLine();
                    insert_Log(Query_Message, 0);
                    break;
                case "LOCAL_EXIT":
                    Application.Current.Dispatcher.Invoke(() => { Application.Current.Shutdown(); }, DispatcherPriority.Send);
                    break;
                default:
                    HP3457A.WriteLine(Command);
                    break;
            }
        }

        private void Measurement_Type_Select()
        {
            if (Measurement_Selected == 0)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "VDC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "VDC";
                    MAX_Type.Content = "VDC";
                    AVG_Type.Content = "VDC";
                    Current_Measurement_Unit = "VDC";
                }));
            }
            else if (Measurement_Selected == 1)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "VAC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "VAC";
                    MAX_Type.Content = "VAC";
                    AVG_Type.Content = "VAC";
                    Current_Measurement_Unit = "VAC";
                }));
            }
            else if (Measurement_Selected == 2)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "Ω";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "Ω";
                    MAX_Type.Content = "Ω";
                    AVG_Type.Content = "Ω";
                    Current_Measurement_Unit = "Ω";
                }));
            }
            else if (Measurement_Selected == 3)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "Ω";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "Ω";
                    MAX_Type.Content = "Ω";
                    AVG_Type.Content = "Ω";
                    Current_Measurement_Unit = "Ω";
                }));
            }
            else if (Measurement_Selected == 4)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "ADC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "ADC";
                    MAX_Type.Content = "ADC";
                    AVG_Type.Content = "ADC";
                    Current_Measurement_Unit = "ADC";
                }));
            }
            else if (Measurement_Selected == 5)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "AAC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "ADC";
                    MAX_Type.Content = "ADC";
                    AVG_Type.Content = "ADC";
                    Current_Measurement_Unit = "ADC";
                }));
            }
            else if (Measurement_Selected == 6)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "VAC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "VAC";
                    MAX_Type.Content = "VAC";
                    AVG_Type.Content = "VAC";
                    Current_Measurement_Unit = "VAC";
                }));
            }
            else if (Measurement_Selected == 7)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "AAC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "AAC";
                    MAX_Type.Content = "AAC";
                    AVG_Type.Content = "AAC";
                    Current_Measurement_Unit = "AAC";
                }));
            }
            else if (Measurement_Selected == 8)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "Hz";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "Hz";
                    MAX_Type.Content = "Hz";
                    AVG_Type.Content = "Hz";
                    Current_Measurement_Unit = "Hz";
                }));
            }
            else if (Measurement_Selected == 9)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
                {
                    Measurement_Type.Content = "SEC";
                    Measurement_Scale.Content = "";
                    Measurement_Value.Content = "";
                    MIN_Type.Content = "SEC";
                    MAX_Type.Content = "SEC";
                    AVG_Type.Content = "SEC";
                    Current_Measurement_Unit = "SEC";
                }));
            }
        }

        //Check if user input is a number and if it is then converts it from string to double.
        private (bool, double) isNumber(string Number)
        {
            bool isNum = double.TryParse(Number, out double number);
            return (isNum, number);
        }

        //inserts message to the output log
        private void insert_Log(string Message, int Code)
        {
            string date = DateTime.Now.ToString("yyyy-MM-dd h:mm:ss tt");
            SolidColorBrush Color;
            this.Dispatcher.Invoke(DispatcherPriority.ContextIdle, new ThreadStart(delegate
            {
                if (Output_Log.Inlines.Count >= Auto_Clear_Output_Log_Count)
                {
                    Output_Log.Text = String.Empty;
                    Output_Log.Inlines.Clear();
                    Output_Log.Inlines.Add(new Run("[" + date + "]" + " " + "Output Log has been auto cleared. \n") { Foreground = Brushes.Green });
                }
            }));
            string Status = "";
            switch (Code)
            {
                case 0:
                    Status = "[Success]";
                    Color = Brushes.Green;
                    break;
                case 1:
                    Status = "[Error]";
                    Color = Brushes.Red;
                    break;
                case 2:
                    Status = "[Warning]";
                    Color = Brushes.Orange;
                    break;
                case 3:
                    Status = "";
                    Color = Brushes.Blue;
                    break;
                case 4:
                    Status = "";
                    Color = Brushes.Black;
                    break;
                case 5:
                    Status = "";
                    Color = Brushes.BlueViolet;
                    break;
                default:
                    Status = "Unknown";
                    Color = Brushes.Black;
                    break;
            }
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
            {
                Output_Log.Inlines.Add(new Run("[" + date + "]" + " " + Status + " " + Message + "\n") { Foreground = Color });
                if (AutoScroll.IsChecked == true)
                {
                    Output_Log_Scroll.ScrollToBottom();
                }
            }));
            //Saves output log to a text file
            if (saveOutputLog == true)
            {
                writeToFile("[" + date + "]" + " " + Status + " " + Message, Serial_COM_Info.folder_Directory, Serial_COM_Info.COM_Port + "_" + "Output Log.txt", true);
            }
        }

        //Writes data to a file
        private void writeToFile(string data, string filePath, string fileName, bool append)
        {
            try
            {
                using (TextWriter datatotxt = new StreamWriter(filePath + @"\" + fileName, append))
                {
                    datatotxt.WriteLine(data.Trim());
                }
            }
            catch (Exception)
            {
                saveOutputLog = false;
                SaveOutputLog.IsChecked = false;
                insert_Log("Cannot write Output Log to text file.", 1);
                insert_Log("Save Output Log option disabled.", 1);
                insert_Log("Enable it again from Data Logger Menu if you wish to try again.", 1);
            }
        }

        //------------------------Config Options-----------------------------------------------

        private void Connect_Click(object sender, RoutedEventArgs e)
        {
            if (COM_Select == null)
            {
                COM_Select = new COM_Select_Window();
                COM_Select.Closed += (a, b) => { COM_Select = null; Serial_COM_Selected(); };
                COM_Select.Owner = this;
                COM_Select.Show();
            }
            else
            {
                COM_Select.Show();
                insert_Log("COM Select Window is already open.", 2);
            }
        }

        private void Stop_Sampling_Click(object sender, RoutedEventArgs e)
        {
            if (Stop_Sampling.IsChecked == false)
            {
                DataSampling = false;
            }
            else
            {
                StartDateTime = DateTime.Now;
                DataSampling = true;
            }
            if (DataSampling == true)
            {
                insert_Log("Software is reading measurement data from multimeter.", 0);
            }
            else
            {
                insert_Log("Software will not read measurement data from multimeter.", 2);
            }
        }

        private void Sampling_Only_Click(object sender, RoutedEventArgs e)
        {
            if (Sampling_Only.IsChecked == true)
            {
                isSamplingOnly = true;
                lockControls();
            }
            else
            {
                isSamplingOnly = false;
                unlockControls();
            }
            if (isSamplingOnly == true)
            {
                insert_Log("Software will now only read measurements from the multimeter.", 2);
                insert_Log("All Write (front panel) operations are disabled.", 2);
            }
            else
            {
                insert_Log("Software will allow commands to be send to the multimeter.", 0);
                insert_Log("Sampling only mode disabled. Returned to normal mode.", 0);
            }
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        //-----------------------------------------------------------------------


        //---------------------------Data Logger--------------------------------------------

        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("explorer.exe", Serial_COM_Info.folder_Directory);
            }
            catch (Exception)
            {
                insert_Log("Cannot open test files directory.", 1);
            }
        }

        private void SaveOutputLog_Click(object sender, RoutedEventArgs e)
        {
            if (SaveOutputLog.IsChecked == true)
            {
                saveOutputLog = true;
            }
            else
            {
                saveOutputLog = false;
            }
            if (saveOutputLog == true)
            {
                insert_Log("Output Log entries will be saved to a text file.", 0);
            }
            else
            {
                insert_Log("Output Log entries will not be saved.", 2);
            }
        }

        private void ClearOutputLog_Click(object sender, RoutedEventArgs e)
        {
            Output_Log.Text = String.Empty;
            Output_Log.Inlines.Clear();
        }

        private void SaveMeasurements_Click(object sender, RoutedEventArgs e)
        {
            if (SaveMeasurements.IsChecked == true)
            {
                saveMeasurements = true;
            }
            else
            {
                saveMeasurements = false;
            }
            if (saveMeasurements == true)
            {
                insert_Log("Measurement data will be saved.", 0);
                saveMeasurements_Timer.Enabled = true;
            }
            else
            {
                insert_Log("Measurement data will not be saved.", 2);
            }
        }

        private void SaveMeasurements_Interval_5Sec_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 5000;
            insert_Log("Save Measurement Interval set to 5 seconds.", 0);
            SaveMeasurements_IntervalSelected(5);
        }

        private void SaveMeasurements_Interval_10Sec_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 10000;
            insert_Log("Save Measurement Interval set to 10 seconds.", 0);
            SaveMeasurements_IntervalSelected(10);
        }

        private void SaveMeasurements_Interval_20Sec_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 20000;
            insert_Log("Save Measurement Interval set to 20 seconds.", 0);
            SaveMeasurements_IntervalSelected(20);
        }

        private void SaveMeasurements_Interval_40Sec_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 40000;
            insert_Log("Save Measurement Interval set to 40 seconds.", 0);
            SaveMeasurements_IntervalSelected(40);
        }

        private void SaveMeasurements_Interval_1Min_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 60000;
            insert_Log("Save Measurement Interval set to 1 Minute.", 0);
            SaveMeasurements_IntervalSelected(60);
        }

        private void SaveMeasurements_Interval_4Min_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 240000;
            insert_Log("Save Measurement Interval set to 4 Minutes.", 0);
            SaveMeasurements_IntervalSelected(240);
        }

        private void SaveMeasurements_Interval_8Min_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 480000;
            insert_Log("Save Measurement Interval set to 8 Minutes.", 0);
            SaveMeasurements_IntervalSelected(480);
        }

        private void SaveMeasurements_Interval_10Min_Click(object sender, RoutedEventArgs e)
        {
            saveMeasurements_Timer.Interval = 600000;
            insert_Log("Save Measurement Interval set to 10 Minutes.", 0);
            SaveMeasurements_IntervalSelected(600);
        }

        private void SaveMeasurements_IntervalSelected(int interval)
        {
            if (interval == 5)
            {
                SaveMeasurements_Interval_5Sec.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_5Sec.IsChecked = false;
            }
            if (interval == 10)
            {
                SaveMeasurements_Interval_10Sec.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_10Sec.IsChecked = false;
            }
            if (interval == 20)
            {
                SaveMeasurements_Interval_20Sec.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_20Sec.IsChecked = false;
            }
            if (interval == 40)
            {
                SaveMeasurements_Interval_40Sec.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_40Sec.IsChecked = false;
            }
            if (interval == 60)
            {
                SaveMeasurements_Interval_1Min.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_1Min.IsChecked = false;
            }
            if (interval == 240)
            {
                SaveMeasurements_Interval_4Min.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_4Min.IsChecked = false;
            }
            if (interval == 480)
            {
                SaveMeasurements_Interval_8Min.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_8Min.IsChecked = false;
            }
            if (interval == 600)
            {
                SaveMeasurements_Interval_10Min.IsChecked = true;
            }
            else
            {
                SaveMeasurements_Interval_10Min.IsChecked = false;
            }
        }

        private void Auto_Clear_20_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Auto_Clear_Output_Log_Count, 20);
            insert_Log("Output Log will be cleared after " + Auto_Clear_Output_Log_Count + " logs are inserted into it.", 0);
            Auto_Clear_20.IsChecked = true;
            Auto_Clear_40.IsChecked = false;
            Auto_Clear_60.IsChecked = false;
        }

        private void Auto_Clear_40_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Auto_Clear_Output_Log_Count, 40);
            insert_Log("Output Log will be cleared after " + Auto_Clear_Output_Log_Count + " logs are inserted into it.", 0);
            Auto_Clear_20.IsChecked = false;
            Auto_Clear_40.IsChecked = true;
            Auto_Clear_60.IsChecked = false;
        }

        private void Auto_Clear_60_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Auto_Clear_Output_Log_Count, 60);
            insert_Log("Output Log will be cleared after " + Auto_Clear_Output_Log_Count + " logs are inserted into it.", 0);
            Auto_Clear_20.IsChecked = false;
            Auto_Clear_40.IsChecked = false;
            Auto_Clear_60.IsChecked = true;
        }

        //-----------------------------------------------------------------------

        //---------------------------------Graph Options--------------------------------------
        private void ShowMeasurementGraph_Click(object sender, RoutedEventArgs e)
        {
            if (HP3457A_Graph_Window == null)
            {
                Create_HP3457A_Graph_Window();
                ShowMeasurementGraph.IsChecked = true;
                AddDataGraph.IsChecked = true;
                save_to_Graph = true;
                Enable_AddDatatoGraph();
                insert_Log("HP3457A Graph Module has been opened.", 0);
            }
            else
            {
                ShowMeasurementGraph.IsChecked = true;
            }
        }

        private void Create_HP3457A_Graph_Window()
        {
            try
            {
                (string Measurement_Unit, string Graph_Y_Axis_Label) = MeasurementUnit_String();
                Thread Waveform_Thread = new Thread(new ThreadStart(() =>
                {
                    HP3457A_Graph_Window = new Graphing_Window(Measurement_Unit, Graph_Y_Axis_Label, "HP 3457A " + Serial_COM_Info.COM_Port);
                    HP3457A_Graph_Window.Show();
                    HP3457A_Graph_Window.Closed += Close_Graph_Event;
                    Dispatcher.Run();
                }));
                Waveform_Thread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.SetApartmentState(ApartmentState.STA);
                Waveform_Thread.IsBackground = true;
                Waveform_Thread.Start();
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("HP3457A Graph Window creation failed.", 1);
            }
        }

        private void Close_Graph_Event(object sender, EventArgs e)
        {
            HP3457A_Graph_Window.Dispatcher.InvokeShutdown();
            HP3457A_Graph_Window = null;
            Close_Graph_Module();
        }

        private void Close_Graph_Module()
        {
            this.Dispatcher.Invoke(() =>
            {
                if (HP3457A_Graph_Window == null & HP3457A_N_Graph_Window == null & HP3457A_DateTime_Graph_Window == null)
                {
                    Save_to_N_Graph = false;
                    Save_to_DateTime_Graph = false;
                    AddDataGraph.IsChecked = false;
                    insert_Log("No Graphs are opened, unchecking Add Data to Graphs option.", 2);
                }
                save_to_Graph = false;
                ShowMeasurementGraph.IsChecked = false;
                insert_Log("HP3457A Graph Module has been closed.", 0);
            });
        }

        private void Try_Graph_Reset()
        {
            try
            {
                (string Measurement_Unit, string Graph_Y_Axis_Label) = MeasurementUnit_String();
                if (HP3457A_Graph_Window != null)
                {
                    HP3457A_Graph_Window.Measurement_Unit = Measurement_Unit;
                    HP3457A_Graph_Window.Graph_Y_Axis_Label = Graph_Y_Axis_Label;
                    HP3457A_Graph_Window.Graph_Reset = true;
                }
                if (HP3457A_N_Graph_Window != null)
                {
                    HP3457A_N_Graph_Window.Measurement_Unit = Measurement_Unit;
                    HP3457A_N_Graph_Window.Graph_Y_Axis_Label = Graph_Y_Axis_Label;
                    HP3457A_N_Graph_Window.Graph_Reset = true;
                }
                if (HP3457A_DateTime_Graph_Window != null)
                {
                    HP3457A_DateTime_Graph_Window.Measurement_Unit = Measurement_Unit;
                    HP3457A_DateTime_Graph_Window.Graph_Y_Axis_Label = Graph_Y_Axis_Label;
                    HP3457A_DateTime_Graph_Window.Graph_Reset = true;
                }
            }
            catch (Exception)
            {
                insert_Log("Graph Reset may have failed, do a manual reset through the graph window.", 2);
            }
        }

        private void AddDataGraph_Click(object sender, RoutedEventArgs e)
        {
            if (AddDataGraph.IsChecked == true & HP3457A_Graph_Window != null)
            {
                save_to_Graph = true;
                insert_Log("Data will be added to Graph.", 0);
                AddDataGraph.IsChecked = true;
            }
            else
            {
                save_to_Graph = false;
            }
            if (AddDataGraph.IsChecked == true & HP3457A_N_Graph_Window != null)
            {
                Save_to_N_Graph = true;
                insert_Log("Data will be added to N Sample Waveform Graph.", 0);
                AddDataGraph.IsChecked = true;
            }
            else
            {
                Save_to_N_Graph = false;
            }
            if (AddDataGraph.IsChecked == true & HP3457A_DateTime_Graph_Window != null)
            {
                Save_to_DateTime_Graph = true;
                insert_Log("Data will be added to DateTime Graph.", 0);
                AddDataGraph.IsChecked = true;
            }
            else
            {
                Save_to_DateTime_Graph = false;
            }
            if (HP3457A_Graph_Window == null & HP3457A_N_Graph_Window == null & HP3457A_DateTime_Graph_Window == null)
            {
                save_to_Graph = false;
                Save_to_N_Graph = false;
                Save_to_DateTime_Graph = false;
                AddDataGraph.IsChecked = false;
                insert_Log("No Graphs are opened, unchecking Add Data to Graphs option.", 2);
            }
        }

        private void Enable_AddDatatoGraph()
        {
            if (HP3457A_Graph_Window != null)
            {
                save_to_Graph = true;
                AddDataGraph.IsChecked = true;
            }
            if (HP3457A_N_Graph_Window != null)
            {
                Save_to_N_Graph = true;
                AddDataGraph.IsChecked = true;
            }
            if (HP3457A_DateTime_Graph_Window != null)
            {
                Save_to_DateTime_Graph = true;
                AddDataGraph.IsChecked = true;
            }
        }

        private void N_Sample_Graph_Button_Click(object sender, RoutedEventArgs e)
        {
            (bool isValidNum, double N_Sample_Value) = Text_Num(N_Sample_Graph_Text.Text, false, true);
            if (isValidNum == true)
            {
                if (N_Sample_Value >= 10)
                {
                    if (HP3457A_N_Graph_Window == null)
                    {
                        Create_HP3457A_N_Sample_Graph_Window((int)N_Sample_Value);
                        Show_N_Sample_Graph.IsChecked = true;
                        AddDataGraph.IsChecked = true;
                        Save_to_N_Graph = true;
                        Enable_AddDatatoGraph();
                        insert_Log("HP3457A N Sample Graph Module has been opened.", 0);
                    }
                }
                else
                {
                    insert_Log("N Sample Graph Creation Value must be a positive integer greater than 10.", 2);
                }
            }
            else
            {
                insert_Log("N Sample Graph Creation Value must be a positive integer greater than 10.", 2);
            }
        }

        private void Create_HP3457A_N_Sample_Graph_Window(int N_Samples)
        {
            try
            {
                (string Measurement_Unit, string Graph_Y_Axis_Label) = MeasurementUnit_String();
                Thread Waveform_Thread = new Thread(new ThreadStart(() =>
                {
                    HP3457A_N_Graph_Window = new N_Sample_Graph_Window(N_Samples, Measurement_Unit, Graph_Y_Axis_Label, "HP 3457A " + Serial_COM_Info.COM_Port);
                    HP3457A_N_Graph_Window.Show();
                    HP3457A_N_Graph_Window.Closed += N_Sample_Close_Graph_Event;
                    Dispatcher.Run();
                }));
                Waveform_Thread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.SetApartmentState(ApartmentState.STA);
                Waveform_Thread.IsBackground = true;
                Waveform_Thread.Start();
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("HP3457A N Sample Graph Window creation failed.", 1);
            }
        }

        private void N_Sample_Close_Graph_Event(object sender, EventArgs e)
        {
            HP3457A_N_Graph_Window.Dispatcher.InvokeShutdown();
            HP3457A_N_Graph_Window = null;
            Close_N_Sample_Graph_Module();
        }

        private void Close_N_Sample_Graph_Module()
        {
            this.Dispatcher.Invoke(() =>
            {
                if (HP3457A_Graph_Window == null & HP3457A_N_Graph_Window == null & HP3457A_DateTime_Graph_Window == null)
                {
                    save_to_Graph = false;
                    Save_to_DateTime_Graph = false;
                    AddDataGraph.IsChecked = false;
                    insert_Log("No Graphs are opened, unchecking Add Data to Graphs option.", 2);
                }
                Save_to_N_Graph = false;
                Show_N_Sample_Graph.IsChecked = false;
                insert_Log("HP3457A N Sample Graph Module has been closed.", 0);
            });
        }

        private void Show_DateTime_Graph_Button_Click(object sender, RoutedEventArgs e)
        {
            if (HP3457A_DateTime_Graph_Window == null)
            {
                Create_HP3457A_DateTime_Graph_Window();
                ShowDateTimeGraph.IsChecked = true;
                AddDataGraph.IsChecked = true;
                Save_to_DateTime_Graph = true;
                Enable_AddDatatoGraph();
                insert_Log("HP3457A DateTime Graph Module has been opened.", 0);
            }
            else
            {
                ShowDateTimeGraph.IsChecked = true;
            }
        }

        private void Create_HP3457A_DateTime_Graph_Window()
        {
            try
            {
                (string Measurement_Unit, string Graph_Y_Axis_Label) = MeasurementUnit_String();
                Thread Waveform_Thread = new Thread(new ThreadStart(() =>
                {
                    HP3457A_DateTime_Graph_Window = new DateTime_Graph_Window(Measurement_Unit, Graph_Y_Axis_Label, "HP 3457A " + Serial_COM_Info.COM_Port);
                    HP3457A_DateTime_Graph_Window.Show();
                    HP3457A_DateTime_Graph_Window.Closed += Close_DateTime_Graph_Event;
                    Dispatcher.Run();
                }));
                Waveform_Thread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                Waveform_Thread.SetApartmentState(ApartmentState.STA);
                Waveform_Thread.IsBackground = true;
                Waveform_Thread.Start();
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("HP3457A Graph Window creation failed.", 1);
            }
        }

        private void Close_DateTime_Graph_Event(object sender, EventArgs e)
        {
            HP3457A_DateTime_Graph_Window.Dispatcher.InvokeShutdown();
            HP3457A_DateTime_Graph_Window = null;
            Close_DateTime_Graph_Module();
        }

        private void Close_DateTime_Graph_Module()
        {
            this.Dispatcher.Invoke(() =>
            {
                if (HP3457A_Graph_Window == null & HP3457A_N_Graph_Window == null & HP3457A_DateTime_Graph_Window == null)
                {
                    save_to_Graph = false;
                    Save_to_N_Graph = false;
                    AddDataGraph.IsChecked = false;
                    insert_Log("No Graphs are opened, unchecking Add Data to Graphs option.", 2);
                }
                Save_to_DateTime_Graph = false;
                ShowDateTimeGraph.IsChecked = false;
                insert_Log("HP3457A DateTime Graph Module has been closed.", 0);
            });
        }

        //-----------------------------------------------------------------------

        //------------------------------Table Options-----------------------------------------

        private void ShowTable_Click(object sender, RoutedEventArgs e)
        {
            if (HP3457A_Table == null)
            {
                Create_HP3457A_Table_Window();
                AddDataTable.IsChecked = true;
                ShowTable.IsChecked = true;
                save_to_Table = true;
                insert_Log("HP3457A Table Window has been opened.", 0);
            }
            else
            {
                ShowTable.IsChecked = true;
            }
        }

        private void AddDataTable_Click(object sender, RoutedEventArgs e)
        {
            if (AddDataTable.IsChecked == true & HP3457A_Table != null)
            {
                save_to_Table = true;
                insert_Log("Data will be added to the table.", 0);
                AddDataTable.IsChecked = true;
            }
            else
            {
                save_to_Table = false;
                AddDataTable.IsChecked = false;
            }
        }

        private void Create_HP3457A_Table_Window()
        {
            try
            {
                Thread Table_Thread = new Thread(new ThreadStart(() =>
                {
                    HP3457A_Table = new Measurement_Data_Table("HP 3457A " + Serial_COM_Info.COM_Port);
                    HP3457A_Table.Show();
                    HP3457A_Table.Closed += Close_Table_Event;
                    Dispatcher.Run();
                }));
                Table_Thread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Table_Thread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");
                Table_Thread.SetApartmentState(ApartmentState.STA);
                Table_Thread.IsBackground = true;
                Table_Thread.Start();
            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 1);
                insert_Log("HP3457A Table Window creation failed.", 1);
            }
        }

        private void Close_Table_Event(object sender, EventArgs e)
        {
            HP3457A_Table.Dispatcher.InvokeShutdown();
            HP3457A_Table = null;
            Close_Table_Window();
        }

        private void Close_Table_Window()
        {
            this.Dispatcher.Invoke(() =>
            {
                save_to_Table = false;
                AddDataTable.IsChecked = false;
                ShowTable.IsChecked = false;
                insert_Log("HP3457A Table Window has been closed.", 0);
            });
        }

        //-----------------------------------------------------------------------

        //----------------------------Speech Options-------------------------------------------

        private void EnableSpeech_Click(object sender, RoutedEventArgs e)
        {
            if (EnableSpeech.IsChecked == true)
            {
                Interlocked.Exchange(ref isSpeechActive, 1);
                insert_Log("The Speech Synthesizer is Enabled.", 4);
            }
            else
            {
                Interlocked.Exchange(ref isSpeechActive, 0);
                insert_Log("The Speech Synthesizer is Disabled.", 4);
            }
        }

        private void VoiceMale_Click(object sender, RoutedEventArgs e)
        {
            Voice.SelectVoiceByHints(VoiceGender.Male);
            VoiceMale.IsChecked = true;
            VoiceFemale.IsChecked = false;
            insert_Log("David will voice your measurements.", 0);
        }

        private void VoiceFemale_Click(object sender, RoutedEventArgs e)
        {
            Voice.SelectVoiceByHints(VoiceGender.Female);
            VoiceMale.IsChecked = false;
            VoiceFemale.IsChecked = true;
            insert_Log("Zira will voice your measurements.", 0);
        }

        private void VoiceSlow_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 0;
            VoiceSpeedSelected(0);
            insert_Log("Voice speed set to slow.", 4);
        }

        private void VoiceMedium_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 1;
            VoiceSpeedSelected(1);
            insert_Log("Voice speed set to medium.", 4);
        }

        private void VoiceFast_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 2;
            VoiceSpeedSelected(2);
            insert_Log("Voice speed set to fast.", 4);
        }

        private void VoiceVeryFast_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 3;
            VoiceSpeedSelected(3);
            insert_Log("Voice speed set to very fast.", 4);
        }

        private void VoiceFastest_Click(object sender, RoutedEventArgs e)
        {
            Voice.Rate = 4;
            VoiceSpeedSelected(4);
            insert_Log("Voice speed set to fastest.", 4);
        }

        private void VoiceSpeedSelected(int speed)
        {
            if (speed == 0)
            {
                VoiceSlow.IsChecked = true;
            }
            else
            {
                VoiceSlow.IsChecked = false;
            }
            if (speed == 1)
            {
                VoiceMedium.IsChecked = true;
            }
            else
            {
                VoiceMedium.IsChecked = false;
            }
            if (speed == 2)
            {
                VoiceFast.IsChecked = true;
            }
            else
            {
                VoiceFast.IsChecked = false;
            }
            if (speed == 3)
            {
                VoiceVeryFast.IsChecked = true;
            }
            else
            {
                VoiceVeryFast.IsChecked = false;
            }
            if (speed == 4)
            {
                VoiceFastest.IsChecked = true;
            }
            else
            {
                VoiceFastest.IsChecked = false;
            }
        }

        private void Voice_Volume_10_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 10;
            VoiceVolumeSelected(0);
            insert_Log("Voice volume set to 10%.", 4);
        }

        private void Voice_Volume_20_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 20;
            VoiceVolumeSelected(1);
            insert_Log("Voice volume set to 20%.", 4);
        }

        private void Voice_Volume_30_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 30;
            VoiceVolumeSelected(2);
            insert_Log("Voice volume set to 30%.", 4);
        }

        private void Voice_Volume_40_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 40;
            VoiceVolumeSelected(3);
            insert_Log("Voice volume set to 40%.", 4);
        }

        private void Voice_Volume_50_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 50;
            VoiceVolumeSelected(4);
            insert_Log("Voice volume set to 50%.", 4);
        }

        private void Voice_Volume_60_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 60;
            VoiceVolumeSelected(5);
            insert_Log("Voice volume set to 60%.", 4);
        }

        private void Voice_Volume_70_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 70;
            VoiceVolumeSelected(6);
            insert_Log("Voice volume set to 70%.", 4);
        }

        private void Voice_Volume_80_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 80;
            VoiceVolumeSelected(7);
            insert_Log("Voice volume set to 80%.", 4);
        }

        private void Voice_Volume_90_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 90;
            VoiceVolumeSelected(8);
            insert_Log("Voice volume set to 90%.", 4);
        }

        private void Voice_Volume_100_Click(object sender, RoutedEventArgs e)
        {
            Voice.Volume = 100;
            VoiceVolumeSelected(9);
            insert_Log("Voice volume set to 100%.", 4);
        }

        private void VoiceVolumeSelected(int volume)
        {
            if (volume == 0)
            {
                Voice_Volume_10.IsChecked = true;
            }
            else
            {
                Voice_Volume_10.IsChecked = false;
            }
            if (volume == 1)
            {
                Voice_Volume_20.IsChecked = true;
            }
            else
            {
                Voice_Volume_20.IsChecked = false;
            }
            if (volume == 2)
            {
                Voice_Volume_30.IsChecked = true;
            }
            else
            {
                Voice_Volume_30.IsChecked = false;
            }
            if (volume == 3)
            {
                Voice_Volume_40.IsChecked = true;
            }
            else
            {
                Voice_Volume_40.IsChecked = false;
            }
            if (volume == 4)
            {
                Voice_Volume_50.IsChecked = true;
            }
            else
            {
                Voice_Volume_50.IsChecked = false;
            }
            if (volume == 5)
            {
                Voice_Volume_60.IsChecked = true;
            }
            else
            {
                Voice_Volume_60.IsChecked = false;
            }
            if (volume == 6)
            {
                Voice_Volume_70.IsChecked = true;
            }
            else
            {
                Voice_Volume_70.IsChecked = false;
            }
            if (volume == 7)
            {
                Voice_Volume_80.IsChecked = true;
            }
            else
            {
                Voice_Volume_80.IsChecked = false;
            }
            if (volume == 8)
            {
                Voice_Volume_90.IsChecked = true;
            }
            else
            {
                Voice_Volume_90.IsChecked = false;
            }
            if (volume == 9)
            {
                Voice_Volume_100.IsChecked = true;
            }
            else
            {
                Voice_Volume_100.IsChecked = false;
            }
        }

        //-----------------------------------------------------------------------

        //----------------------------About Options-------------------------------------------

        private void DeviceSupport_Click(object sender, RoutedEventArgs e)
        {
            insert_Log("This software was created for HP 3457A.", 4);
            insert_Log("You will need an AR488 Arduino GPIB adapter.", 4);
        }

        private void Credits_Click(object sender, RoutedEventArgs e)
        {
            insert_Log("Created by Niravk Patel.", 4);
            insert_Log("Email: niravkp97@gmail.com", 4);
            insert_Log("This program was created using C# WPF .Net Framework 4.7.2", 4);
            insert_Log("Supports Windows 10, 8, 8.1, and 7", 4);
        }

        //-----------------------------------------------------------------------

        //--------------------------Measurements Options---------------------------------------------


        //------------------------------Main Measurements---------------------------------------------
        private void VDC_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(0);
            Range_Tab_Selector(0);
            insert_Log("DCV Measurement Selected.", 3);
            VDC_Range_Indicator(0); //Auto
            SerialWriteQueue.Add("DCV");
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
        }

        private void VAC_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(1);
            Range_Tab_Selector(1);
            insert_Log("ACV Measurement Selected.", 3);
            VAC_Range_Indicator(0); //Auto
            SerialWriteQueue.Add("ACV");
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
        }

        private void TwoOhms_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(2);
            Range_Tab_Selector(2);
            insert_Log("2 Wire Ohms Measurement Selected.", 3);
            Ohms_Range_Indicator(0); //Auto
            SerialWriteQueue.Add("OHM");
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
            if (isNDigit_7_Selected == true) 
            {
                insert_Log("7½ is not available for the extended Ohms Range, 300MΩ ~ 3GΩ", 2);
            }
        }

        private void FourOhms_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(3);
            Range_Tab_Selector(2);
            insert_Log("4 Wire Ohms Measurement Selected.", 3);
            Ohms_Range_Indicator(0); //Auto
            SerialWriteQueue.Add("OHMF");
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
            if (isNDigit_7_Selected == true)
            {
                insert_Log("7½ is not available for the extended Ohms Range, 300MΩ ~ 3GΩ", 2);
            }
        }

        private void ADC_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(4);
            Range_Tab_Selector(3);
            insert_Log("DCI Measurement Selected.", 3);
            ADC_Range_Indicator(0); //Auto
            SerialWriteQueue.Add("DCI");
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
        }

        private void AAC_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(5);
            Range_Tab_Selector(4);
            insert_Log("ACI Measurement Selected.", 3);
            AAC_Range_Indicator(0); //Auto
            SerialWriteQueue.Add("ACI");
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
        }

        private void ACDCV_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(6);
            Range_Tab_Selector(5);
            insert_Log("ACDCV Measurement Selected.", 3);
            ACDCV_Range_Indicator(0); //Auto
            SerialWriteQueue.Add("ACDCV");
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
        }

        private void ACDCI_Button_Click(object sender, RoutedEventArgs e)
        {
            MesurementSelector(7);
            Range_Tab_Selector(6);
            insert_Log("ACDCI Measurement Selected.", 3);
            ACDCI_Range_Indicator(0); //Auto
            SerialWriteQueue.Add("ACDCI");
            lockControls();
            isUserSendCommand = true;
            Try_Graph_Reset();
            Speedup_Interval();
        }

        private void FREQ_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                MesurementSelector(8);
                Range_Tab_Selector(7);
                insert_Log("FREQ Measurement Selected.", 3);
                SerialWriteQueue.Add("FREQ");
                lockControls();
                isUserSendCommand = true;
                Try_Graph_Reset();
                Speedup_Interval();
            }
            else 
            {
                insert_Log("FREQ Measurement cannot be selected while N Digit 7½ is selected.", 2);
            }
        }

        private void PER_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                MesurementSelector(9);
                Range_Tab_Selector(7);
                insert_Log("PER Measurement Selected.", 3);
                SerialWriteQueue.Add("PER");
                lockControls();
                isUserSendCommand = true;
                Try_Graph_Reset();
                Speedup_Interval();
            }
            else 
            {
                insert_Log("PER Measurement cannot be selected while N Digit 7½ is selected.", 2);
            }
        }

        private void Range_Tab_Selector(int RangeType)
        {
            if (RangeType == 0) //VDC
            {
                VDC_Tab.IsSelected = true;
            }
            else
            {
                VDC_Tab.IsSelected = false;
            }
            if (RangeType == 1) //VAC
            {
                VAC_Tab.IsSelected = true;
            }
            else
            {
                VAC_Tab.IsSelected = false;
            }
            if (RangeType == 2) //Ohms
            {
                Ohms_Tab.IsSelected = true;
            }
            else
            {
                Ohms_Tab.IsSelected = false;
            }
            if (RangeType == 3) //ADC
            {
                ADC_Tab.IsSelected = true;
            }
            else
            {
                ADC_Tab.IsSelected = false;
            }
            if (RangeType == 4) //AAC
            {
                AAC_Tab.IsSelected = true;
            }
            else
            {
                AAC_Tab.IsSelected = false;
            }
            if (RangeType == 5) //ACDCV
            {
                ACDCV_Tab.IsSelected = true;
            }
            else
            {
                ACDCV_Tab.IsSelected = false;
            }
            if (RangeType == 6) //ACDCI
            {
                ACDCI_Tab.IsSelected = true;
            }
            else
            {
                ACDCI_Tab.IsSelected = false;
            }
            if (RangeType == 7) //FSOURCE
            {
                FSOURCE_Tab.IsSelected = true;
            }
            else
            {
                FSOURCE_Tab.IsSelected = false;
            }

        }

        private void MesurementSelector(int MeasurementChoice)
        {
            if (MeasurementChoice == 0) //VDC
            {
                VDC_Border.BorderBrush = Selected;
                Measurement_Selected = 0;
                Interlocked.Exchange(ref Selected_Measurement_type, 0);
            }
            else
            {
                VDC_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 1) //VAC
            {
                VAC_Border.BorderBrush = Selected;
                Measurement_Selected = 1;
                Interlocked.Exchange(ref Selected_Measurement_type, 1);
            }
            else
            {
                VAC_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 2) //2 Wire Ohms
            {
                TwoOhms_Border.BorderBrush = Selected;
                Measurement_Selected = 2;
                Interlocked.Exchange(ref Selected_Measurement_type, 2);
            }
            else
            {
                TwoOhms_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 3) //4 Wire Ohms
            {
                FourOhms_Border.BorderBrush = Selected;
                Measurement_Selected = 3;
                Interlocked.Exchange(ref Selected_Measurement_type, 3);
            }
            else
            {
                FourOhms_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 4) //ADC
            {
                ADC_Border.BorderBrush = Selected;
                Measurement_Selected = 4;
                Interlocked.Exchange(ref Selected_Measurement_type, 4);
            }
            else
            {
                ADC_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 5) //AAC
            {
                AAC_Border.BorderBrush = Selected;
                Measurement_Selected = 5;
                Interlocked.Exchange(ref Selected_Measurement_type, 5);
            }
            else
            {
                AAC_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 6) //ACDCV
            {
                ACDCV_Border.BorderBrush = Selected;
                Measurement_Selected = 6;
                Interlocked.Exchange(ref Selected_Measurement_type, 6);
            }
            else
            {
                ACDCV_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 7) //ACDCI
            {
                ACDCI_Border.BorderBrush = Selected;
                Measurement_Selected = 7;
                Interlocked.Exchange(ref Selected_Measurement_type, 7);
            }
            else
            {
                ACDCI_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 8) //FREQ
            {
                FREQ_Border.BorderBrush = Selected;
                Measurement_Selected = 8;
                Interlocked.Exchange(ref Selected_Measurement_type, 8);
            }
            else
            {
                FREQ_Border.BorderBrush = Deselected;
            }
            if (MeasurementChoice == 9) //PER
            {
                PER_Border.BorderBrush = Selected;
                Measurement_Selected = 9;
                Interlocked.Exchange(ref Selected_Measurement_type, 9);
            }
            else
            {
                PER_Border.BorderBrush = Deselected;
            }
        }

        //------------------------------Main Measurements---------------------------------------------

        //----------------------------------Measurements Ranges-------------------------------------

        //------------------------------------VDC Start-----------------------------------------------

        private void VDC_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VDC Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set VDC Range when VDC Measurement is not selected.", 2);
            }
        }

        private void VDC_30mV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VDC Range set to 30mV.", 5);
            }
            else
            {
                insert_Log("Cannot set VDC Range when VDC Measurement is not selected.", 2);
            }
        }

        private void VDC_300mV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VDC Range set to 300mV.", 5);
            }
            else
            {
                insert_Log("Cannot set VDC Range when VDC Measurement is not selected.", 2);
            }
        }

        private void VDC_3V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VDC Range set to 3V.", 5);
            }
            else
            {
                insert_Log("Cannot set VDC Range when VDC Measurement is not selected.", 2);
            }
        }

        private void VDC_30V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(4));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VDC Range set to 30V.", 5);
            }
            else
            {
                insert_Log("Cannot set VDC Range when VDC Measurement is not selected.", 2);
            }
        }

        private void VDC_300V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 0)
            {
                SerialWriteQueue.Add(VDC_Range_Indicator(5));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VDC Range set to 300V.", 5);
            }
            else
            {
                insert_Log("Cannot set VDC Range when VDC Measurement is not selected.", 2);
            }
        }

        private string VDC_Range_Indicator(int Range)
        {
            string RangeCommand = "RANGE AUTO";
            if (Range == 0)
            {
                VDC_Auto_Border.BorderBrush = Selected;
                RangeCommand = "RANGE AUTO";
            }
            else
            {
                VDC_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 1)
            {
                VDC_30mV_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.03";
            }
            else
            {
                VDC_30mV_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                VDC_300mV_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.3";
            }
            else
            {
                VDC_300mV_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                VDC_3V_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 3";
            }
            else
            {
                VDC_3V_Border.BorderBrush = Deselected;
            }
            if (Range == 4)
            {
                VDC_30V_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 30";
            }
            else
            {
                VDC_30V_Border.BorderBrush = Deselected;
            }
            if (Range == 5)
            {
                VDC_300V_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 300";
            }
            else
            {
                VDC_300V_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //------------------------------------VDC END-----------------------------------------------

        //------------------------VAC Range-----------------------------------------------

        private void VAC_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VAC Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set VAC Range when VAC Measurement is not selected.", 2);
            }
        }

        private void VAC_30mV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VAC Range set to 30mV.", 5);
            }
            else
            {
                insert_Log("Cannot set VAC Range when VAC Measurement is not selected.", 2);
            }
        }

        private void VAC_300mV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VAC Range set to 300mV.", 5);
            }
            else
            {
                insert_Log("Cannot set VAC Range when VAC Measurement is not selected.", 2);
            }
        }

        private void VAC_3V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VAC Range set to 3V.", 5);
            }
            else
            {
                insert_Log("Cannot set VAC Range when VAC Measurement is not selected.", 2);
            }
        }

        private void VAC_30V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(4));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VAC Range set to 30V.", 5);
            }
            else
            {
                insert_Log("Cannot set VAC Range when VAC Measurement is not selected.", 2);
            }
        }

        private void VAC_300V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 1)
            {
                SerialWriteQueue.Add(VAC_Range_Indicator(5));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("VAC Range set to 300V.", 5);
            }
            else
            {
                insert_Log("Cannot set VAC Range when VAC Measurement is not selected.", 2);
            }
        }

        private string VAC_Range_Indicator(int Range)
        {
            string RangeCommand = "RANGE AUTO";
            if (Range == 0)
            {
                VAC_Auto_Border.BorderBrush = Selected;
                RangeCommand = "RANGE AUTO";
            }
            else
            {
                VAC_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 1)
            {
                VAC_30mV_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.03";
            }
            else
            {
                VAC_30mV_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                VAC_300mV_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.3";
            }
            else
            {
                VAC_300mV_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                VAC_3V_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 3";
            }
            else
            {
                VAC_3V_Border.BorderBrush = Deselected;
            }
            if (Range == 4)
            {
                VAC_30V_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 30";
            }
            else
            {
                VAC_30V_Border.BorderBrush = Deselected;
            }
            if (Range == 5)
            {
                VAC_300V_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 300";
            }
            else
            {
                VAC_300V_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //------------------------------------VAC-----------------------------------------------

        //-------------------------Ohms Range----------------------------------------------

        private void Ohms_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_30_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 30Ω.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_300_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 300Ω.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_3K_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 3KΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_30K_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(4));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 30KΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_300K_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(5));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 300KΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_3M_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(6));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 3MΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_30M_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(7));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 30MΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private void Ohms_3G_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 2 || Measurement_Selected == 3)
            {
                SerialWriteQueue.Add(Ohms_Range_Indicator(8));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Ohms Range set to 3GΩ.", 5);
            }
            else
            {
                insert_Log("Cannot set Ω Range when 2/4 Wire Ω Measurement is not selected.", 2);
            }
        }

        private string Ohms_Range_Indicator(int Range)
        {
            string RangeCommand = "RANGE AUTO";
            if (Range == 0)
            {
                Ohms_Auto_Border.BorderBrush = Selected;
                RangeCommand = "RANGE AUTO";
            }
            else
            {
                Ohms_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 1)
            {
                Ohms_30_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 30";
            }
            else
            {
                Ohms_30_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                Ohms_300_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 300";
            }
            else
            {
                Ohms_300_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                Ohms_3K_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 3E3";
            }
            else
            {
                Ohms_3K_Border.BorderBrush = Deselected;
            }
            if (Range == 4)
            {
                Ohms_30K_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 30E3";
            }
            else
            {
                Ohms_30K_Border.BorderBrush = Deselected;
            }
            if (Range == 5)
            {
                Ohms_300K_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 300E3";
            }
            else
            {
                Ohms_300K_Border.BorderBrush = Deselected;
            }
            if (Range == 6)
            {
                Ohms_3M_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 3E6";
            }
            else
            {
                Ohms_3M_Border.BorderBrush = Deselected;
            }
            if (Range == 7)
            {
                Ohms_30M_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 30E6";
            }
            else
            {
                Ohms_30M_Border.BorderBrush = Deselected;
            }
            if (Range == 8)
            {
                Ohms_3G_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 3E9";
            }
            else
            {
                Ohms_3G_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //-----------------------------------------------------------------------

        //------------------------ACDCV Range-----------------------------------------------

        private void ACDCV_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 6)
            {
                SerialWriteQueue.Add(ACDCV_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCV Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCV Range when ACDCV Measurement is not selected.", 2);
            }
        }

        private void ACDCV_30mV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 6)
            {
                SerialWriteQueue.Add(ACDCV_Range_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCV Range set to 30mV.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCV Range when ACDCV Measurement is not selected.", 2);
            }
        }

        private void ACDCV_300mV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 6)
            {
                SerialWriteQueue.Add(ACDCV_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCV Range set to 300mV.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCV Range when ACDCV Measurement is not selected.", 2);
            }
        }

        private void ACDCV_3V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 6)
            {
                SerialWriteQueue.Add(ACDCV_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCV Range set to 3V.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCV Range when ACDCV Measurement is not selected.", 2);
            }
        }

        private void ACDCV_30V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 6)
            {
                SerialWriteQueue.Add(ACDCV_Range_Indicator(4));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCV Range set to 30V.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCV Range when ACDCV Measurement is not selected.", 2);
            }
        }

        private void ACDCV_300V_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 6)
            {
                SerialWriteQueue.Add(ACDCV_Range_Indicator(5));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCV Range set to 300V.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCV Range when ACDCV Measurement is not selected.", 2);
            }
        }

        private string ACDCV_Range_Indicator(int Range)
        {
            string RangeCommand = "RANGE AUTO";
            if (Range == 0)
            {
                ACDCV_Auto_Border.BorderBrush = Selected;
                RangeCommand = "RANGE AUTO";
            }
            else
            {
                ACDCV_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 1)
            {
                ACDCV_30mV_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.03";
            }
            else
            {
                ACDCV_30mV_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                ACDCV_300mV_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.3";
            }
            else
            {
                ACDCV_300mV_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                ACDCV_3V_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 3";
            }
            else
            {
                ACDCV_3V_Border.BorderBrush = Deselected;
            }
            if (Range == 4)
            {
                ACDCV_30V_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 30";
            }
            else
            {
                ACDCV_30V_Border.BorderBrush = Deselected;
            }
            if (Range == 5)
            {
                ACDCV_300V_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 300";
            }
            else
            {
                ACDCV_300V_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //------------------------------------ACDCV-----------------------------------------------

        //-------------------------ADC Range----------------------------------------------

        private void ADC_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 4)
            {
                SerialWriteQueue.Add(ADC_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ADC Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set ADC Range when ADC Measurement is not selected.", 2);
            }
        }

        private void ADC_300u_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 4)
            {
                SerialWriteQueue.Add(ADC_Range_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ADC Range set to 300uA.", 5);
            }
            else
            {
                insert_Log("Cannot set ADC Range when ADC Measurement is not selected.", 2);
            }
        }

        private void ADC_3m_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 4)
            {
                SerialWriteQueue.Add(ADC_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ADC Range set to 3mA.", 5);
            }
            else
            {
                insert_Log("Cannot set ADC Range when ADC Measurement is not selected.", 2);
            }
        }

        private void ADC_30m_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 4)
            {
                SerialWriteQueue.Add(ADC_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ADC Range set to 30mA.", 5);
            }
            else
            {
                insert_Log("Cannot set ADC Range when ADC Measurement is not selected.", 2);
            }
        }

        private void ADC_300m_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 4)
            {
                SerialWriteQueue.Add(ADC_Range_Indicator(4));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ADC Range set to 300mA.", 5);
            }
            else
            {
                insert_Log("Cannot set ADC Range when ADC Measurement is not selected.", 2);
            }
        }

        private void ADC_1_5_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 4)
            {
                SerialWriteQueue.Add(ADC_Range_Indicator(5));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ADC Range set to 1.5A.", 5);
            }
            else
            {
                insert_Log("Cannot set ADC Range when ADC Measurement is not selected.", 2);
            }
        }

        private string ADC_Range_Indicator(int Range)
        {
            string RangeCommand = "RANGE AUTO";
            if (Range == 0)
            {
                ADC_Auto_Border.BorderBrush = Selected;
                RangeCommand = "RANGE AUTO";
            }
            else
            {
                ADC_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 1)
            {
                ADC_300u_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 300E-6";
            }
            else
            {
                ADC_300u_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                ADC_3m_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.003";
            }
            else
            {
                ADC_3m_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                ADC_30m_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.03";
            }
            else
            {
                ADC_30m_Border.BorderBrush = Deselected;
            }
            if (Range == 4)
            {
                ADC_300m_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.3";
            }
            else
            {
                ADC_300m_Border.BorderBrush = Deselected;
            }
            if (Range == 5)
            {
                ADC_1_5A_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 3";
            }
            else
            {
                ADC_1_5A_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //-----------------------------------------------------------------------

        //----------------------------AAC Range-------------------------------------------

        private void AAC_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(AAC_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("AAC Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set AAC Range when AAC Measurement is not selected.", 2);
            }
        }

        private void AAC_30m_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(AAC_Range_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("AAC Range set to 30mA.", 5);
            }
            else
            {
                insert_Log("Cannot set AAC Range when AAC Measurement is not selected.", 2);
            }
        }

        private void AAC_300m_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(AAC_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("AAC Range set to 300mA.", 5);
            }
            else
            {
                insert_Log("Cannot set AAC Range when AAC Measurement is not selected.", 2);
            }
        }

        private void AAC_1_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 5)
            {
                SerialWriteQueue.Add(AAC_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("AAC Range set to 1A.", 5);
            }
            else
            {
                insert_Log("Cannot set AAC Range when AAC Measurement is not selected.", 2);
            }
        }

        private string AAC_Range_Indicator(int Range)
        {
            string RangeCommand = "RANGE AUTO";
            if (Range == 0)
            {
                AAC_Auto_Border.BorderBrush = Selected;
                RangeCommand = "RANGE AUTO";
            }
            else
            {
                AAC_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 1)
            {
                AAC_30m_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.03";
            }
            else
            {
                AAC_30m_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                AAC_300m_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.3";
            }
            else
            {
                AAC_300m_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                AAC_1_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 3";
            }
            else
            {
                AAC_1_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //-----------------------------------------------------------------------

        //----------------------------AAC Range-------------------------------------------

        private void ACDCI_Auto_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 7)
            {
                SerialWriteQueue.Add(ACDCI_Range_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCI Range set to Auto.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCI Range when ACDCI Measurement is not selected.", 2);
            }
        }

        private void ACDCI_30m_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 7)
            {
                SerialWriteQueue.Add(ACDCI_Range_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCI Range set to 30mA.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCI Range when ACDCI Measurement is not selected.", 2);
            }
        }

        private void ACDCI_300m_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 7)
            {
                SerialWriteQueue.Add(ACDCI_Range_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCI Range set to 300mA.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCI Range when ACDCI Measurement is not selected.", 2);
            }
        }

        private void ACDCI_1_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 7)
            {
                SerialWriteQueue.Add(ACDCI_Range_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("ACDCI Range set to 1A.", 5);
            }
            else
            {
                insert_Log("Cannot set ACDCI Range when ACDCI Measurement is not selected.", 2);
            }
        }

        private string ACDCI_Range_Indicator(int Range)
        {
            string RangeCommand = "RANGE AUTO";
            if (Range == 0)
            {
                ACDCI_Auto_Border.BorderBrush = Selected;
                RangeCommand = "RANGE AUTO";
            }
            else
            {
                ACDCI_Auto_Border.BorderBrush = Deselected;
            }
            if (Range == 1)
            {
                ACDCI_30m_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.03";
            }
            else
            {
                ACDCI_30m_Border.BorderBrush = Deselected;
            }
            if (Range == 2)
            {
                ACDCI_300m_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 0.3";
            }
            else
            {
                ACDCI_300m_Border.BorderBrush = Deselected;
            }
            if (Range == 3)
            {
                ACDCI_1_Border.BorderBrush = Selected;
                RangeCommand = "RANGE 3";
            }
            else
            {
                ACDCI_1_Border.BorderBrush = Deselected;
            }
            return RangeCommand;
        }

        //-----------------------------------------------------------------------

        //----------------------------FSOURCE-------------------------------------------

        private void FSOURCE_ACV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                SerialWriteQueue.Add(FSOURCE_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("FSOURCE set to ACV.", 5);
            }
            else
            {
                insert_Log("Cannot set FSOURCE when FREQ/PERIOD is not selected.", 2);
            }
        }

        private void FSOURCE_ACDCV_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                SerialWriteQueue.Add(FSOURCE_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("FSOURCE set to ACDCV.", 5);
            }
            else
            {
                insert_Log("Cannot set FSOURCE when FREQ/PERIOD is not selected.", 2);
            }
        }

        private void FSOURCE_ACI_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                SerialWriteQueue.Add(FSOURCE_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("FSOURCE set to ACI.", 5);
            }
            else
            {
                insert_Log("Cannot set FSOURCE when FREQ/PERIOD is not selected.", 2);
            }
        }

        private void FSOURCE_ACDCI_Button_Click(object sender, RoutedEventArgs e)
        {
            if (Measurement_Selected == 8 || Measurement_Selected == 9)
            {
                SerialWriteQueue.Add(FSOURCE_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("FSOURCE set to ACDCI.", 5);
            }
            else
            {
                insert_Log("Cannot set FSOURCE when FREQ/PERIOD is not selected.", 2);
            }
        }

        private string FSOURCE_Indicator(int Select)
        {
            string FSOURCE_Select = "FSOURCE ACV";
            if (Select == 0)
            {
                FSOURCE_ACV_Border.BorderBrush = Selected;
                FSOURCE_Select = "FSOURCE ACV";
            }
            else
            {
                FSOURCE_ACV_Border.BorderBrush = Deselected;
            }
            if (Select == 1)
            {
                FSOURCE_ACDCV_Border.BorderBrush = Selected;
                FSOURCE_Select = "FSOURCE ACDCV";
            }
            else
            {
                FSOURCE_ACDCV_Border.BorderBrush = Deselected;
            }
            if (Select == 2)
            {
                FSOURCE_ACI_Border.BorderBrush = Selected;
                FSOURCE_Select = "FSOURCE ACI";
            }
            else
            {
                FSOURCE_ACI_Border.BorderBrush = Deselected;
            }
            if (Select == 3)
            {
                FSOURCE_ACDCI_Border.BorderBrush = Selected;
                FSOURCE_Select = "FSOURCE ACDCI";
            }
            else
            {
                FSOURCE_ACDCI_Border.BorderBrush = Deselected;
            }
            return FSOURCE_Select;
        }

        //-----------------------------------------------------------------------

        //----------------------------------Measurements Ranges-------------------------------------

        //----------------------------------N Digits------------------------------------------------------
        private void NDIGIT_3_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NDIGIT_Indicator(0));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("N Digit set to 3½.", 5);
        }

        private void NDIGIT_4_Button_Click(object sender, RoutedEventArgs e)
        {

            SerialWriteQueue.Add(NDIGIT_Indicator(1));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("N Digit set to 4½.", 5);

        }

        private void NDIGIT_5_Button_Click(object sender, RoutedEventArgs e)
        {

            SerialWriteQueue.Add(NDIGIT_Indicator(2));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("N Digit set to 5½.", 5);

        }

        private void NDIGIT_6_Button_Click(object sender, RoutedEventArgs e)
        {

            SerialWriteQueue.Add(NDIGIT_Indicator(3));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("N Digit set to 6½.", 5);

        }

        private void NDIGIT_7_Button_Click(object sender, RoutedEventArgs e)
        {
            if ((Measurement_Selected != 8) & (Measurement_Selected != 9))
            {
                SerialWriteQueue.Add(NDIGIT_Indicator(4));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("N Digit set to 7½.", 5);
            }
            else 
            {
                insert_Log("N Digit cannot be set to 7½ while FREQ or PER is enabled.", 2);
            }

        }

        private string NDIGIT_Indicator(int Select)
        {
            string NDIGIT_Select = "NDIG 5";
            if (Select == 0)
            {
                Digit_Three_Border.BorderBrush = Selected;
                NDIGIT_Select = "NDIG 3";
            }
            else
            {
                Digit_Three_Border.BorderBrush = Deselected;
            }
            if (Select == 1)
            {
                Digit_Four_Border.BorderBrush = Selected;
                NDIGIT_Select = "NDIG 4";
            }
            else
            {
                Digit_Four_Border.BorderBrush = Deselected;
            }
            if (Select == 2)
            {
                Digit_Five_Border.BorderBrush = Selected;
                NDIGIT_Select = "NDIG 5";
            }
            else
            {
                Digit_Five_Border.BorderBrush = Deselected;
            }
            if (Select == 3)
            {
                Digit_Six_Border.BorderBrush = Selected;
                NDIGIT_Select = "NDIG 6";
            }
            else
            {
                Digit_Six_Border.BorderBrush = Deselected;
            }
            if (Select == 4)
            {
                Digit_Seven_Border.BorderBrush = Selected;
                NDIGIT_Select = "NDIG 7";
            }
            else
            {
                Digit_Seven_Border.BorderBrush = Deselected;
            }
            return NDIGIT_Select;
        }
        //----------------------------------N Digits------------------------------------------------------

        //----------------------------------NPLC------------------------------------------------------
        private void NPLC_0005_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                SerialWriteQueue.Add(NPLC_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("NPLC set to 0.0005", 5);
                insert_Log("Resolution is 3½", 5);
            }
            else
            {
                insert_Log("Cannot set NPLC to 0.0005 while N Digit is set to 7½", 2);
            }
        }

        private void NPLC_005_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                SerialWriteQueue.Add(NPLC_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("NPLC set to 0.005", 5);
                insert_Log("Resolution is 4½", 5);
            }
            else
            {
                insert_Log("Cannot set NPLC to 0.0005 while N Digit is set to 7½", 2);
            }
        }

        private void NPLC_01_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                SerialWriteQueue.Add(NPLC_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("NPLC set to 0.1", 5);
                insert_Log("Resolution is 5½", 5);
            }
            else
            {
                insert_Log("Cannot set NPLC to 0.0005 while N Digit is set to 7½", 2);
            }
        }

        private void NPLC_1_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                SerialWriteQueue.Add(NPLC_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("NPLC set to 1", 5);
                insert_Log("Resolution is 6½", 5);
            }
            else
            {
                insert_Log("Cannot set NPLC to 0.0005 while N Digit is set to 7½", 2);
            }
        }

        private void NPLC_10_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NPLC_Indicator(4));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("NPLC set to 10", 5);
            insert_Log("Resolution is 6½", 5);
            insert_Log("7½ Resolution is available to be selected from N Digit tab.", 5);
        }

        private void NPLC_100_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(NPLC_Indicator(5));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("NPLC set to 100", 5);
            insert_Log("Resolution is 6½", 5);
            insert_Log("7½ Resolution is available to be selected from N Digit tab.", 5);
        }

        private void NPLC_Query_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("NPLC?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("NPLC Query Command Send.", 5);
        }

        private string NPLC_Indicator(int Select)
        {
            string NPLC_Select = "NPLC 10";
            if (Select == 0)
            {
                NPLC_0005_Border.BorderBrush = Selected;
                NPLC_Select = "NPLC 0.0005"; //3.5
            }
            else
            {
                NPLC_0005_Border.BorderBrush = Deselected;
            }
            if (Select == 1)
            {
                NPLC_005_Border.BorderBrush = Selected;
                NPLC_Select = "NPLC 0.005"; //4.5
            }
            else
            {
                NPLC_005_Border.BorderBrush = Deselected;
            }
            if (Select == 2)
            {
                NPLC_01_Border.BorderBrush = Selected;
                NPLC_Select = "NPLC 0.1"; //5.5
            }
            else
            {
                NPLC_01_Border.BorderBrush = Deselected;
            }
            if (Select == 3)
            {
                NPLC_1_Border.BorderBrush = Selected;
                NPLC_Select = "NPLC 1"; //6.5
            }
            else
            {
                NPLC_1_Border.BorderBrush = Deselected;
            }
            if (Select == 4)
            {
                NPLC_10_Border.BorderBrush = Selected;
                NPLC_Select = "NPLC 10"; //6.5
            }
            else
            {
                NPLC_10_Border.BorderBrush = Deselected;
            }
            if (Select == 5)
            {
                NPLC_100_Border.BorderBrush = Selected;
                NPLC_Select = "NPLC 100"; //6.5
            }
            else
            {
                NPLC_100_Border.BorderBrush = Deselected;
            }
            return NPLC_Select;
        }
        //----------------------------------NPLC------------------------------------------------------

        //----------------------------------Auto Zero------------------------------------------------------
        private void AutoZero_On_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(AutoZero_Indicator(0));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("AutoZero set to On.", 5);
        }

        private void AutoZero_Off_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(AutoZero_Indicator(1));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("AutoZero set to Off.", 5);
        }

        private string AutoZero_Indicator(int Select)
        {
            string AutoZero_Select = "AZERO 1";
            if (Select == 0)
            {
                AutoZero_On_Border.BorderBrush = Selected;
                AutoZero_Select = "AZERO 1";
            }
            else
            {
                AutoZero_On_Border.BorderBrush = Deselected;
            }
            if (Select == 1)
            {
                AutoZero_Off_Border.BorderBrush = Selected;
                AutoZero_Select = "AZERO 0";
            }
            else
            {
                AutoZero_Off_Border.BorderBrush = Deselected;
            }
            return AutoZero_Select;
        }
        //----------------------------------Auto Zero------------------------------------------------------

        //----------------------------------Trigger------------------------------------------------------
        private void Trigger_Internal_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                SerialWriteQueue.Add(Trigger_Indicator(0));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Trigger set to Internal.", 5);
            }
            else 
            {
                insert_Log("Cannot set Trigger while N Digit is set to 7½", 2);
            }
        }

        private void Trigger_External_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                SerialWriteQueue.Add(Trigger_Indicator(1));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Trigger set to External.", 5);
            }
            else
            {
                insert_Log("Cannot set Trigger while N Digit is set to 7½", 2);
            }
        }
        private void Trigger_Hold_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                SerialWriteQueue.Add(Trigger_Indicator(2));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Trigger set to Hold.", 5);
            }
            else
            {
                insert_Log("Cannot set Trigger while N Digit is set to 7½", 2);
            }
        }

        private void Trigger_Single_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isNDigit_7_Selected == false)
            {
                SerialWriteQueue.Add(Trigger_Indicator(3));
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Trigger set to Single.", 5);
            }
            else
            {
                insert_Log("Cannot set Trigger while N Digit is set to 7½", 2);
            }
        }

        private string Trigger_Indicator(int Select)
        {
            string Trigger_Select = "TRIG 1";
            if (Select == 0)
            {
                Trigger_Internal_Border.BorderBrush = Selected;
                Trigger_Select = "TRIG 1";
            }
            else
            {
                Trigger_Internal_Border.BorderBrush = Deselected;
            }
            if (Select == 1)
            {
                Trigger_External_Border.BorderBrush = Selected;
                Trigger_Select = "TRIG 2";
            }
            else
            {
                Trigger_External_Border.BorderBrush = Deselected;
            }
            if (Select == 2)
            {
                Trigger_Hold_Border.BorderBrush = Selected;
                Trigger_Select = "TRIG 4";
            }
            else
            {
                Trigger_Hold_Border.BorderBrush = Deselected;
            }
            if (Select == 3)
            {
                Trigger_Single_Border.BorderBrush = Selected;
                Trigger_Select = "TRIG SGL";
            }
            else
            {
                Trigger_Single_Border.BorderBrush = Deselected;
            }
            return Trigger_Select;
        }
        //----------------------------------Trigger------------------------------------------------------

        //----------------------------------Terminal------------------------------------------------------
        private void Terminal_Front_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Terminal_Indicator(0));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Terminal set to Front.", 5);
        }

        private void Terminal_Rear_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Terminal_Indicator(1));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Terminal set to Rear.", 5);
        }

        private void Terminal_Open_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Terminal_Indicator(2));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Terminal set to Open.", 5);
        }

        private void Terminal_Query_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("TERM?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Terminal Query Send.", 5);
        }

        private string Terminal_Indicator(int Select)
        {
            string Terminal_Select = "TERM 1";
            if (Select == 0)
            {
                Terminal_Front_Border.BorderBrush = Selected;
                Terminal_Select = "TERM 1";
            }
            else
            {
                Terminal_Front_Border.BorderBrush = Deselected;
            }
            if (Select == 1)
            {
                Terminal_Rear_Border.BorderBrush = Selected;
                Terminal_Select = "TERM 2";
            }
            else
            {
                Terminal_Rear_Border.BorderBrush = Deselected;
            }
            if (Select == 2)
            {
                Terminal_Open_Border.BorderBrush = Selected;
                Terminal_Select = "TERM OPEN";
            }
            else
            {
                Terminal_Open_Border.BorderBrush = Deselected;
            }
            return Terminal_Select;
        }
        //----------------------------------Terminal------------------------------------------------------

        //----------------------------------Display------------------------------------------------------
        //--------------------------------Display Options---------------------------------------

        private void Display_On_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Display_Selector(0));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Front Panel Display is On.", 3);
        }

        private void Display_Off_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add(Display_Selector(1));
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Front Panel Display is Off.", 3);
        }

        private void Display_Text_Button_Click(object sender, RoutedEventArgs e)
        {
            Regex Characters_Allowed = new Regex("^[a-zA-Z0-9]*$");
            string Message = Display_Text_Input.Text;
            if (Characters_Allowed.IsMatch(Message))
            {
                Message = Message.Replace(" ", "_");
                SerialWriteQueue.Add("DISP 2," + Message);
                lockControls();
                isUserSendCommand = true;
                Speedup_Interval();
                insert_Log("Front Panel Display Text is set.", 3);
            }
            else 
            {
                insert_Log("Front Panel Display Text Not Set. Message must be alphanumeric.", 2);
            }
        }

        private void Display_Clear_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("DISP 1");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Front Panel Display Text is cleared.", 3);
        }

        private void Lock_On_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("LOCK ON");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Front Panel Keyboard is Disabled.", 3);
            Lock_On_Border.BorderBrush = Selected;
            Lock_Off_Border.BorderBrush = Deselected;
        }

        private void Lock_Off_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("LOCK OFF");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Front Panel Keyboard is Enabled.", 3);
            Lock_On_Border.BorderBrush = Deselected;
            Lock_Off_Border.BorderBrush = Selected;
        }


        private string Display_Selector(int Display)
        {
            string Command = "DISP 1";
            if (Display == 0)
            {
                Display_On_Border.BorderBrush = Selected;
                Command = "DISP 1";
            }
            else
            {
                Display_On_Border.BorderBrush = Deselected;
            }
            if (Display == 1)
            {
                Display_Off_Border.BorderBrush = Selected;
                Command = "DISP 0";
            }
            else
            {
                Display_Off_Border.BorderBrush = Deselected;
            }
            return Command;
        }

        //-----------------------------------------------------------------------

        //----------------------------------Display------------------------------------------------------

        //-----------------------------------Queries-----------------------------------------------------

        private void ID_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("ID?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("ID? Instrument will send its ID back.", 5);
        }

        private void STB_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("STB?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("STB? Instrument will send its Status Byte back.", 5);
        }

        private void ERR_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("ERR?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("ERR? Instrument will send its Error Register Byte back.", 5);
        }

        private void AUXERR_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("AUXERR?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("AUXERR? Instrument will send its Aux Error Byte back.", 5);
        }

        private void CALNUM_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("CALNUM?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("CALNUM? How many times its been calibrated.", 5);
        }

        private void REV_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("REV?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("REV? Shows revision info.", 5);
        }

        //-----------------------------------Queries End-----------------------------------------------------

        //----------------------------------FIXEDZ----------------------------------------------------------

        private void FIXEDZ_ON_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("FIXEDZ ON");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("10MΩ Input Impedance for DCV 0.03V, 0.3V, 3V, 30V, 300V ranges", 5);
        }

        private void FIXEDZ_OFF_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("FIXEDZ OFF");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("10GΩ Input Impedance for DCV 0.03V, 0.3V, and 3V ranges, 10MΩ for 30V, 300V ranges.", 5);
        }

        private void FIXEDZ_Query_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("FIXEDZ?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Shows Input Impedance Info for DCV.", 5);
        }

        //----------------------------------FIXEDZ----------------------------------------------------------

        private void LFREQ50_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("LFREQ 50");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Line Frequency set to 50Hz.", 5);
        }

        private void LFREQ60_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("LFREQ 60");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Line Frequency set to 60Hz.", 5);
        }

        private void LFREQ_Query_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("LFREQ?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Shows what Line Frequency is set.", 5);
        }

        private void LINE_Query_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("LINE?");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("Shows Measured Line Frequency.", 5);
        }

        //--------------------------Measurements Options---------------------------------------------

        private void ACAL_AC_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("ACAL AC");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("ALCAL for AC Measurement will begin.", 5);
        }

        private void ACAL_OHMS_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("ACAL OHMS");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("ALCAL for Ohms Measurement will begin.", 5);
        }

        private void ACBAND_SLOW_Query_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("ACBAND 200");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("ACBAND set to Slow, affects ACV, ACI, FREQ, PER.", 5);
        }

        private void ACBAND_FAST_Query_Button_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("ACBAND 5000");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
            insert_Log("ACBAND set to Fast, affects ACV, ACI, FREQ, PER.", 5);
        }


        //----------------------------Speech Setup-------------------------------------------

        private void Speech_Continuous_Set_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                (bool isNum, double value) = isNumber(Speech_Continuous_Value.Text);
                if (isNum == true)
                {
                    if (value > 0)
                    {
                        Interlocked.Exchange(ref isSpeechContinuous, 1);
                        Speech_Measurement_Interval.Interval = (value * 60000);
                        Continuous_Selector(0);
                        insert_Log("Continuously voice measurement every " + value + " minutes.", 0);
                        Speech_Measurement_Interval.Start();
                    }
                    else
                    {
                        insert_Log("Continuous voice value must be a positive number.", 1);
                    }
                }
                else
                {
                    insert_Log("Continuous voice value must be a positive number.", 1);
                }
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void Speech_Continuous_Clear_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                Interlocked.Exchange(ref isSpeechContinuous, 0);
                Speech_Continuous_Value.Text = string.Empty;
                Continuous_Selector(1);
                insert_Log("Continuous voice measurement is cleared.", 0);
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }

        }

        private void Continuous_Selector(int status)
        {
            if (status == 0)
            {
                Speech_Continuous_Set_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_Continuous_Set_Border.BorderBrush = Deselected;
            }
            if (status == 1)
            {
                Speech_Continuous_Clear_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_Continuous_Clear_Border.BorderBrush = Deselected;
            }
        }

        private void Speech_MIN_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                (bool isNum, double value) = isNumber(Speech_MIN_Value.Text);
                if (isNum == true)
                {
                    Interlocked.Exchange(ref Speech_min_value, value);
                    Interlocked.Exchange(ref isSpeechMIN, 1);
                    MIN_Selector(0);
                    insert_Log("Voice measurement less than " + value, 0);
                    Speech_MIN_Max.Start();
                }
                else
                {
                    insert_Log("MIN voice value must be a number.", 1);
                }
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void Speech_MIN_Clear_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                Interlocked.Exchange(ref isSpeechMIN, 0);
                Interlocked.Exchange(ref Speech_min_value, 0);
                Speech_MIN_Value.Text = string.Empty;
                MIN_Selector(1);
                insert_Log("MIN voice measurement is cleared.", 0);
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void MIN_Selector(int status)
        {
            if (status == 0)
            {
                Speech_MIN_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_MIN_Border.BorderBrush = Deselected;
            }
            if (status == 1)
            {
                Speech_MIN_Clear_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_MIN_Clear_Border.BorderBrush = Deselected;
            }
        }

        private void Speech_MAX_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                (bool isNum, double value) = isNumber(Speech_MAX_Value.Text);
                if (isNum == true)
                {
                    Interlocked.Exchange(ref Speech_max_value, value);
                    Interlocked.Exchange(ref isSpeechMAX, 1);
                    MAX_Selector(0);
                    insert_Log("Voice measurement greater than " + value, 0);
                    Speech_MIN_Max.Start();
                }
                else
                {
                    insert_Log("MAX voice value must be a number.", 1);
                }
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void Speech_MAX_Clear_Button_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeechActive == 1)
            {
                Interlocked.Exchange(ref isSpeechMAX, 0);
                Interlocked.Exchange(ref Speech_max_value, 0);
                Speech_MAX_Value.Text = string.Empty;
                MAX_Selector(1);
                insert_Log("MAX voice measurement is cleared.", 0);
            }
            else
            {
                insert_Log("The Speech Synthesizer is not enabled. Enable it from Speech Menu.", 2);
            }
        }

        private void MAX_Selector(int status)
        {
            if (status == 0)
            {
                Speech_MAX_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_MAX_Border.BorderBrush = Deselected;
            }
            if (status == 1)
            {
                Speech_MAX_Clear_Border.BorderBrush = Selected;
            }
            else
            {
                Speech_MAX_Clear_Border.BorderBrush = Deselected;
            }
        }

        //-----------------------------------------------------------------------

        //-----------------------------Update Speed Options------------------------------------------

        private void UpdateSpeed_Value_Set_Button_Click(object sender, RoutedEventArgs e)
        {
            (bool isNum, double value) = isNumber(UpdateSpeed_Value.Text);
            if (isNum == true)
            {
                if (value > 0)
                {
                    insert_Log("You may need to wait for " + (UpdateSpeed / 1000) + " seconds before your new update speed takes effect.", 2);
                    insert_Log("Update Speed set to " + value + " seconds Command Send.", 5);
                    value = value * 1000;
                    UpdateSpeed = value;
                    UpdateSpeed_Selector(0);
                    isUpdateSpeed_Changed = true;
                }
                else
                {
                    insert_Log("Update Speed must be number greater than 0. Minimum value can be 0.01 seconds.", 1);
                }
            }
            else
            {
                insert_Log("Update Speed must be number.", 1);
            }
        }

        private void UpdateSpeed_Default_Set_Button_Click(object sender, RoutedEventArgs e)
        {
            insert_Log("You may to wait for " + (UpdateSpeed / 1000) + " seconds before your new update speed takes effect.", 2);
            UpdateSpeed = 1000;
            insert_Log("Update Speed set to " + (UpdateSpeed / 1000) + " seconds Command Send.", 5);
            UpdateSpeed_Selector(1);
            isUpdateSpeed_Changed = true;
        }

        private void UpdateSpeed_Fast_Set_Button_Click(object sender, RoutedEventArgs e)
        {
            insert_Log("You may to wait for " + (UpdateSpeed / 1000) + " seconds before your new update speed takes effect.", 2);
            UpdateSpeed = 10;
            insert_Log("Update Speed set to " + (UpdateSpeed / 1000) + " seconds Command Send.", 5);
            UpdateSpeed_Selector(2);
            isUpdateSpeed_Changed = true;
        }

        private void UpdateSpeed_Selector(int status)
        {
            if (status == 0)
            {
                UpdateSpeed_Value_Set_Border.BorderBrush = Selected;
            }
            else
            {
                UpdateSpeed_Value_Set_Border.BorderBrush = Deselected;
            }
            if (status == 1)
            {
                UpdateSpeed_Default_Set_Border.BorderBrush = Selected;
            }
            else
            {
                UpdateSpeed_Default_Set_Border.BorderBrush = Deselected;
            }
            if (status == 2)
            {
                UpdateSpeed_Fast_Set_Border.BorderBrush = Selected;
            }
            else
            {
                UpdateSpeed_Fast_Set_Border.BorderBrush = Deselected;
            }
        }

        //-----------------------------------------------------------------------

        private void Measurement_Green_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(0);
            Measurement_Color("#FF00FF17");

        }

        private void Measurement_Blue_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(1);
            Measurement_Color("#FF00C0FF");
        }

        private void Measurement_Red_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(2);
            Measurement_Color("Red");
        }

        private void Measurement_Yellow_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(3);
            Measurement_Color("#FFFFFF00");
        }

        private void Measurement_Orange_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(4);
            Measurement_Color("DarkOrange");
        }

        private void Measurement_Pink_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(5);
            Measurement_Color("DeepPink");
        }

        private void Measurement_White_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(6);
            Measurement_Color("White");
        }

        private void Measurement_Black_Click(object sender, RoutedEventArgs e)
        {
            Measurement_Color_Checker(7);
            Measurement_Color("Black");
        }

        private void Measurement_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            Measurement_Value.Foreground = Color;
            Measurement_Scale.Foreground = Color;
            Measurement_Type.Foreground = Color;
        }

        private void Measurement_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                Measurement_Green.IsChecked = true;
            }
            else
            {
                Measurement_Green.IsChecked = false;
            }
            if (Check == 1)
            {
                Measurement_Blue.IsChecked = true;
            }
            else
            {
                Measurement_Blue.IsChecked = false;
            }
            if (Check == 2)
            {
                Measurement_Red.IsChecked = true;
            }
            else
            {
                Measurement_Red.IsChecked = false;
            }
            if (Check == 3)
            {
                Measurement_Yellow.IsChecked = true;
            }
            else
            {
                Measurement_Yellow.IsChecked = false;
            }
            if (Check == 4)
            {
                Measurement_Orange.IsChecked = true;
            }
            else
            {
                Measurement_Orange.IsChecked = false;
            }
            if (Check == 5)
            {
                Measurement_Pink.IsChecked = true;
            }
            else
            {
                Measurement_Pink.IsChecked = false;
            }
            if (Check == 6)
            {
                Measurement_White.IsChecked = true;
            }
            else
            {
                Measurement_White.IsChecked = false;
            }
            if (Check == 7)
            {
                Measurement_Black.IsChecked = true;
            }
            else
            {
                Measurement_Black.IsChecked = false;
            }
        }

        private void MIN_Green_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(0);
            MIN_Color("#FF00FF17");
        }

        private void MIN_Blue_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(1);
            MIN_Color("#FF00C0FF");
        }

        private void MIN_Red_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(2);
            MIN_Color("Red");
        }

        private void MIN_Yellow_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(3);
            MIN_Color("#FFFFFF00");
        }

        private void MIN_Orange_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(4);
            MIN_Color("DarkOrange");
        }

        private void MIN_Pink_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(5);
            MIN_Color("DeepPink");
        }

        private void MIN_White_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(6);
            MIN_Color("White");
        }

        private void MIN_Black_Click(object sender, RoutedEventArgs e)
        {
            MIN_Color_Checker(7);
            MIN_Color("Black");
        }

        private void MIN_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            MIN_Value.Foreground = Color;
            MIN_Scale.Foreground = Color;
            MIN_Type.Foreground = Color;
            MIN_Label.Foreground = Color;
        }

        private void MIN_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                MIN_Green.IsChecked = true;
            }
            else
            {
                MIN_Green.IsChecked = false;
            }
            if (Check == 1)
            {
                MIN_Blue.IsChecked = true;
            }
            else
            {
                MIN_Blue.IsChecked = false;
            }
            if (Check == 2)
            {
                MIN_Red.IsChecked = true;
            }
            else
            {
                MIN_Red.IsChecked = false;
            }
            if (Check == 3)
            {
                MIN_Yellow.IsChecked = true;
            }
            else
            {
                MIN_Yellow.IsChecked = false;
            }
            if (Check == 4)
            {
                MIN_Orange.IsChecked = true;
            }
            else
            {
                MIN_Orange.IsChecked = false;
            }
            if (Check == 5)
            {
                MIN_Pink.IsChecked = true;
            }
            else
            {
                MIN_Pink.IsChecked = false;
            }
            if (Check == 6)
            {
                MIN_White.IsChecked = true;
            }
            else
            {
                MIN_White.IsChecked = false;
            }
            if (Check == 7)
            {
                MIN_Black.IsChecked = true;
            }
            else
            {
                MIN_Black.IsChecked = false;
            }
        }

        private void MAX_Green_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(0);
            MAX_Color("#FF00FF17");
        }

        private void MAX_Blue_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(1);
            MAX_Color("#FF00C0FF");
        }

        private void MAX_Red_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(2);
            MAX_Color("Red");
        }

        private void MAX_Yellow_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(3);
            MAX_Color("#FFFFFF00");
        }

        private void MAX_Orange_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(4);
            MAX_Color("DarkOrange");
        }

        private void MAX_Pink_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(5);
            MAX_Color("DeepPink");
        }

        private void MAX_White_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(6);
            MAX_Color("White");
        }

        private void MAX_Black_Click(object sender, RoutedEventArgs e)
        {
            MAX_Color_Checker(7);
            MAX_Color("Black");
        }

        private void MAX_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            MAX_Value.Foreground = Color;
            MAX_Scale.Foreground = Color;
            MAX_Type.Foreground = Color;
            MAX_Label.Foreground = Color;
        }

        private void MAX_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                MAX_Green.IsChecked = true;
            }
            else
            {
                MAX_Green.IsChecked = false;
            }
            if (Check == 1)
            {
                MAX_Blue.IsChecked = true;
            }
            else
            {
                MAX_Blue.IsChecked = false;
            }
            if (Check == 2)
            {
                MAX_Red.IsChecked = true;
            }
            else
            {
                MAX_Red.IsChecked = false;
            }
            if (Check == 3)
            {
                MAX_Yellow.IsChecked = true;
            }
            else
            {
                MAX_Yellow.IsChecked = false;
            }
            if (Check == 4)
            {
                MAX_Orange.IsChecked = true;
            }
            else
            {
                MAX_Orange.IsChecked = false;
            }
            if (Check == 5)
            {
                MAX_Pink.IsChecked = true;
            }
            else
            {
                MAX_Pink.IsChecked = false;
            }
            if (Check == 6)
            {
                MAX_White.IsChecked = true;
            }
            else
            {
                MAX_White.IsChecked = false;
            }
            if (Check == 7)
            {
                MAX_Black.IsChecked = true;
            }
            else
            {
                MAX_Black.IsChecked = false;
            }
        }

        private void AVG_Green_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(0);
            AVG_Color("#FF00FF17");
        }

        private void AVG_Blue_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(1);
            AVG_Color("#FF00C0FF");
        }

        private void AVG_Red_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(2);
            AVG_Color("Red");
        }

        private void AVG_Yellow_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(3);
            AVG_Color("#FFFFFF00");
        }

        private void AVG_Orange_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(4);
            AVG_Color("DarkOrange");
        }

        private void AVG_Pink_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(5);
            AVG_Color("DeepPink");
        }

        private void AVG_White_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(6);
            AVG_Color("White");
        }

        private void AVG_Black_Click(object sender, RoutedEventArgs e)
        {
            AVG_Color_Checker(7);
            AVG_Color("Black");
        }

        private void AVG_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            AVG_Value.Foreground = Color;
            AVG_Scale.Foreground = Color;
            AVG_Type.Foreground = Color;
            AVG_Label.Foreground = Color;
        }

        private void AVG_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                AVG_Green.IsChecked = true;
            }
            else
            {
                AVG_Green.IsChecked = false;
            }
            if (Check == 1)
            {
                AVG_Blue.IsChecked = true;
            }
            else
            {
                AVG_Blue.IsChecked = false;
            }
            if (Check == 2)
            {
                AVG_Red.IsChecked = true;
            }
            else
            {
                AVG_Red.IsChecked = false;
            }
            if (Check == 3)
            {
                AVG_Yellow.IsChecked = true;
            }
            else
            {
                AVG_Yellow.IsChecked = false;
            }
            if (Check == 4)
            {
                AVG_Orange.IsChecked = true;
            }
            else
            {
                AVG_Orange.IsChecked = false;
            }
            if (Check == 5)
            {
                AVG_Pink.IsChecked = true;
            }
            else
            {
                AVG_Pink.IsChecked = false;
            }
            if (Check == 6)
            {
                AVG_White.IsChecked = true;
            }
            else
            {
                AVG_White.IsChecked = false;
            }
            if (Check == 7)
            {
                AVG_Black.IsChecked = true;
            }
            else
            {
                AVG_Black.IsChecked = false;
            }
        }

        private void Background_White_Click(object sender, RoutedEventArgs e)
        {
            Background_Color_Checker(0);
            Background_Color("White");
        }

        private void Background_Black_Click(object sender, RoutedEventArgs e)
        {
            Background_Color_Checker(1);
            Background_Color("Black");
        }

        private void Background_Color(string HexValue)
        {
            SolidColorBrush Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString(HexValue));
            DisplayPanel_Background.Background = Color;
        }

        private void Background_Color_Checker(int Check)
        {
            if (Check == 0)
            {
                Background_White.IsChecked = true;
            }
            else
            {
                Background_White.IsChecked = false;
            }
            if (Check == 1)
            {
                Background_Black.IsChecked = true;
            }
            else
            {
                Background_Black.IsChecked = false;
            }
        }

        private void Randomize_Display_Colors(object sender, RoutedEventArgs e)
        {
            Random RGB_Value = new Random();
            int Value_Red = RGB_Value.Next(0, 255);
            int Value_Green = RGB_Value.Next(0, 255);
            int Value_Blue = RGB_Value.Next(0, 255);

            Set_Measurement_Color(Value_Red, Value_Green, Value_Blue);
            Set_MIN_Color(Value_Red, Value_Green, Value_Blue);
            Set_MAX_Color(Value_Red, Value_Green, Value_Blue);
            Set_AVG_Color(Value_Red, Value_Green, Value_Blue);

            insert_Log(Value_Red + "," + Value_Green + "," + Value_Blue + "," + "Measurement_Colors_Selected_RGB", 4);
        }

        private void FSI_Display_Click(object sender, RoutedEventArgs e)
        {
            B_FSI_Double = true;
            B_RTZ_Display = false;
            PSI_Display.IsChecked = false;
            FSI_Display.IsChecked = true;
            RTZ_Display.IsChecked = false;
        }

        private void PSI_Display_Click(object sender, RoutedEventArgs e)
        {
            B_FSI_Double = false;
            B_RTZ_Display = false;
            PSI_Display.IsChecked = true;
            FSI_Display.IsChecked = false;
            RTZ_Display.IsChecked = false;
        }

        private void RTZ_Display_Click(object sender, RoutedEventArgs e)
        {
            B_FSI_Double = false;
            B_RTZ_Display = true;
            PSI_Display.IsChecked = false;
            FSI_Display.IsChecked = false;
            RTZ_Display.IsChecked = true;
        }

        private void Calculate_AVG_Click(object sender, RoutedEventArgs e)
        {
            if (Calculate_AVG.IsChecked == true)
            {
                Interlocked.Exchange(ref AVG_Calculate, 1);
                insert_Log("Average will be calculated.", 0);
            }
            else
            {
                Interlocked.Exchange(ref AVG_Calculate, 0);
                insert_Log("Average will not be calculated.", 2);
            }
        }

        private void AVG_Res_2_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 2);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_3_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 3);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_4_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 4);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_5_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 5);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_6_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_resolution, 6);
            insert_Log("Average's resolution set to " + avg_resolution + ".", 0);
            AVG_Res_Selected();
        }

        private void AVG_Res_Selected()
        {
            if (avg_resolution == 2)
            {
                AVG_Res_2.IsChecked = true;
            }
            else
            {
                AVG_Res_2.IsChecked = false;
            }
            if (avg_resolution == 3)
            {
                AVG_Res_3.IsChecked = true;
            }
            else
            {
                AVG_Res_3.IsChecked = false;
            }
            if (avg_resolution == 4)
            {
                AVG_Res_4.IsChecked = true;
            }
            else
            {
                AVG_Res_4.IsChecked = false;
            }
            if (avg_resolution == 5)
            {
                AVG_Res_5.IsChecked = true;
            }
            else
            {
                AVG_Res_5.IsChecked = false;
            }
            if (avg_resolution == 6)
            {
                AVG_Res_6.IsChecked = true;
            }
            else
            {
                AVG_Res_6.IsChecked = false;
            }
        }

        private void Factor_50_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 50);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_100_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 100);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_200_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 200);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_400_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 400);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_800_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 800);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void Factor_1000_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref avg_factor, 1000);
            insert_Log("Average's factor set to " + avg_factor + ".", 0);
            AVG_Fac_Selected();
        }

        private void AVG_Fac_Selected()
        {
            if (avg_factor == 50)
            {
                Factor_50.IsChecked = true;
            }
            else
            {
                Factor_50.IsChecked = false;
            }
            if (avg_factor == 100)
            {
                Factor_100.IsChecked = true;
            }
            else
            {
                Factor_100.IsChecked = false;
            }
            if (avg_factor == 200)
            {
                Factor_200.IsChecked = true;
            }
            else
            {
                Factor_200.IsChecked = false;
            }
            if (avg_factor == 400)
            {
                Factor_400.IsChecked = true;
            }
            else
            {
                Factor_400.IsChecked = false;
            }
            if (avg_factor == 800)
            {
                Factor_800.IsChecked = true;
            }
            else
            {
                Factor_800.IsChecked = false;
            }
            if (avg_factor == 1000)
            {
                Factor_1000.IsChecked = true;
            }
            else
            {
                Factor_1000.IsChecked = false;
            }
        }

        private void Show_7Digit_Info_Click(object sender, RoutedEventArgs e)
        {
            try 
            {
                if (Show_7Digit_Info.IsChecked == true)
                {
                    show_7Digit_Info = true;
                }
                else 
                {
                    show_7Digit_Info = false;
                }
            } 
            catch (Exception) 
            {
                
            }
        }

        private void Voice_Precision_0_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 0);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_1_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 1);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_2_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 2);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_3_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 3);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_4_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 4);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_5_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 5);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_6_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 6);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_7_Click(object sender, RoutedEventArgs e)
        {
            Interlocked.Exchange(ref Speech_Value_Precision, 7);
            insert_Log("Speech's precision is set to " + Speech_Value_Precision + ".", 0);
            Voice_Precision_Selected();
        }

        private void Voice_Precision_Selected()
        {
            if (Speech_Value_Precision == 0)
            {
                Voice_Precision_0.IsChecked = true;
            }
            else
            {
                Voice_Precision_0.IsChecked = false;
            }
            if (Speech_Value_Precision == 1)
            {
                Voice_Precision_1.IsChecked = true;
            }
            else
            {
                Voice_Precision_1.IsChecked = false;
            }
            if (Speech_Value_Precision == 2)
            {
                Voice_Precision_2.IsChecked = true;
            }
            else
            {
                Voice_Precision_2.IsChecked = false;
            }
            if (Speech_Value_Precision == 3)
            {
                Voice_Precision_3.IsChecked = true;
            }
            else
            {
                Voice_Precision_3.IsChecked = false;
            }
            if (Speech_Value_Precision == 4)
            {
                Voice_Precision_4.IsChecked = true;
            }
            else
            {
                Voice_Precision_4.IsChecked = false;
            }
            if (Speech_Value_Precision == 5)
            {
                Voice_Precision_5.IsChecked = true;
            }
            else
            {
                Voice_Precision_5.IsChecked = false;
            }
            if (Speech_Value_Precision == 6)
            {
                Voice_Precision_6.IsChecked = true;
            }
            else
            {
                Voice_Precision_6.IsChecked = false;
            }
            if (Speech_Value_Precision == 7)
            {
                Voice_Precision_7.IsChecked = true;
            }
            else
            {
                Voice_Precision_7.IsChecked = false;
            }
        }

        private void Main_Window_Closed(object sender, EventArgs e)
        {
            try
            {
                if (Serial_COM_Info.isConnected == true)
                {
                    HP3457A.Close();
                    HP3457A.Dispose();
                }
            }
            catch (Exception) 
            {
                
            }
        }

        private void Local_Exit_Click(object sender, RoutedEventArgs e)
        {
            SerialWriteQueue.Add("LOCAL_EXIT");
            lockControls();
            isUserSendCommand = true;
            Speedup_Interval();
        }

        private void Load_Main_Window_Settings()
        {
            try
            {
                List<String> Config_Lines = new List<string>();
                string Software_Location = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\" + "Settings.txt";
                string[] Config_Parts;
                using (var readFile = new StreamReader(Software_Location))
                {
                    Config_Parts = readFile.ReadLine().Split(',');
                    Set_Measurement_Color(int.Parse(Config_Parts[0]), int.Parse(Config_Parts[1]), int.Parse(Config_Parts[2]));
                    Config_Parts = readFile.ReadLine().Split(',');
                    Set_MIN_Color(int.Parse(Config_Parts[0]), int.Parse(Config_Parts[1]), int.Parse(Config_Parts[2]));
                    Config_Parts = readFile.ReadLine().Split(',');
                    Set_MAX_Color(int.Parse(Config_Parts[0]), int.Parse(Config_Parts[1]), int.Parse(Config_Parts[2]));
                    Config_Parts = readFile.ReadLine().Split(',');
                    Set_AVG_Color(int.Parse(Config_Parts[0]), int.Parse(Config_Parts[1]), int.Parse(Config_Parts[2]));
                    insert_Log("Settings.txt file loaded.", 0);
                }

            }
            catch (Exception Ex)
            {
                insert_Log(Ex.Message, 2);
                insert_Log("Could not load Settings.txt file, try again.", 2);
            }
        }

        private void Set_Measurement_Color(int Red, int Green, int Blue)
        {
            Measurement_Color_Checker(9);
            if ((Red <= 255 & Red >= 0) & (Green <= 255 & Green >= 0) & (Blue <= 255 & Blue >= 0))
            {
                Measurement_Value.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                Measurement_Scale.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                Measurement_Type.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
            }
            else
            {
                insert_Log("Measurement_Value_Color: RGB Values must be between 0 to 255, try again.", 2);
            }
        }

        private void Set_MIN_Color(int Red, int Green, int Blue)
        {
            MIN_Color_Checker(9);
            if ((Red <= 255 & Red >= 0) & (Green <= 255 & Green >= 0) & (Blue <= 255 & Blue >= 0))
            {
                MIN_Value.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MIN_Scale.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MIN_Type.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MIN_Label.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
            }
            else
            {
                insert_Log("MIN_Value_Color: RGB Values must be between 0 to 255, try again.", 2);
            }
        }

        private void Set_MAX_Color(int Red, int Green, int Blue)
        {
            MAX_Color_Checker(9);
            if ((Red <= 255 & Red >= 0) & (Green <= 255 & Green >= 0) & (Blue <= 255 & Blue >= 0))
            {
                MAX_Value.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MAX_Scale.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MAX_Type.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                MAX_Label.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
            }
            else
            {
                insert_Log("MAX_Value_Color: RGB Values must be between 0 to 255, try again.", 2);
            }
        }

        private void Set_AVG_Color(int Red, int Green, int Blue)
        {
            AVG_Color_Checker(9);
            if ((Red <= 255 & Red >= 0) & (Green <= 255 & Green >= 0) & (Blue <= 255 & Blue >= 0))
            {
                AVG_Value.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                AVG_Scale.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                AVG_Type.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
                AVG_Label.Foreground = new SolidColorBrush(Color.FromArgb(255, (byte)(Red), (byte)(Green), (byte)(Blue)));
            }
            else
            {
                insert_Log("AVG_Value_Color: RGB Values must be between 0 to 255, try again.", 2);
            }
        }
    }
}
