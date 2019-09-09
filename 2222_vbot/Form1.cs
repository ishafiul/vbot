using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
namespace _2222_vbot
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        StreamReader CommandsReader = new StreamReader(@"C:\users\" + Environment.UserName.ToString() + @"\documents\commands.txt");
        SpeechSynthesizer sSynth = new SpeechSynthesizer();
        PromptBuilder pBuilder = new PromptBuilder();
        SpeechRecognitionEngine sRecognize = new SpeechRecognitionEngine();
        string name = "Microsoft Zira Desktop";
        Point lastPosition;

        private void Form1_Load(object sender, EventArgs e)
        {
            initializeSpeach();
        }
        GrammarBuilder gbuilder = new GrammarBuilder();
        public void initializeSpeach()
        {


            Choices sList = new Choices();

            //Add the words



            try
            {

                gbuilder.Append(new Choices(System.IO.File.ReadAllLines(@"C:\users\" + Environment.UserName.ToString() + @"\documents\commands.txt")));
            }
            catch { MessageBox.Show("The 'Commands' file must not contain empty lines.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); pr.StartInfo.FileName = @"C:\users\" + Environment.UserName.ToString() + @"\documents\commands.txt"; pr.Start(); Application.Exit(); return; }

            Grammar gr = new Grammar(gbuilder);
            try
            {
                sRecognize.UnloadAllGrammars();
                sRecognize.RecognizeAsyncCancel();
                sRecognize.RequestRecognizerUpdate();
                sRecognize.LoadGrammar(gr);
                sRecognize.SpeechRecognized += sRecognize_SpeechRecognized;

                sRecognize.SetInputToDefaultAudioDevice();
                sRecognize.RecognizeAsync(RecognizeMode.Multiple);




            }

            catch
            {
                MessageBox.Show("Grammar Builder Error");
                return;
            }

        }


        bool start = false;

        Process pr = new Process();
        public void speakText(string textSpeak)
        {
            sRecognize.RecognizeAsyncCancel();
            sRecognize.RecognizeAsyncStop();
            pBuilder.ClearContent();
            pBuilder.AppendText(textSpeak.ToString());
            sSynth.SelectVoice(name);
            sSynth.SpeakAsync(pBuilder);
            sRecognize.RecognizeAsyncCancel();
            sRecognize.RecognizeAsyncStop();
            sRecognize.RecognizeAsync(RecognizeMode.Multiple);
        }
        public void lockComputer()
        {

            System.Diagnostics.Process.Start(@"C:\WINDOWS\system32\rundll32.exe", "user32.dll,LockWorkStation");
            return;
        }
        internal void cancelSpeech()
        {
            sSynth.SpeakAsyncCancelAll();
        }
        bool exitCondition = false;


        private void sRecognize_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (exitCondition)
            {
                Thread.Sleep(100);
                if (e.Result.Text == "yes")
                {
                    sSynth.SpeakAsyncCancelAll();
                    sRecognize.RecognizeAsyncCancel();
                    Application.Exit();

                    return;
                }
                else { exitCondition = false; speakText("Exit Cancelled"); return; }
            }
            switch (e.Result.Text)
            {
                case "hello":
                    listBox2.Items.Add(e.Result.Text.ToString());
                    speakText("Hello " + Environment.UserName.ToString());
                    break;
                case "invisable":
                    if (ActiveForm.Visible == true)
                    {
                        listBox2.Items.Add(e.Result.Text.ToString());
                        ActiveForm.ShowInTaskbar = false; ActiveForm.Hide();


                        speakText("I am now invisible. You can access me by clicking on the icon down here in the tray.");
                        break;
                    }
                    else
                    {

                        break;
                    }
                case "hi":
                    listBox2.Items.Add(e.Result.Text.ToString());
                    speakText("Hello, " + Environment.UserName.ToString());
                    break;

                case "exit":
                    listBox2.Items.Add(e.Result.Text.ToString());
                    speakText("Are you sure you want to exit?");

                    exitCondition = true;
                    break;

                case "thank you":
                    listBox2.Items.Add(e.Result.Text.ToString());
                    speakText("You're welcome");
                    break;

                case "stop talking":
                    listBox2.Items.Add(e.Result.Text.ToString());
                    cancelSpeech();
                    break;
                case "lock computer":
                    listBox2.Items.Add(e.Result.Text.ToString());
                    lockComputer();
                    break;



                case "be quiet":
                    listBox2.Items.Add(e.Result.Text.ToString());
                    cancelSpeech();
                    sSynth.SpeakAsyncCancelAll();
                    break;

                case "internet":
                    listBox2.Items.Add(e.Result.Text.ToString());
                    speakText("One moment.");
                    Process pr = new Process();
                    pr.StartInfo.FileName = "http://www.google.com/";
                    pr.Start();
                    break;

                case "stop listening":
                    listBox2.Items.Add(e.Result.Text.ToString());
                    speakText("Ok.");
                    sRecognize.RecognizeAsyncCancel();
                    sRecognize.RecognizeAsyncStop();

                    break;

                case "hide":
                    break;

            }







        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
