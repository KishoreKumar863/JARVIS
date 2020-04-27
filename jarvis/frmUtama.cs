using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Speech.Recognition;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Speech.Synthesis;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VoiceRecognition
{
    public partial class frmUtama : MetroFramework.Forms.MetroForm
    {

        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();

        public frmUtama()
        {
            InitializeComponent();
        }
       

        string[] data_personal =  { "who are you","what is your name", "Jarvis"
                                    , "hello", "whats up", "your from"
                                    ,"what is your age","thank you"};
         
        string[] data_todo = {  "play a video", "google","do you know CORTANA"
                                    , "tell me about CORTANA", "lock", "command prompt"
                                    , "task manager","write my name","play a song"
                                    , "windows explorer","print all comands","localhost","clear","notepad" };

        void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            try
            {
                string t_ = "Processing ";


               
            
                if (e.Result.Text == data_personal[0])
                {
                    frmlist.addData(t_+"-> " + data_personal[0]);
                    synthesizer.SpeakAsync("I am a personal assistant");

                }
                else if (e.Result.Text == data_personal[1])
                {
                    frmlist.addData(t_ + "-> " + data_personal[1]);
                    synthesizer.SpeakAsync("My name is Jarvis");
                }
                else if (e.Result.Text == data_personal[2])
                {
                    frmlist.addData(t_ + "-> " + data_personal[2]);
                    synthesizer.SpeakAsync("Yes my boss, What can i do for you ?");
                }
                else if (e.Result.Text == data_personal[3])
                {
                    frmlist.addData(t_ + "-> " + data_personal[3]);
                    synthesizer.SpeakAsync("Yes my boss, i am here");
                }
                else if (e.Result.Text == data_personal[4])
                {
                    frmlist.addData(t_ + "-> " + data_personal[4]);
                    synthesizer.SpeakAsync("Nothing my boss, I am do nothing");
                }
                else if (e.Result.Text == data_personal[5])
                {
                    frmlist.addData(t_ + "-> " + data_personal[5]);
                    synthesizer.SpeakAsync("i come from another galaxy and i live on earth now");
                }
                else if (e.Result.Text == data_personal[6])
                {
                    frmlist.addData(t_ + "-> " + data_personal[6]);
                    synthesizer.SpeakAsync("i do not have age my boss");
                }
               
                else if (e.Result.Text == data_personal[7])
                {
                    frmlist.addData(t_ + "-> " + data_personal[7]);
                    synthesizer.SpeakAsync("you are welcome my boss");
                }
               

                // =====================================
                else if (e.Result.Text == data_todo[0])
                {
                    frmlist.addData(t_ + "-> " + data_todo[0]);
                    synthesizer.SpeakAsync("ok my boss, please wait, i'll play a video !");
                    Process.Start("C:/Users/Kishore Kumar V/Downloads/kis/videos/editz/amm.mp4");
                }
                else if (e.Result.Text == data_todo[1])
                {
                    frmlist.addData(t_ + "-> " + data_todo[1]);
                    synthesizer.SpeakAsync("ok my boss, please wait, i'm open google for you !");
                    Process.Start("https://google.com/");
                }
                else if (e.Result.Text == data_todo[2])
                {
                    frmlist.addData(t_ + "-> " + data_todo[2]);
                    synthesizer.SpeakAsync("Yes, i know cortana my boss");
                }
                else if (e.Result.Text == data_todo[3])
                {
                    frmlist.addData(t_ + "-> " + data_todo[3]);
                    synthesizer.SpeakAsync("Cortana is Artificial Intelligence that was created by windows, but cortana is not smart. Cortana and me is same. Do you understand me my bos ?");
                }

                else if (e.Result.Text == data_todo[4])
                {
                    frmlist.addData(t_ + "-> " + data_todo[4]);
                    synthesizer.SpeakAsync("Lock computer");
                    LockWorkStation();
                  
                }
                else if (e.Result.Text == data_todo[5])
                {
                    frmlist.addData(t_ + "-> " + data_todo[5]);
                    synthesizer.SpeakAsync("ok my boss, please wait, i'm open command prompt for you !");
                    Process.Start("cmd");
                }
                else if (e.Result.Text == data_todo[6])
                {
                    frmlist.addData(t_ + "-> " + data_todo[6]);
                    synthesizer.SpeakAsync("ok my boss, please wait, i'm open task manager for you !");
                    Process.Start("taskmgr");
                }
                else if (e.Result.Text == data_todo[7])
                {
                    rt_echo.Text = "";
                    rt_echo.Text = "Kishore Kumar V";
                    frmlist.addData(t_ + "-> " + data_todo[7]);
                    synthesizer.SpeakAsync("ok my boss, your name has been writed");
                   
                }
                else if (e.Result.Text == data_todo[8])
                {
                    frmlist.addData(t_ + "-> " + data_todo[8]);
                    synthesizer.SpeakAsync("ok my boss, please wait, i'll play a for you !");
                    Process.Start("C:/Users/Kishore Kumar V/Downloads/kis/songs/ENGLISH/EVERY_BODY.mp3");
                }
                else if (e.Result.Text == data_todo[9])
                {
                    frmlist.addData(t_ + "-> " + data_todo[9]);
                    synthesizer.SpeakAsync("ok my boss, please wait, i'm opening windows explorer for you !");
                    Process.Start("explorer");
                }
                else if (e.Result.Text == data_todo[10])
                {
                  
                    synthesizer.SpeakAsync("Oke my boss. I am print all comand !");
                    printAllComand();

                }
                else if (e.Result.Text == data_todo[11])
                {
                    frmlist.addData(t_ + "-> " + data_todo[11]);
                    synthesizer.SpeakAsync("Oke my boss. I will open localhost");
                    Process.Start("C:/xampp/htdocs");

                }
                else if (e.Result.Text == data_todo[12])
                {

                    synthesizer.SpeakAsync("Oke my boss. I will clear list");
                    frmlist.clearCommand();
                }
                else if (e.Result.Text == data_todo[13])
                {
                    frmlist.addData(t_ + "-> " + data_todo[13]);
                    synthesizer.SpeakAsync("Oke my boss. I will open notepad");
                    Process.Start("notepad");
                    
                }
                else
                {
                    frmlist.addData("Searching...");
                    Process.Start("www.google.com/search?site=&source=hp&q=" + e.Result.Text);
                }



            }
            catch (Exception ex)
            {
                frmlist.addData( "Error : "+ex.ToString());
            }


        }
        [DllImport("user32")]
        public static extern void LockWorkStation();
      

        private FrmList frmlist = new FrmList();
        private void frmUtama_Load(object sender, EventArgs e)
        {
            btn_disable_voice.Enabled = false;

            frmlist.Show();


            try
            {
                Choices commands = new Choices();
                commands.Add(new string[] {
                    data_personal[0]
                ,   data_personal[1]
                ,   data_personal[2]
                ,   data_personal[3]
                ,   data_personal[4]
                ,   data_personal[5]
                ,   data_personal[6]
                ,   data_personal[7]
                ,   data_todo[0]
                ,   data_todo[1]
                ,   data_todo[2]
                ,   data_todo[3]
                ,   data_todo[4]
                ,   data_todo[5]
                ,   data_todo[6]
                ,   data_todo[7]
                ,   data_todo[8]
                ,   data_todo[9]
                ,   data_todo[10]
                ,   data_todo[11]
                ,   data_todo[12]
                 ,   data_todo[13]
                });




                GrammarBuilder gBuilder = new GrammarBuilder();
                gBuilder.Append(commands);
                Grammar grammar = new Grammar(gBuilder);

                recEngine.LoadGrammarAsync(grammar);
                recEngine.SetInputToDefaultAudioDevice();
                recEngine.SpeechRecognized += recEngine_SpeechRecognized;

               
            }
            catch
            {

            }
        }

        private void btn_enable_voice_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsync(RecognizeMode.Multiple);
            btn_enable_voice.Enabled = false;
            btn_disable_voice.Enabled = true;
        }

        private void btn_disable_voice_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsyncStop();
            btn_enable_voice.Enabled = true;
            btn_disable_voice.Enabled = false;
        }

        private void btn_talk_Click(object sender, EventArgs e)
        {
            if (rt_echo.Text == string.Empty)
            {
                return;
            }
            synthesizer.SpeakAsync(rt_echo.Text);
        }


        private void printAllComand()
        {
            int lengtData = data_personal.Length;
            frmlist.clearCommand();
            for (int i = 0; i < lengtData; i++)
            {
                frmlist.addData("- " + data_personal[i]);
            }
            int lengthTodo = data_todo.Length;

            for (int j = 0; j < lengthTodo; j++)
            {
                frmlist.addData(" - " + data_todo[j]);
            }
        }

      

        private void btn_print_command_Click(object sender, EventArgs e)
        {
            printAllComand();
        }
    }
}
