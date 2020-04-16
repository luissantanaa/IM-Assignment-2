using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Xml.Linq;
using mmisharp;
using Newtonsoft.Json;
using System.Threading;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Speech.Recognition;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Timers;


namespace AppGui
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static MmiCommunication mmiC;

        private Microsoft.Office.Interop.PowerPoint.Application PPTAPP = new Microsoft.Office.Interop.PowerPoint.Application();
        private Presentation pptPresentation;
        private Slides slides;
        private Boolean WOKE = false;
        private Boolean OPENED = false;
        private int index = 1;
        private SlideShowView objSlideShowView;
        private String intpart;
        private double slideTimeLimit;
        private static System.Timers.Timer timerPerSlide = null;
        private static System.Timers.Timer timerFull = null;
        private static int min = 60000;
        private static LifeCycleEvents lce;
        private static string exNot;
        public MainWindow()
        {
           
          
            
        InitializeComponent();



            lce = new LifeCycleEvents("GUI", "TTS", "speech-1", "acoustic", "command");
            mmiC = new MmiCommunication("localhost",8000, "User1", "GUI");
            mmiC.Send(lce.NewContextRequest());
            mmiC.Message += MmiC_Message;

            mmiC.Start();

        }

        private void MmiC_Message(object sender, MmiEventArgs e)
        {
            Console.WriteLine(e.Message);
            var doc = XDocument.Parse(e.Message);
            try {
                //var wak = doc.Descendants("wake").FirstOrDefault().Value;
                var com = doc.Descendants("command").FirstOrDefault().Value;
                Console.WriteLine(":::::" + com);
                dynamic json = JsonConvert.DeserializeObject(com);
                Console.WriteLine(":::::" + json);
                Console.WriteLine((string)json.recognized[0].ToString());
                Console.WriteLine((string)json.recognized[1].ToString());


                var wake = ((string)json.recognized[0].ToString());
                var command = ((string)json.recognized[1].ToString());


                string[] split = command.Split(':');

                command = split[0];
                try
                {
                    intpart = split[1];
                    Console.WriteLine("sl " + intpart);
                }
                catch
                {
                    intpart = "";
                }
                Console.WriteLine(wake);
                Console.WriteLine(command);

                if (WOKE && OPENED)
                {
                    switch (command)
                    {
                        case "nAprest":

                            if (!(PPTAPP.SlideShowWindows.Count > 0))
                            {
                                exNot = lce.ExtensionNotification("", "", 0, "De momento não esta no modo de apresentação");
                                mmiC.Send(exNot);
                            }
                            else
                            {
                                objSlideShowView.Exit();
                                //restartSre();
                            }
                            break;
                        case "aprest":

                            if (slides.Count.Equals(0))
                            {
                                exNot = lce.ExtensionNotification("", "", 0, "A apresentação tem que ter pelo menos um slide");
                                mmiC.Send(exNot);
                                //tts.Speak("A apresentação tem que ter pelo menos um slide");
                            }
                            else
                            {
                                pptPresentation.SlideShowSettings.ShowPresenterView = MsoTriState.msoFalse;
                                pptPresentation.SlideShowSettings.Run();
                                objSlideShowView = pptPresentation.SlideShowWindow.View;
                                objSlideShowView.Application.SlideShowWindows[1].Activate();

                            }

                            break;

                        case "nots":
                            try
                            {
                                slides = pptPresentation.Slides;
                                if (intpart == "")
                                {
                                    var slide = slides[index];
                                    exNot = lce.ExtensionNotification("", "", 0, slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                                    mmiC.Send(exNot);
                                    //tts.Speak(slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                                }
                                else
                                {
                                    var slide = slides[Int32.Parse(intpart)];
                                    exNot = lce.ExtensionNotification("", "", 0, slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                                    mmiC.Send(exNot);
                                    // tts.Speak(slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                                }
                            }
                            catch
                            {
                                exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel ler notas do diapositivo " + intpart);
                                mmiC.Send(exNot);
                                //tts.Speak("Desculpe, não é possivel ler notas do diapositivo " + intpart);
                            }

                            break;

                        case "salt":

                            slides = pptPresentation.Slides;
                            try
                            {
                                slides[Int32.Parse(intpart)].Select();
                                //restartSre();
                            }
                            catch
                            {
                                exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel avançar para o diapositivo " + intpart);
                                mmiC.Send(exNot);
                                //tts.Speak("Desculpe, não é possivel avançar para o diapositivo " + intpart);
                            }

                            break;

                        case "limit":

                            try
                            {
                                slides = pptPresentation.Slides;
                                int size = slides.Count;
                                slideTimeLimit = Int32.Parse(intpart);
                                Console.WriteLine(slideTimeLimit);
                                Console.WriteLine(slideTimeLimit * min);

                                double timePerSlide = slideTimeLimit / size;
                                timerFull = new System.Timers.Timer(slideTimeLimit * min);
                                timerFull.Elapsed += OnTimedEventFull;
                                timerFull.Enabled = true;

                                timerPerSlide = new System.Timers.Timer(timePerSlide * min);
                                timerPerSlide.Elapsed += OnTimedEvent;
                                timerPerSlide.Enabled = true;
                                exNot = lce.ExtensionNotification("", "", 0, "Limite de " + intpart + " minutos definido");
                                mmiC.Send(exNot);
                                //tts.Speak("Limite de " + intpart + " minutos definido");

                            }
                            catch
                            {
                                exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel definir limite");
                                mmiC.Send(exNot);
                                //tts.Speak("Desculpe, não é possivel definir limite");
                            }

                            break;

                        case "avn":

                            slides = pptPresentation.Slides;

                            try
                            {
                                if (!(PPTAPP.SlideShowWindows.Count > 0))
                                {
                                    index++;
                                    slides[index].Select();
                                }
                                else
                                {
                                    index++;
                                    pptPresentation.SlideShowWindow.View.Next();
                                }

                                
                                if (timerPerSlide != null)
                                {
                                    if (timerPerSlide.Enabled)
                                    {
                                        timerPerSlide.Stop();
                                        timerPerSlide.Start();
                                    }
                                    else
                                    {
                                        timerPerSlide.Start();
                                    }
                                }
                                //restartSre();
                            
                            }
                            catch
                            {
                                index--;
                                exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel avançar para o diapositivo seguinte");
                                mmiC.Send(exNot);
                                //tts.Speak("Desculpe, não é possivel avançar para o diapositivo seguinte");
                            }

                            break;

                        case "rec":
                            slides = pptPresentation.Slides;
                            try
                            {
                                if (!(PPTAPP.SlideShowWindows.Count > 0))
                                {
                                    index--;
                                    slides[index].Select();
                                }
                                else
                                {
                                    index--;
                                    pptPresentation.SlideShowWindow.View.Previous();
                                }
                            }
                            catch
                            {
                                index++;
                                exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel recuar para o diapositivo anterior");
                                mmiC.Send(exNot);
                                //tts.Speak("Desculpe, não é possivel recuar para o diapositivo anterior");
                            }
                            break;


                        case "acab":

                            try
                            {
                                pptPresentation.Close();
                                //restartSre();
                            }
                            catch
                            {
                                exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possível terminar a apresentação");
                                mmiC.Send(exNot);

                                //tts.Speak("Desculpe, não é possível terminar a apresentação");
                            }
                            break;

                        case "adi":
                            try
                            {
                                slides = pptPresentation.Slides;
                                slides.Add(pptPresentation.Slides.Count + 1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                                //restartSre();
                            }
                            catch
                            {
                                exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possível adicionar slide");
                                mmiC.Send(exNot);
                                //tts.Speak("Desculpe, não é possível adicionar slide");
                            }

                            break;

                        case "rem":
                            try
                            {
                                slides = pptPresentation.Slides;
                                slides[index].Delete();
                                index--;
                                //restartSre();
                            }
                            catch
                            {
                                exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possível remover slide");
                                mmiC.Send(exNot);
                                //tts.Speak("Desculpe, não é possível remover slide");
                            }

                            break;

                        case "grdrppt":
                            try
                            {
                                pptPresentation.SaveAs("temp", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                                //restartSre();
                            }
                            catch
                            {
                                exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não foi possivel guardar a apresentação");
                                mmiC.Send(exNot);
                                //tts.Speak("Desculpe, não foi possivel guardar a apresentação");
                            }
                            break;




                    }
                }
                else if (!(OPENED) && WOKE)
                {
                    if (command.Equals("abr"))
                    {
                        OPENED = true;
                        try
                        {
                            PPTAPP.Visible = MsoTriState.msoTrue;
                            Presentations ppPresens = PPTAPP.Presentations;
                            pptPresentation = ppPresens.Open("temp", MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
                        }
                        catch (System.IO.FileNotFoundException)
                        {
                            pptPresentation = PPTAPP.Presentations.Add(MsoTriState.msoTrue);
                        }
                        //restartSre();
                    }
                    else
                    {
                        exNot = lce.ExtensionNotification("", "", 0, "Por favor, use o comando 'abrir' para iniciar a aplicação");
                        mmiC.Send(exNot);
                        //tts.Speak("Por favor, use o comando 'abrir' para iniciar a aplicação");
                    }

                }
                else if (!(WOKE) && !(OPENED))
                {
                    if (!wake.Equals("") && !command.Equals(""))
                    {
                        switch (command)
                        {

                            case "abr":
                                try
                                {
                                    OPENED = true;
                                    PPTAPP.Visible = MsoTriState.msoTrue;
                                    Presentations ppPresens = PPTAPP.Presentations;
                                    pptPresentation = ppPresens.Open("temp", MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
                                    //restartSre();
                                }
                                catch (System.IO.FileNotFoundException)
                                {
                                    pptPresentation = PPTAPP.Presentations.Add(MsoTriState.msoTrue);
                                    //restartSre();
                                }
                                break;

                            default:
                                exNot = lce.ExtensionNotification("", "", 0, "Por favor use o comando 'Powerpoint abrir' para iniciar a aplicação");
                                mmiC.Send(exNot);
                                //tts.Speak("Por favor use o comando 'Powerpoint abrir' para iniciar a aplicação");
                                break;
                        }
                    }

                    if (wake.Equals("ppt") && command.Equals(""))
                    {
                        try
                        {


                            OPENED = true;
                            WOKE = true;
                            PPTAPP.Visible = MsoTriState.msoTrue;
                            Presentations ppPresens = PPTAPP.Presentations;
                            pptPresentation = ppPresens.Open("temp", MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
                            exNot = lce.ExtensionNotification("", "", 0, "Sim?");
                            mmiC.Send(exNot);
                            //tts.Speak("Sim?");


                        }
                        catch (System.IO.FileNotFoundException)
                        {
                            pptPresentation = PPTAPP.Presentations.Add(MsoTriState.msoTrue);
                            //restartSre();
                        }

                    }
                }
                else
                {
                    if (!wake.Equals("") && !command.Equals(""))
                    {
                        switch (command)
                        {

                            case "nAprest":

                                if (!(PPTAPP.SlideShowWindows.Count > 0))
                                {
                                    exNot = lce.ExtensionNotification("", "", 0, "De momento não esta no modo de apresentação");
                                    mmiC.Send(exNot);
                                    //tts.Speak("De momento não esta no modo de apresentação");
                                }
                                else
                                {
                                    objSlideShowView.Exit();
                                    //restartSre();
                                }
                                break;
                            case "aprest":

                                if (slides.Count.Equals(0))
                                {
                                    exNot = lce.ExtensionNotification("", "", 0, "A apresentação tem que ter pelo menos um slide");
                                    mmiC.Send(exNot);
                                    // tts.Speak("A apresentação tem que ter pelo menos um slide");
                                }
                                else
                                {
                                    pptPresentation.SlideShowSettings.ShowPresenterView = MsoTriState.msoFalse;
                                    pptPresentation.SlideShowSettings.Run();
                                    objSlideShowView = pptPresentation.SlideShowWindow.View;
                                    objSlideShowView.Application.SlideShowWindows[1].Activate();
                                    //restartSre();
                                }

                                break;

                            case "nots":
                                try
                                {
                                    slides = pptPresentation.Slides;
                                    if (intpart == "")
                                    {
                                        var slide = slides[index];
                                        exNot = lce.ExtensionNotification("", "", 0, slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                                        mmiC.Send(exNot);
                                        //tts.Speak(slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                                    }
                                    else
                                    {
                                        var slide = slides[Int32.Parse(intpart)];
                                        exNot = lce.ExtensionNotification("", "", 0, slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                                        mmiC.Send(exNot);
                                        //tts.Speak(slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                                    }
                                }
                                catch
                                {
                                    exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel ler notas do diapositivo " + intpart);
                                    mmiC.Send(exNot);
                                    //tts.Speak("Desculpe, não é possivel ler notas do diapositivo " + intpart);
                                }

                                break;

                            case "salt":

                                slides = pptPresentation.Slides;
                                try
                                {
                                    slides[Int32.Parse(intpart)].Select();
                                    //restartSre();
                                }
                                catch
                                {
                                    exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel avançar para o diapositivo " + intpart);
                                    mmiC.Send(exNot);
                                    //tts.Speak("Desculpe, não é possivel avançar para o diapositivo " + intpart);
                                }

                                break;

                            case "limit":

                                try
                                {
                                    slides = pptPresentation.Slides;
                                    int size = slides.Count;
                                    slideTimeLimit = Int32.Parse(intpart);
                                    Console.WriteLine(slideTimeLimit);
                                    Console.WriteLine(slideTimeLimit * min);

                                    double timePerSlide = slideTimeLimit / size;
                                    timerFull = new System.Timers.Timer(slideTimeLimit * min);
                                    timerFull.Elapsed += OnTimedEventFull;
                                    timerFull.Enabled = true;

                                    timerPerSlide = new System.Timers.Timer(timePerSlide * min);
                                    timerPerSlide.Elapsed += OnTimedEvent;
                                    timerPerSlide.Enabled = true;
                                    exNot = lce.ExtensionNotification("", "", 0, "Limite de " + intpart + " minutos definido");
                                    mmiC.Send(exNot);
                                    //tts.Speak("Limite de " + intpart + " minutos definido");

                                }
                                catch
                                {
                                    exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel definir limite");
                                    mmiC.Send(exNot);
                                    //tts.Speak("Desculpe, não é possivel definir limite");
                                }

                                break;

                            case "avn":

                                slides = pptPresentation.Slides;

                                try
                                {

                                    index++;
                                    slides[index].Select();
                                    if (timerPerSlide != null)
                                    {
                                        if (timerPerSlide.Enabled)
                                        {
                                            timerPerSlide.Stop();
                                            timerPerSlide.Start();
                                        }
                                        else
                                        {
                                            timerPerSlide.Start();
                                        }
                                    }
                                    //restartSre();
                                }
                                catch
                                {
                                    index--;
                                    exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel avançar para o diapositivo seguinte");
                                    mmiC.Send(exNot);
                                    //tts.Speak("Desculpe, não é possivel avançar para o diapositivo seguinte");
                                }

                                break;

                            case "rec":
                                slides = pptPresentation.Slides;
                                try
                                {
                                    index--;
                                    slides[index].Select();
                                    //restartSre();
                                }
                                catch
                                {
                                    index++;
                                    exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possivel recuar para o diapositivo anterior");
                                    mmiC.Send(exNot);
                                    //tts.Speak("Desculpe, não é possivel recuar para o diapositivo anterior");
                                }
                                break;


                            case "acab":

                                try
                                {
                                    pptPresentation.Close();
                                    //restartSre();
                                }
                                catch
                                {
                                    exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possível terminar a apresentação");
                                    mmiC.Send(exNot);
                                    //tts.Speak("Desculpe, não é possível terminar a apresentação");
                                }
                                break;

                            case "adi":
                                try
                                {
                                    slides = pptPresentation.Slides;
                                    slides.Add(pptPresentation.Slides.Count + 1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                                    //restartSre();
                                }
                                catch
                                {
                                    exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possível adicionar slide");
                                    mmiC.Send(exNot);
                                    //tts.Speak("Desculpe, não é possível adicionar slide");
                                }

                                break;

                            case "rem":
                                try
                                {
                                    slides = pptPresentation.Slides;
                                    slides[index].Delete();
                                    index--;
                                    //restartSre();
                                }
                                catch
                                {
                                    exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não é possível remover slide");
                                    mmiC.Send(exNot);
                                    //tts.Speak("Desculpe, não é possível remover slide");
                                }

                                break;

                            case "grdrppt":
                                try
                                {
                                    pptPresentation.SaveAs("temp", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                                    //restartSre();
                                }
                                catch
                                {
                                    exNot = lce.ExtensionNotification("", "", 0, "Desculpe, não foi possivel guardar a apresentação");
                                    mmiC.Send(exNot);
                                    //tts.Speak("Desculpe, não foi possivel guardar a apresentação");
                                }
                                break;



                            default:
                                exNot = lce.ExtensionNotification("", "", 0, "Por favor use o comando 'Powerpoint abrir' para iniciar a aplicação");
                                mmiC.Send(exNot);
                                // tts.Speak("Por favor use o comando 'Powerpoint abrir' para iniciar a aplicação");
                                break;

                        }
                    }
                }
            } catch {
            }
            }
        private void OnTimedEventFull(object sender, ElapsedEventArgs e)
        {
            timerFull.Stop();
            timerFull.Enabled = false;
            
            exNot = lce.ExtensionNotification("", "", 0, "O tempo da apresentação esgotou-se");
            mmiC.Send(exNot);
            // tts.Speak("O tempo da apresentação esgotou-se");
        }

        private static void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            timerPerSlide.Stop();
            timerPerSlide.Enabled = false;
            
            exNot = lce.ExtensionNotification("", "", 0, "O tempo da apresentação por slide esgotou-se");
            mmiC.Send(exNot);
            //tts.Speak("O tempo da apresentação por slide esgotou-se");
        }
    }
}
