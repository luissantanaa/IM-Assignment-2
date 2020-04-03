using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using mmisharp;
using Microsoft.Speech.Recognition;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using Newtonsoft.Json;

namespace speechModality
{
    public class SpeechMod
    {
        private SpeechRecognitionEngine sre;
        private Grammar gr;
        private Tts tts = new Tts();
        private Microsoft.Office.Interop.PowerPoint.Application PPTAPP = new Microsoft.Office.Interop.PowerPoint.Application();
        private Presentation pptPresentation;
        private Slides slides;
        private Boolean WOKE = false;
        private Boolean OPENED = false;
        private int index = 1;
        private SlideShowView objSlideShowView;

        public event EventHandler<SpeechEventArg> Recognized;
        protected virtual void onRecognized(SpeechEventArg msg)
        {
            EventHandler<SpeechEventArg> handler = Recognized;
            if (handler != null)
            {
                handler(this, msg);
            }
        }

        private LifeCycleEvents lce;
        private MmiCommunication mmic;

        public SpeechMod()
        {
            //init LifeCycleEvents..
            lce = new LifeCycleEvents("ASR", "FUSION","speech-1", "acoustic", "command"); // LifeCycleEvents(string source, string target, string id, string medium, string mode)
            //mmic = new MmiCommunication("localhost",9876,"User1", "ASR");  //PORT TO FUSION - uncomment this line to work with fusion later
            mmic = new MmiCommunication("localhost", 8000, "User1", "ASR"); // MmiCommunication(string IMhost, int portIM, string UserOD, string thisModalityName)

            mmic.Send(lce.NewContextRequest());
            
            //load pt recognizer
            sre = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("pt-PT"));
            gr = new Grammar(Environment.CurrentDirectory + "\\ptG.grxml", "rootRule");
            sre.LoadGrammar(gr);

            
            sre.SetInputToDefaultAudioDevice();
            sre.RecognizeAsync(RecognizeMode.Multiple);
            sre.SpeechRecognized += Sre_SpeechRecognized;
            sre.SpeechHypothesized += Sre_SpeechHypothesized;

        }

        private void Sre_SpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            onRecognized(new SpeechEventArg() { Text = e.Result.Text, Confidence = e.Result.Confidence, Final = false });
        }

        private void Sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            int ID;
            onRecognized(new SpeechEventArg(){Text = e.Result.Text, Confidence = e.Result.Confidence, Final = true});
            
            //SEND
            // IMPORTANT TO KEEP THE FORMAT {"recognized":["SHAPE","COLOR"]}
            string json = "{ \"recognized\": [";
            foreach (var resultSemantic in e.Result.Semantics)
            {
                json+= "\"" + resultSemantic.Value.Value +"\", ";
            }
            json = json.Substring(0, json.Length - 2);
            json += "] }";

            var exNot = lce.ExtensionNotification(e.Result.Audio.StartTime+"", e.Result.Audio.StartTime.Add(e.Result.Audio.Duration)+"",e.Result.Confidence, json);
            mmic.Send(exNot);

            if (e.Result.Confidence < 0.5)
            {
                tts.Speak("Desculpe, não percebi. Por favor repita");
            }
            else
            {
                var wake = e.Result.Semantics.First().Value.Value;
                var command = e.Result.Semantics.Last().Value.Value;
                Console.WriteLine(wake);
                Console.WriteLine(command);

                if (WOKE && OPENED)
                {
                    switch (command)
                    {
                        case "nAprest":
                            if (!(PPTAPP.SlideShowWindows.Count > 0))
                            {
                                tts.Speak("De momento não esta no modo de apresentação");
                            }
                            else
                            {
                                objSlideShowView.Exit();
                            }
                            break;
                        case "aprest":
                            if (slides.Count.Equals(0))
                            {
                                tts.Speak("A apresentação tem que ter pelo menos um slide");
                            }
                            else
                            {
                                pptPresentation.SlideShowSettings.ShowPresenterView = MsoTriState.msoFalse;
                                pptPresentation.SlideShowSettings.Run();
                                objSlideShowView = pptPresentation.SlideShowWindow.View;
                                objSlideShowView.Application.SlideShowWindows[1].Activate();
                            }

                            break;

                        case "avn":
                            //pptPresentation.SlideShowSettings.ShowPresenterView = MsoTriState.msoFalse;
                            ////Run the presentation
                            //pptPresentation.SlideShowSettings.Run();
                            ////Hold a reference to the SlideShowWindow
                            //SlideShowView objSlideShowView = pptPresentation.SlideShowWindow.View;
                            //objSlideShowView.Application.SlideShowWindows[1].Activate();
                            //objSlideShowView.Next();
                            slides = pptPresentation.Slides;
                            try
                            {
                                index++;
                                slides[index].Select();
                            }
                            catch
                            {
                                index--;
                                tts.Speak("Desculpe, não é possivel avançar para o diapositivo seguinte");
                            }
                            
                            break;

                        case "rec":
                            slides = pptPresentation.Slides;
                            try
                            {
                                index--;
                                slides[index].Select();
                            }
                            catch
                            {
                                index++;
                                tts.Speak("Desculpe, não é possivel recuar para o diapositivo anterior");
                            }
                            break;

                        case "abr":
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
                            break;

                        case "adi":
                            slides = pptPresentation.Slides;
                            slides.Add(pptPresentation.Slides.Count + 1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
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
                    }
                    else
                    {
                        tts.Speak("Por favor, use o comando 'abrir' para iniciar a aplicação");
                    }

                }
                else if (!(WOKE) && !(OPENED))
                {
                    if (!wake.Equals("") && !command.Equals(""))
                    {
                        switch (command){ 

                            case "abr":
                                try
                                {
                                    OPENED = true;
                                    PPTAPP.Visible = MsoTriState.msoTrue;
                                    Presentations ppPresens = PPTAPP.Presentations;
                                    pptPresentation = ppPresens.Open("temp", MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
                                }
                                catch (System.IO.FileNotFoundException)
                                {
                                    pptPresentation = PPTAPP.Presentations.Add(MsoTriState.msoTrue);
                                }
                                break;

                            default:
                                tts.Speak("Por favor use o comando 'Powerpoint abrir' para iniciar a aplicação");
                                break;
                            //case "adi":
                            //    Slides slides = pptPresentation.Slides;
                            //    slides.Add(pptPresentation.Slides.Count + 1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                            //    break;
                        }
                    }

                    if (wake.Equals("ppt") && command.Equals(""))
                    {
                        tts.Speak("Sim?");
                        WOKE = true;
                    }
                }
                else
                {
                    if (!wake.Equals("") && !command.Equals(""))
                    {
                        switch (command)
                        {

                            case "nAprest":
                                if (!(PPTAPP.SlideShowWindows.Count > 0)) {
                                    tts.Speak("De momento não esta no modo de apresentação");
                                }
                                else
                                {
                                    objSlideShowView.Exit();
                                }
                                break;
                            case "aprest":
                                if (slides == null)
                                {
                                    tts.Speak("A apresentação tem que ter pelo menos um slide");
                                }
                                else
                                {
                                    pptPresentation.SlideShowSettings.ShowPresenterView = MsoTriState.msoFalse;
                                    pptPresentation.SlideShowSettings.Run();
                                    objSlideShowView = pptPresentation.SlideShowWindow.View;
                                    objSlideShowView.Application.SlideShowWindows[1].Activate();
                                }
                               
                                break;

                            case "avn":
                               
                                slides = pptPresentation.Slides;
                                try
                                {
                                    index++;
                                    slides[index].Select();
                                }
                                catch
                                {
                                    index--;
                                    tts.Speak("Desculpe, não é possivel avançar para o diapositivo seguinte");
                                }

                                break;

                            case "rec":
                                slides = pptPresentation.Slides;
                                try
                                {
                                    index--;
                                    slides[index].Select();
                                }
                                catch
                                {
                                    index++;
                                    tts.Speak("Desculpe, não é possivel recuar para o diapositivo anterior");
                                }
                                break;

                            case "abr":
                                try
                                {
                                    OPENED = true;
                                    PPTAPP.Visible = MsoTriState.msoTrue;
                                    Presentations ppPresens = PPTAPP.Presentations;
                                    pptPresentation = ppPresens.Open("temp", MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
                                }
                                catch (System.IO.FileNotFoundException)
                                {
                                    pptPresentation = PPTAPP.Presentations.Add(MsoTriState.msoTrue);
                                }
                                break;

                            case "adi":
                                slides = pptPresentation.Slides;
                                slides.Add(pptPresentation.Slides.Count + 1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                                break;

                            default:
                                tts.Speak("Por favor use o comando 'Powerpoint abrir' para iniciar a aplicação");
                                break;

                        }
                    }
                }
            }

            
            

            //if (e.Result.Semantics.ContainsKey("open")) {
            //    try
            //    {
            //        PPTAPP.Visible = MsoTriState.msoTrue;
            //        Presentations ppPresens = PPTAPP.Presentations;
            //        pptPresentation = ppPresens.Open("temp", MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
            //    }
            //    catch (System.IO.FileNotFoundException)
            //    {
            //        pptPresentation = PPTAPP.Presentations.Add(MsoTriState.msoTrue);
            //        //pptPresentation.SaveAs(@"temp", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            //    }
            //}

            //if (e.Result.Semantics.ContainsKey("command")) {
            //    switch (e.Result.Semantics["command"].Value.ToString())
            //    {
            //        case "avn":
            //            tts.Speak("Avançar Slide");

            //            //if (pptPresentation == null)
            //            //{
            //            //    try
            //            //    {
            //            //        var col = PPTAPP.Presentations.GetEnumerator();

            //            //        while (pptPresentation == null) {
            //            //            pptPresentation = (Presentation)col.Current;
            //            //            col.MoveNext();
            //            //        }
            //            //      }
            //            //    catch
            //            //    {

            //            //    }
            //            //}

            //            //try
            //            //{
            //            //    SlideShowWindow slides = pptPresentation.SlideShowWindow;

            //            //}
            //            //catch(System.Runtime.InteropServices.COMException)
            //            //{
            //            //    tts.Speak("A apresentação não tem mais slides");
            //            //}
            //            break;

            //        case "rec":
            //            tts.Speak("Recuar Slide");
            //            break;

            //        case "ppt":
            //            tts.Speak("Sim?");
            //            break;
            //    }
            //}



        }
    }
}
