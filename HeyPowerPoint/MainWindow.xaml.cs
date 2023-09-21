using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.CognitiveServices.Speech;

namespace HeyPowerPoint
{
    public partial class MainWindow : Window
    {
        private string subscriptionKey = "";
        private string region = "";
        private PowerPoint.Application powerPointApp;
        private PowerPoint.Presentation presentation;
        private PowerPoint.SlideShowWindow slideShowWindow;
        private readonly SpeechRecognizer recognizer;
        private List<string> transitionTriggers = new List<string>
        {
            "start",
            "next",
            "back",
            "end"
        };
        private int slideCount;

        public MainWindow()
        {
            InitializeComponent();

            var speechConfig = SpeechConfig.FromSubscription(subscriptionKey, region);
            recognizer = new SpeechRecognizer(speechConfig);

            recognizer.Recognized += Recognizer_Recognized;
            recognizer.Canceled += Recognizer_Canceled;
        }

        private void Recognizer_Recognized(object sender, SpeechRecognitionEventArgs e)
        {
            if (e.Result.Reason == ResultReason.RecognizedSpeech)
            {
                string recognizedText = e.Result.Text.ToLowerInvariant().TrimEnd('.');
                Dispatcher.Invoke(() =>
                {
                    RecognizedTextBox.Text = recognizedText;

                    if (CustomPhraseMatch(recognizedText))
                    {
                        CustomTransition(recognizedText);
                        return;
                    }

                    int slideNumber;
                    if (SlideNumberMatch(recognizedText, out slideNumber))
                    {
                        SlideNumberTransition(slideNumber);
                        return;
                    }

                    RecognizedTextBox.Text = $"Recognized phrase: `{recognizedText}` doesn't match any predefined patterns";
                });
            }
            else if (e.Result.Reason == ResultReason.NoMatch)
            {
                Dispatcher.Invoke(() =>
                {
                    RecognizedTextBox.Text = "No speech recognized.";
                });
            }
        }

        private void Recognizer_Canceled(object sender, SpeechRecognitionCanceledEventArgs e)
        {
            if (e.Reason == CancellationReason.Error)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"Recognition canceled. Reason: {e.Reason}\nError Details: {e.ErrorDetails}");
                });
            }
        }

        private bool CustomPhraseMatch(string recognizedText)
        {
            return transitionTriggers.Any(phrase => recognizedText.Contains(phrase));
        }

        private void CustomTransition(string recognizedText)
        {
            if (recognizedText.Contains("start"))
            {
                Dispatcher.Invoke(() =>
                {
                    StartSlideshowButton_Click(this, new RoutedEventArgs());
                });
            }
            else if (recognizedText.Contains("next"))
            {
                Dispatcher.Invoke(() =>
                {
                    NextSlideshowButton_Click(this, new RoutedEventArgs());
                });
            }
            else if (recognizedText.Contains("back"))
            {
                Dispatcher.Invoke(() =>
                {
                    PreviousSlideshowButton_Click(this, new RoutedEventArgs());
                });
            }
            else if (recognizedText.Contains("end"))
            {
                Dispatcher.Invoke(() =>
                {
                    EndSlideshowButton_Click(this, new RoutedEventArgs());
                });
            }
            else
            {
                Console.WriteLine($"RecognizedText `{recognizedText}` doesn't contain any of the predefined patterns");
            }
        }

        private bool SlideNumberMatch(string recognizedText, out int slideNumber)
        {
            Console.WriteLine($"SlideNumberMatch: {recognizedText}");
            return int.TryParse(recognizedText, out slideNumber);
        }

        private void SlideNumberTransition(int slideNumber)
        {
            if (slideNumber < 1 || slideNumber > slideCount)
            {
                MessageBox.Show($"Error: Slide number {slideNumber} is out of range.");
                return;
            }

            Dispatcher.Invoke(() =>
            {
                TransitionToSlide(slideNumber);
            });
        }

        private void TransitionToSlide(int i)
        {
            if (presentation == null)
            {
                MessageBox.Show("Error: No active PowerPoint presentation detected.");
                return;
            }

            PowerPoint.SlideShowWindow slideShowWindow = presentation.SlideShowWindow;
            if (slideShowWindow == null)
            {
                MessageBox.Show("Error: No slide show detected.");
                return;
            }

            try
            {
                slideShowWindow.View.GotoSlide(i);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void ConnectPowerPointButton_Click(object sender, RoutedEventArgs e)
        {
            powerPointApp = (PowerPoint.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
            if (powerPointApp == null)
            {
                MessageBox.Show("Error: PowerPoint was not detected.");
                return;
            }

            presentation = powerPointApp.ActivePresentation;
            slideCount = presentation.Slides.Count;
            Console.WriteLine($"Slide count: {slideCount}");
            if (presentation == null)
            {
                MessageBox.Show("Error: No Active PowerPoint Presentation detected.");
                return;
            }
        }

        private async void StartSlideshowButton_Click(object sender, RoutedEventArgs e)
        {
            if (presentation == null)
            {
                await Task.Run(() => ConnectPowerPointButton_Click(sender, e));
            }

            slideShowWindow = presentation.SlideShowSettings.Run();
        }

        private void EndSlideshowButton_Click(object sender, RoutedEventArgs e)
        {
            if (presentation == null)
                return;

            if (slideShowWindow == null)
                return;

            slideShowWindow.View.Exit();
        }

        private void NextSlideshowButton_Click(object sender, RoutedEventArgs e)
        {
            if (presentation == null)
            {
                MessageBox.Show("Error: No active PowerPoint presentation detected.");
                return;
            }

            PowerPoint.SlideShowWindow slideShowWindow = presentation.SlideShowWindow;

            if (slideShowWindow == null)
            {
                MessageBox.Show("Error: No slide show detected.");
                return;
            }

            try
            {
                slideShowWindow.View.Next();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void PreviousSlideshowButton_Click(object sender, RoutedEventArgs e)
        {
            if (presentation == null)
            {
                MessageBox.Show("Error: No active PowerPoint presentation detected.");
                return;
            }

            PowerPoint.SlideShowWindow slideShowWindow = presentation.SlideShowWindow;

            if (slideShowWindow == null)
            {
                MessageBox.Show("Error: No slide show detected.");
                return;
            }

            try
            {
                slideShowWindow.View.Previous();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private async void StartListeningButton_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                await recognizer.StartContinuousRecognitionAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void StopListeningButton_Click(object sender, RoutedEventArgs e)
        {
            recognizer.StopContinuousRecognitionAsync();
        }
    }
}
