using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Speech.Synthesis;

namespace PowerPointSetAudio
{
    public partial class Ribbon1
    {
        private int selctedRate = 0;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            for (int i = -10; i <= 10; i++)
            {
                var item = this.Factory.CreateRibbonDropDownItem();
                item.Label = i.ToString();
                this.speedBox.Items.Add(item);
            }

            this.speedBox.Text = "2".ToString();

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            // iterate all slides in current opened presentation
            foreach (PowerPoint.Slide slide in Globals.ThisAddIn.Application.ActivePresentation.Slides)
            {

                var synthesizer = new SpeechSynthesizer();

                var noteText = slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text;
                // generate a filename for the audio file
                var fn = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".wav";

                // get a audio stream
                synthesizer.SetOutputToWaveFile(fn);
                // call text to speech


                synthesizer.Rate = selctedRate;
                synthesizer.Speak(noteText);
                // add audio shape to slide
                var audioShape = slide.Shapes.AddMediaObject2(fn, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 100, 100);
                audioShape.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue;
                audioShape.AnimationSettings.PlaySettings.HideWhileNotPlaying = MsoTriState.msoTrue;

                // update progressbar in ribbon


            }
        }

        private void buttonExportMP4_Click(object sender, RibbonControlEventArgs e)
        {
            // export current file to mp4 with the same name 
            var path = Globals.ThisAddIn.Application.ActivePresentation.FullName;
            var newPath = path.Replace(".pptx", ".mp4");
            Globals.ThisAddIn.Application.ActivePresentation.SaveCopyAs(newPath, PowerPoint.PpSaveAsFileType.ppSaveAsMP4);
        }

        private void buttonClearAudio_Click(object sender, RibbonControlEventArgs e)
        {
            var audioCount = 0;
            // iterate all slides in current opened presentation
            foreach (PowerPoint.Slide slide in Globals.ThisAddIn.Application.ActivePresentation.Slides)
            {
                // iterate all shapes in current slide
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    // check if shape is an audio shape
                    if (shape.Type == MsoShapeType.msoMedia)
                    {
                        // cast shape to audio shape
                        if (shape.MediaType == PowerPoint.PpMediaType.ppMediaTypeSound)
                        {
                            // remove audio shape
                            shape.Delete();
                            audioCount++;
                        }
                    }
                }
            }
            // Show messagebox to say finished, and how many slides were processed
            var msg = $"{audioCount} audio files deleted in {Globals.ThisAddIn.Application.ActivePresentation.Slides.Count} slides";
            System.Windows.Forms.MessageBox.Show(msg);
        }

        private void speedBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var ret = this.speedBox.Text;
            try
            {
                selctedRate = int.Parse(ret);
            }
            catch (Exception ex)
            {
                // do nothing
            };


        }
    }
}
