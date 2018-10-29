namespace BalaReva.PowerPoint
{
    using BalaReva.PowerPoint.Design;
    using Microsoft.Office.Core;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using PowerPointObj = Microsoft.Office.Interop.PowerPoint;

    [DisplayName("Replace Text")]
    [Designer(typeof(ReplaceTextDesign))]
    [Description("Replace a text along the whole presentation, or for a specific slide, with another given one")]
    public class ReplaceText : BasePowerPoint
    {
        [Category("Input"), RequiredArgument]
        [Description("Text to find in the presentation")]
        [DisplayName("Text to find")]
        public InArgument<string> TextToFind { get; set; }

        [Category("Input"), RequiredArgument]
        [Description("Text to replace in the presentation")]
        [DisplayName("Text to replace")]
        public InArgument<string> TextToReplace { get; set; }

        [Category("Input"), RequiredArgument]
        [Description("Slide Index starts with one. Zero to replace in all the slides")]
        [DisplayName("Slide Index")]
        public InArgument<int> SlideIndex { get; set; } = 0;

        private string strTextToFind;
        private string strTextToReplace;
        private int intSlideIndex;

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadValues(context);

                this.strTextToFind = TextToFind.Get(context);
                this.strTextToReplace = TextToReplace.Get(context);
                this.intSlideIndex = SlideIndex.Get(context) > 0 ? SlideIndex.Get(context) : -1;

                this.DoReplaceText();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoReplaceText()
        {
            try
            {
                base.InitPresentation();

                if (PptPresentation != null)
                {
                    // We've got a real presentation. We can operate with it
                    // Guess the replacement is for all the slides
                    int firstSlideNbr = 1;
                    int lastSlideNbr = PptPresentation.Slides.Count;
                    PptPresentation.PageSetup.FirstSlideNumber = firstSlideNbr;
                    if (intSlideIndex <= lastSlideNbr)
                    {
                        // If there is a specific slide we need to change the scope of the replacement
                        if (this.intSlideIndex != -1)
                        {
                            firstSlideNbr = this.intSlideIndex;
                            lastSlideNbr = this.intSlideIndex;
                        }
                        for (int slideNbr = firstSlideNbr; slideNbr <= lastSlideNbr; ++slideNbr)
                        {
                            foreach (PowerPointObj.Shape shape in PptPresentation.Slides[slideNbr].Shapes)
                            {
                                if (shape.HasTextFrame == MsoTriState.msoTrue)
                                {
                                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                                    {
                                        PowerPointObj.TextRange textRange = shape.TextFrame.TextRange;
                                        if (textRange.Text.Contains(this.strTextToFind))
                                        {
                                            textRange.Text = textRange.Text.Replace(this.strTextToFind, this.strTextToReplace);
                                        }
                                    }
                                }
                            }
                        }
                        base.SavePresentation();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
