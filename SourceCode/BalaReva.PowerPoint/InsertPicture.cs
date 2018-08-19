namespace BalaReva.PowerPoint
{
    using Design;
    using Microsoft.Office.Core;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Insert Picture")]
    [Designer(typeof(InsertPictureDesign))]
    public class InsertPicture : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Select PPT File with path")]
        [DisplayName("File Path")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Select Image to insert")]
        [DisplayName("Image Path")]
        public InArgument<string> ImagePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Slide Index starts with one")]
        [DisplayName("Slide Index")]
        public InArgument<int> SlideIndex { get; set; } = 1;

        [Category("Input")]
        [RequiredArgument]
        public ImageSize PictureSize { get; set; } = new ImageSize();

        private string strFile { get; set; }
        private int intSlideIndex { get; set; }
        private string strImage { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                this.strFile = FilePath.Get(context);
                this.intSlideIndex = SlideIndex.Get(context);
                this.strImage = ImagePath.Get(context);

                this.DoInsertPicture();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private void DoInsertPicture()
        {
            this.Validate();

            Microsoft.Office.Interop.PowerPoint._Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();

            Microsoft.Office.Interop.PowerPoint._Presentation pptPresentation =
                pptApplication.Presentations.Open(strFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            pptPresentation.Slides[intSlideIndex].Shapes.AddPicture
                (this.strImage, MsoTriState.msoFalse, MsoTriState.msoTrue,
                PictureSize.Left, PictureSize.Top, PictureSize.Width, PictureSize.Height);

            pptPresentation.Save();
            pptPresentation.Close();
            pptApplication.Quit();

            
            this.releaseObject(pptPresentation);
            this.releaseObject(pptApplication);
        }

        private void Validate()
        {
            if (!System.IO.File.Exists(strFile))
            {
                throw new Exception("File is not exists");
            }

            else if (!System.IO.File.Exists(strImage))
            {
                throw new Exception("Invalid Image Path");
            }
        }


        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
