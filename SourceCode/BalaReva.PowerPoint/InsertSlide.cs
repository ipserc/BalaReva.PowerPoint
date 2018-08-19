namespace BalaReva.PowerPoint
{
    using BalaReva.PowerPoint.Design;
    using Microsoft.Office.Core;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Insert Slide")]
    [Designer(typeof(InsertSlideDesign))]
    public class InsertSlide : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Select PPT File with path")]
        [DisplayName("File Path")]
        public InArgument<string> FilePath { get; set; }

        private string strFile { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                this.strFile = FilePath.Get(context);

                this.DoInsertSlide();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoInsertSlide()
        {
            this.Validate();

            Microsoft.Office.Interop.PowerPoint._Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();

            Microsoft.Office.Interop.PowerPoint._Presentation pptPresentation =
                pptApplication.Presentations.Open(strFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout =
                pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            pptPresentation.Slides.AddSlide(pptPresentation.Slides.Count + 1, customLayout);

            pptPresentation.Save();
            pptPresentation.Close();
            pptApplication.Quit();

            this.releaseObject(customLayout);
            this.releaseObject(pptPresentation);
            this.releaseObject(pptApplication);
        }

        private void Validate()
        {
            if (!System.IO.File.Exists(strFile))
            {
                throw new Exception("File is not exists");
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
