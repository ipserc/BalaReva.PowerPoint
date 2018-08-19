using BalaReva.PowerPoint.Design;
using Microsoft.Office.Core;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BalaReva.PowerPoint
{
    [DisplayName("Delete Slide")]
    [Designer(typeof(InsertSlideDesign))]
    public class DeleteSlide : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Select PPT File with path")]
        [DisplayName("File Path")]
        public InArgument<string> FilePath { get; set; }


        [Category("Input")]
        [RequiredArgument]
        [Description("Index start wiht one")]
        [DisplayName("Slide Index")]
        public InArgument<int> SlideIndex { get; set; } = 1;

       // public InArgument<double> X;
        ///float y, float width, float height



        private string strFile { get; set; }
        private int intSlideIndex { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {

                RectangleF rect = new RectangleF(50, 100, 600, 400);

                this.strFile = FilePath.Get(context);
                this.intSlideIndex = SlideIndex.Get(context);

                this.DoDeleteSlide();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoDeleteSlide()
        {
            this.Validate();

            Microsoft.Office.Interop.PowerPoint._Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();

            Microsoft.Office.Interop.PowerPoint._Presentation pptPresentation =
                pptApplication.Presentations.Open(strFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            if (pptPresentation.Slides.Count > intSlideIndex)
            {
                pptPresentation.Slides[intSlideIndex].Delete();

                pptPresentation.Save();
                pptPresentation.Close();
                pptApplication.Quit();

                this.releaseObject(pptPresentation);
                this.releaseObject(pptApplication);
            }
            else
            {
                pptPresentation.Close();
                pptApplication.Quit();

                this.releaseObject(pptPresentation);
                this.releaseObject(pptApplication);

                throw new Exception("Invalid slide index");
            }
        }

        private void Validate()
        {
            if (!System.IO.File.Exists(strFile))
            {
                throw new Exception("File is not exists");
            }
            else if (intSlideIndex <= 0)
            {
                throw new Exception("Invalid slide index");
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
