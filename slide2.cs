using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.PowerBI.Api.Models;
using Syncfusion.Presentation;

namespace ConsoleApp1ToEditPPTX
{
    class Slide2
    {

        public string titleText;
        public double titleTop, titleLeft, titleW, titleH;
        public string arr1, arr2, arr3, arr4;
        public double arrTop, arr1Left, arr2Left, arr3Left, arr4Left;
        public double arrH, arrW;
        public double ty, t1Left, t2Left, t3Left, t4Left;

        //string 
        private Slide2() { }
        public Slide2(ISlide s2)
        {
            int i = 0;
            double coof = 1;
            foreach (IShape shape in s2.Shapes)
            {
                Console.WriteLine(shape.ShapeName);

                if (shape.ShapeName.Contains("Title 1"))
                {

                    titleText = shape.TextBody.Text;
                    //titleFont = shape.TextBody;
                    titleTop = shape.Top;
                    titleLeft = shape.Left;
                    titleW = shape.Width;
                    titleH = shape.Height;



                }

                if (shape.ShapeName.Contains("Arrow"))
                {
                    if (i == 0)
                    {
                        arrH = shape.Height * coof;
                        arrW = shape.Width * coof;
                        arrTop = shape.Top * coof;
                        arr1Left = shape.Left * coof;
                        arr1 = "Begin";
                    }
                    else if (i == 1)
                    {
                        arr2Left = shape.Left * coof;
                        arr2 = "Step 1";
                    }
                    else if (i == 2)
                    {
                        arr3Left = shape.Left * coof;
                        arr3 = "Step 1";

                    }
                    else if (i == 3)
                    {
                        arr4Left = shape.Left * coof;
                        arr4 = "Step 1";

                    }

                    else
                    {
                        //break;
                    }
                    i++;
                    //titleText = shape.TextBody.Text;
                    //titleFont = shape.TextBody;


                }


                if (shape.ShapeName.Contains("TextBox 6"))
                {
                    ty = shape.Top;
                    t1Left = shape.Left;
                }
                if (shape.ShapeName.Contains("TextBox 19"))
                {
                    t2Left = shape.Left;

                }
                if (shape.ShapeName.Contains("TextBox 20"))
                {
                    t3Left = shape.Left;

                }
                if (shape.ShapeName.Contains("TextBox 21"))
                {
                    t4Left = shape.Left;

                }



            }
            
        }
    }
}
