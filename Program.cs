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

    class Program
    {
        static ISlide slide1, s2;

        static void Main(string[] args)
        {

            //Open the presentation
            string path = "auxi C# Interview-1.pptx";
            IPresentation powerpointDoc = Presentation.Open(path);
            slide1 = powerpointDoc.Slides[0];
            s2 = powerpointDoc.Slides[1];
            /////////////////////////////////

            IShape shapee = slide1.Shapes[5] as IShape;
            slide1.Shapes.Remove(shapee);
            shapee = slide1.Shapes[5] as IShape;
            slide1.Shapes.Remove(shapee);
            shapee = slide1.Shapes[5] as IShape;
            slide1.Shapes.Remove(shapee);
            shapee = slide1.Shapes[5] as IShape;
            slide1.Shapes.Remove(shapee);
            Getslide2();
            Console.WriteLine("done");
            powerpointDoc.Save("Resaved.pptx");
            powerpointDoc.Close();
            System.Console.Read();

        }

        private static void Getslide2()
        {
            Slide2 newslide = new Slide2(s2);

            int i = 0;
            foreach (IShape shape in slide1.Shapes)
            {
                //Console.WriteLine(shape.ShapeName);
                if (shape.ShapeName.Contains("TextBox 2") || shape.ShapeName.Contains("TextBox 3") || shape.ShapeName.Contains("TextBox 4") || shape.ShapeName.Contains("TextBox 5"))
                {
                    //slide1.Shapes.Remove(shape);

                }
                if (shape.ShapeName.Contains("Title"))
                {

                    shape.TextBody.Text = newslide.titleText;
                    //titleFont = shape.TextBody;
                    shape.Width = newslide.titleW;
                    shape.Height = newslide.titleH;
                    shape.Top = newslide.titleTop;
                    shape.Left = newslide.titleLeft;
                    shape.TextBody.Paragraphs[0].HorizontalAlignment = HorizontalAlignmentType.Center;
                    IParagraph paragraph = shape.TextBody.Paragraphs[0];
                    ITextPart textPart = paragraph.TextParts[0];
                    //textPart.Font.Bold = true;
                    textPart.Font.FontName = "Beirut";




                    /////////////////////////////////////
                    //shape.TextFrame.TextRange.Words(3).Font.Bold = true;
                    continue;

                }

                if (shape.ShapeName.Contains("Arrow"))
                {
                    if (i == 0)
                    {
                        shape.Height = newslide.arrH;
                        shape.Width = newslide.arrW;
                        shape.Top = newslide.arrTop;
                        shape.Left = newslide.arr1Left;
                        shape.TextBody.Text = newslide.arr1;


                    }
                    else if (i == 1)
                    {
                        shape.Height = newslide.arrH;
                        shape.Width = newslide.arrW;
                        shape.Top = newslide.arrTop;
                        shape.Left = newslide.arr2Left;
                        shape.TextBody.Text = newslide.arr2;


                    }
                    else if (i == 2)
                    {
                        shape.Height = newslide.arrH;
                        shape.Width = newslide.arrW;
                        shape.Top = newslide.arrTop;
                        shape.Left = newslide.arr3Left;
                        shape.TextBody.Text = newslide.arr3;

                    }
                    else if (i == 3)
                    {
                        shape.Height = newslide.arrH;
                        shape.Width = newslide.arrW;
                        shape.Top = newslide.arrTop;
                        shape.Left = newslide.arr4Left;
                        shape.TextBody.Text = newslide.arr4;

                    }

                    else
                    {
                        continue;
                    }
                    i++;
                }
                if (shape.ShapeName.Contains("TextBox 6"))
                {
                    shape.Top = newslide.ty;
                    shape.Left = newslide.t1Left;

                    IParagraph paragraph = shape.TextBody.Paragraphs[0];
                    ITextPart textPart = paragraph.TextParts[0];
                    textPart.Font.FontName = "Beirut";
                    textPart.Font.Bold = false;
                    paragraph.ListFormat.Type = ListType.Bulleted;
                    //paragraph.ListFormat.Type = ListType.Bulleted;
                    textPart.Font.Underline = 0;
                    paragraph.Font.Bold = false;

                }
                if (shape.ShapeName.Contains("TextBox 12"))
                {
                    shape.Top = newslide.ty;
                    shape.Left = newslide.t2Left;

                    IParagraph paragraph = shape.TextBody.Paragraphs[0];
                    ITextPart textPart = paragraph.TextParts[0];
                    //textPart.Font.Bold = true;
                    textPart.Font.FontName = "Beirut";
                    //textPart.Font.Bold = false;
                    paragraph.Font.Bold = false;

                    paragraph.ListFormat.Type = ListType.Bulleted;
                    paragraph.ListFormat.NumberStyle = NumberedListStyle.ThaiNumPeriod;

                }
                if (shape.ShapeName.Contains("TextBox 14"))
                {
                    shape.Top = newslide.ty;
                    shape.Left = newslide.t3Left;

                    IParagraph paragraph = shape.TextBody.Paragraphs[0];
                    ITextPart textPart = paragraph.TextParts[0];
                    //textPart.Font.Bold = true;
                    textPart.Font.FontName = "Beirut";
                    //textPart.Font.Bold = false;
                    paragraph.Font.Bold = false;

                    paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicDbPeriod;
                    
                }
                if (shape.ShapeName.Contains("TextBox 15"))
                {
                    shape.Top = newslide.ty;
                    shape.Left = newslide.t4Left;

                    IParagraph paragraph = shape.TextBody.Paragraphs[0];
                    ITextPart textPart = paragraph.TextParts[0];

                    //textPart.Font.Bold = true;
                    textPart.Font.FontName = "Beirut";
                    textPart.Font.Bold = false;
                    
                    paragraph.Font.Bold = false;


                }

                continue;


            }
            
        }
    }
}
