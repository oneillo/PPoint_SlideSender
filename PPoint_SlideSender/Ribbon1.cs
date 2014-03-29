using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

namespace PPoint_SlideSender
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // Hide any alerts at the start. Will unhide at the end
             Globals.SlideSender.Application.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
            //PowerPoint.Application app = Globals.SlideSender.Application;
            //app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
            
            
            
            // Create a new file name that will be used to save the temp presentation based on the old presentations path
            //string fileName = Globals.SlideSender.Application.ActivePresentation.Path + "\\" + Globals.SlideSender.Application.ActivePresentation.Name;
            string fileName = @"C:\Temp\" + Globals.SlideSender.Application.ActivePresentation.Name;
            string fileNameNoExt = Path.GetFileNameWithoutExtension(fileName);
            //string newFileName = Globals.SlideSender.Application.ActivePresentation.Path + "\\" + fileNameNoExt + "_slides";
            string newFileName = @"C:\Temp\" + fileNameNoExt + "_slides";
            PowerPoint.Presentation pres = Globals.SlideSender.Application.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
            //Trace.WriteLine("Created a new presentation");
            //PowerPoint.CustomLayout layout = pres.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutTitle];

            // Globals.SlideSender.Application.ActiveWindow.Selection.SlideRange.Count; 
            // Now go through and add the selected slides into the new presentation
            int[] selectedSlides = new int[Globals.SlideSender.Application.ActiveWindow.Selection.SlideRange.Count];
            int count = 0;
            foreach (PowerPoint.Slide sld in Globals.SlideSender.Application.ActiveWindow.Selection.SlideRange)
            {
                selectedSlides[count] = sld.SlideNumber;
                //Trace.WriteLine("Selected slide: " + s);
                count++;
            }

            Array.Sort(selectedSlides);
            foreach (int slideNumber in selectedSlides.Reverse())
            {
                Globals.SlideSender.Application.ActivePresentation.Slides[slideNumber].Copy();
                pres.Slides.Paste(1);
            }
            

            // Save the new presentation file name
            //Trace.WriteLine(" ");
            //Trace.WriteLine("Saved presentation to " + newFileName);
            //Trace.WriteLine(" ");
            pres.SaveAs(newFileName, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Office.MsoTriState.msoTrue);
            string attachmentFileName = pres.Path + "\\" + pres.Name;
            string pptFileName = pres.Name;
            pres.Close();

            //Wait until file is available - in case it takes a moment for it to sync from cloud
            /*bool fileDoesNotExist = !File.Exists(attachmentFileName);
            int fileCheckCount = 0;
            while(fileDoesNotExist && fileCheckCount < 500)
            {
                Trace.WriteLine("Checking if the file exists"); 
                fileDoesNotExist = File.Exists(attachmentFileName);
                fileCheckCount++;
            }

            
            if(fileDoesNotExist)
            {
                System.Windows.Forms.MessageBox.Show("Wasn't able to create the temporary presentation to email out. Is your hard disk full?");
                Environment.Exit(0);
            }*/
            
             
            Outlook.Application outlook = null;
        
            //string fileName = CreatePresentation();
            try
            {
                object app = Marshal.GetActiveObject("Outlook.Application");
                outlook = app as Outlook.Application;
            }
            catch (COMException)
            {
                if (outlook == null)
                {
                    Trace.WriteLine("Couldn't get pointer to existing Outlook application. Opening a new one");
         
                    outlook = new Outlook.Application();
                }
            }

                if (outlook != null)
                {
                    //Trace.WriteLine("Sending the email now");
                    //string attachmentFileName = Globals.SlideSender.Application.ActivePresentation.Path + "\\" + Globals.SlideSender.Application.ActivePresentation.Name;
                    

                    Outlook.MailItem mail = (Outlook.MailItem)outlook.CreateItem(Outlook.OlItemType.olMailItem);
                    //mail.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
                    //mail.Body = "This is a test";

                    //Trace.WriteLine("attachmentFileName =" + attachmentFileName);
                    mail.Attachments.Add((object)attachmentFileName, Outlook.OlAttachmentType.olOLE, 1, pptFileName);
                    //mail.Attachments.Add(new Outlook.Attachment(newFileName));
                    mail.Display(false);

                    //mail.Attachments.Add((object)newFileName, (object)Outlook.OlAttachmentType.olOLE, (object)1, (object)"Nothing");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Unable to process the request!");
                }
            
            
            // Delete the temporary presentation that was created with the selected slides
            System.IO.File.Delete(attachmentFileName);

            // Start showing alerts from powerpoint again
            Globals.SlideSender.Application.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;

            // Now force clean up all of the COMs objects
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

       /* private string CreatePresentation()
        {
            string fileName = Globals.SlideSender.Application.ActivePresentation.Path + Globals.SlideSender.Application.Name + "_excerpt";
            PowerPoint.Presentation pres = Globals.SlideSender.Application.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
            Console.WriteLine("A new presentation is created");
            //PowerPoint.CustomLayout layout = pres.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutTitle];

            Console.WriteLine("Saved presentation to " + fileName);
            pres.SaveAs(fileName, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Office.MsoTriState.msoTrue);

            return fileName;
        }
        */
    }
}
