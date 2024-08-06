#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;

#endregion

namespace PipeTagger
{
    class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            //0. Get assemblyName & Img for button data
            String assembName = Assembly.GetExecutingAssembly().Location;

            #region ###Icon image import

            // Drag your image into projet properties -> Recources
            // Image: 32 pixel x 32 pixel, 96 DPI (https://convert.town/image-dpi)
            Image Img1 = Properties.Resources._1__管排標籤設定96;
            Image Img2 = Properties.Resources._2__放置管排管段標籤96;
            Image Img3 = Properties.Resources._3__管排高程標籤設定96;
            Image Img4 = Properties.Resources._4__放置管排定點高程96;
            Image Img5 = Properties.Resources._5__選取管段標籤96;
            Image Img6 = Properties.Resources._6__選取定點高程96;

            ImageSource imgSrc1 = Common.UI_Tool.GetImageSource(Img1);
            ImageSource imgSrc2 = Common.UI_Tool.GetImageSource(Img2);
            ImageSource imgSrc3 = Common.UI_Tool.GetImageSource(Img3);
            ImageSource imgSrc4 = Common.UI_Tool.GetImageSource(Img4);
            ImageSource imgSrc5 = Common.UI_Tool.GetImageSource(Img5);
            ImageSource imgSrc6 = Common.UI_Tool.GetImageSource(Img6);

            #endregion

            //1. Create a tab
            String TabName = "【CEC Tag】";

            try {application.CreateRibbonTab(TabName); }// get the ribbon tab
            catch (Exception) { } //tab alreadt exists


            //2. Create a panel
            RibbonPanel panel = application.CreateRibbonPanel(TabName, "【CEC Tag】");

            //3. Create Button data x6 
            PushButtonData btnData1 = new PushButtonData("Btn1", "管排標籤設定",     assembName, "PipeTagger.Tag_setting");  // Last input = "Proj name" + "Function name"
            PushButtonData btnData2 = new PushButtonData("Btn2", "放置管排管段標籤", assembName, "PipeTagger.Place_tag");    // Last input = "Proj name" + "Function name"
            //PushButtonData btnData3 = new PushButtonData("Btn3", "管排高程標籤設定", assembName, "PipeTagger.NotFinish");    // Last input = "Proj name" + "Function name"
            PushButtonData btnData4 = new PushButtonData("Btn4", "放置管排定點高程", assembName, "PipeTagger.Place_spot");   // Last input = "Proj name" + "Function name"
            //PushButtonData btnData5 = new PushButtonData("Btn5", "選取管段標籤",     assembName, "PipeTagger.Pick_tag");    // Last input = "Proj name" + "Function name"
            //PushButtonData btnData6 = new PushButtonData("Btn6", "選取定點高程",     assembName, "PipeTagger.NotFinish");    // Last input = "Proj name" + "Function name"

            #region ###Button imageSrc

            btnData1.LargeImage = imgSrc1;
            btnData2.LargeImage = imgSrc2;
            //btnData3.LargeImage = imgSrc3;
            btnData4.LargeImage = imgSrc4;
            //btnData5.LargeImage = imgSrc5;
            //btnData6.LargeImage = imgSrc6;

            #endregion

            //4. Arrange buttons' position in panel

            panel.AddItem(btnData1);
            panel.AddItem(btnData2);
            //panel.AddItem(btnData3);
            panel.AddItem(btnData4);
            //panel.AddSeparator();
            //panel.AddItem(btnData5);
            //panel.AddItem(btnData6);
            
            return Result.Succeeded;
        }

        
       
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Failed;
        }


    }
}
