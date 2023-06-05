using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;



namespace Лючки
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SldWorks swApp;
        ModelDocExtension swModelDocExt;
        IModelDoc2 swModel;
        DrawingDoc swDrawing;
        SelectionMgr swSelmgr;
        SolidWorks.Interop.sldworks.View swView;
        ISketchManager swModel2;
        IModelDocExtension swModel3;



        //  ModelDoc2 swModel;


        private void button1_Click_1(object sender, EventArgs e)
        {

            Process[] processes = Process.GetProcessesByName("SLDWORKS");
            foreach (Process process in processes)
            {
                process.CloseMainWindow();
                process.Kill();
            }

            //Запуск SW
            Guid myGuid1 = new Guid("d134b411-3689-497d-b2d7-a27cb1066648");
            object processSW = System.Activator.CreateInstance(System.Type.GetTypeFromCLSID(myGuid1));

            swApp = (SldWorks)processSW;
            swApp.Visible = true;

            // Создание детали
            swApp.NewPart();
            swModel = swApp.IActiveDoc2;

            //Построение первого лючка

            swModel.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            swModel.ClearSelection2(true);

            double con = 1.240;
            double d = 1000;
            var len1 = (double)length1.Value / d;
            var len2 = (double)length2.Value / d;
            var len3 = (double)length3.Value / d;
            var len4 = (double)length4.Value / d;
            var len5 = (double)length5.Value / d;
            var len6 = (double)length6.Value / d;
            var len7 = (double)length7.Value / d;
            var len8 = (double)length8.Value / d;
            var hole1 = (double)dimHole1.Value / d;
            var hole2 = (double)dimHole2.Value / d;
            var hole3 = (double)dimHole3.Value / d;
            var hole4 = (double)dimHole4.Value / d;
            var hole5 = (double)dimHole5.Value / d;
            var hole6 = (double)dimHole6.Value / d;
            var hole7 = (double)dimHole7.Value / d;
            var hole8 = (double)dimHole8.Value / d;
            var slotLen1 = ((double)lenghtSlot1.Value / d) / 2;
            var slotLen2 = ((double)lenghtSlot2.Value / d) / 2;
            var slotLen3 = ((double)lenghtSlot3.Value / d) / 2;
            var slotLen4 = ((double)lenghtSlot4.Value / d) / 2;
            var slotLen5 = ((double)lenghtSlot5.Value / d) / 2;
            var slotLen6 = ((double)lenghtSlot6.Value / d) / 2;
            var slotLen7 = ((double)lenghtSlot7.Value / d) / 2;
            var slotLen8 = ((double)lenghtSlot8.Value / d) / 2;
            var heigthToCut1 = ((double)heightToCutout1.Value / d);
            var heigthToCut2 = ((double)heigthToCutout2.Value / d);
            var heigthToCut3 = ((double)heigthToCutout3.Value / d);
            var heigthToCut4 = ((double)heigthToCutout4.Value / d);
            var heigthToCut5 = ((double)heigthToCutout5.Value / d);
            var heigthToCut6 = ((double)heigthToCutout6.Value / d);
            var heigthToCut7 = ((double)heigthToCutout7.Value / d);
            var heigthToCut8 = ((double)heigthToCutout8.Value / d);
            var heigthCut1 = ((double)hegthCutout1.Value / d);
            var heigthCut2 = ((double)hegthCutout2.Value / d);
            var heigthCut3 = ((double)hegthCutout3.Value / d);
            var heigthCut4 = ((double)hegthCutout4.Value / d);
            var heigthCut5 = ((double)hegthCutout5.Value / d);
            var heigthCut6 = ((double)hegthCutout6.Value / d);
            var heigthCut7 = ((double)hegthCutout7.Value / d);
            var heigthCut8 = ((double)hegthCutout8.Value / d);
            var depthCut1 = ((double)depthCutout1.Value / d);
            var depthCut2 = ((double)depthCutout2.Value / d);
            var depthCut3 = ((double)depthCutout3.Value / d);
            var depthCut4 = ((double)depthCutout4.Value / d);
            var depthCut5 = ((double)depthCutout5.Value / d);
            var depthCut6 = ((double)depthCutout6.Value / d);
            var depthCut7 = ((double)depthCutout7.Value / d);
            var depthCut8 = ((double)depthCutout8.Value / d);




            var line = "Линия";
            var namberLine = 2;
            double wid = (double)width.Value / d;
            var pen = new double[2];
            bool per = false;
            var numberLuc = 1;
            


            // Перенос карандаша

            double[] TransprtPen(double[] pin, double x, bool pir)
            {
                if (pir == true)
                {
                    pin[0] = 0;
                    pin[1] = pin[1] + 0.005 + wid;
                }
                else
                {
                    pin[0] = pin[0] + x;
                    pin[1] = pin[1];
                }
                return pin;
            }


            // Построение отверстия
            void BuildHole(double[] pin, string hole, double x, double slot)
            {
                switch (hole)
                {
                    case "Прорезь":
                        swModel.SketchManager.CreateSketchSlot(1, 1, 0.114, pin[0] + x, pin[1] + (wid / 2), 0, pin[0] + x - slot, pin[1] + (wid / 2), 0, 0, 0, 0, 1, false);
                        break;
                    case "Отверстие":
                        swModel.SketchManager.CreateCircleByRadius(pin[0] + x, pin[1] + (wid / 2), 0, 0.057);
                        break;
                    case "нет":
                        break;

                }

            }


            // Построение лючка

            void BildLuc(bool Chek, double x, double x2)
            {
                if (Chek == true)
                {
                    if (con - (pen[0] + x) < x2)
                        per = true;
                    else
                        per = false;
                    TransprtPen(pen, x, per);
                    swModel.SketchManager.CreateCornerRectangle(pen[0], pen[1], 0, pen[0] + x2, pen[1] + wid, 0);
                    namberLine = namberLine + 4;
                    numberLuc++;

                }
            }
            
            // построение отверстии для интедефикации лючка

            void OtvLuc(bool chek, double x, double y,int num)
            {
                if (chek==true)
                for (int i = 0; i < num; i++)
                {
                    swModel.SketchManager.CreateCircleByRadius(x + 0.01+(i*0.01), y + wid - 0.005, 0, 0.0025);
                }
                                   
            }

            // Построение выреза под трубу

            void BildСutout(double[] pin, string cutout, double y, double yHeigth, double x, double length)
            {
                switch (cutout)
                {
                    case "Слева":
                        swModel.SketchManager.CreateCornerRectangle(pin[0], pin[1]+y , 0, pin[0] + x, pin[1]+y+yHeigth, 0);
                        swModel.Extension.SelectByID2(line+namberLine, "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        swModel.SketchManager.SketchTrim(0, pen[0], pen[1] + y + 0.001, 0);
                        namberLine=namberLine + 4;
                        swModel.Extension.SelectByID2(line+namberLine, "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        swModel.SketchManager.SketchTrim(0, pen[0], pen[1] + y + 0.001, 0);
                        namberLine = namberLine + 1;
                        break;
                    case "Справа":
                        swModel.SketchManager.CreateCornerRectangle(pen[0]+length, pen[1]+y, 0, pen[0]+length-x, pen[1] + y+yHeigth, 0);
                        namberLine = namberLine + 2;
                        swModel.Extension.SelectByID2(line + namberLine, "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        swModel.SketchManager.SketchTrim(0, pen[0]+length, pen[1] + y + 0.001, 0);
                        namberLine = namberLine + 2;
                        swModel.Extension.SelectByID2(line + namberLine, "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        swModel.SketchManager.SketchTrim(0, pen[0]+length, pen[1] + y + 0.001, 0);
                        namberLine = namberLine + 1;
                        break;
                    case "Нет":
                        break;
                }
                
                }

            // Постороение первого лючка

            swModel.SketchManager.CreateCornerRectangle(0, 0, 0, len1, wid, 0);
            swModel.SketchManager.CreateCircleByRadius(pen[0] + 0.01, pen[1]+wid-0.005, 0, 0.0025);
            BildСutout(pen, cutoutChek1.Text, heigthToCut1, heigthCut1, depthCut1,len1);
            BuildHole(pen, holeBox1.Text, hole1, slotLen1);

            //Построение 2-8 лючка
            
            BildLuc(l2.Checked, len1, len2);
            BildСutout(pen, cutoutChek2.Text, heigthToCut2, heigthCut2, depthCut2,len2);
            BuildHole(pen, holeBox2.Text, hole2, slotLen2);
            OtvLuc(l2.Checked, pen[0], pen[1], numberLuc);

            BildLuc(l3.Checked, len2, len3);
            BildСutout(pen, cutoutChek3.Text, heigthToCut3, heigthCut3, depthCut3,len3);
            BuildHole(pen, holeBox3.Text, hole3, slotLen3);
            OtvLuc(l3.Checked, pen[0], pen[1], numberLuc);

            BildLuc(l4.Checked, len3, len4);
            BildСutout(pen, cutoutChek4.Text, heigthToCut4, heigthCut4, depthCut4,len4);
            BuildHole(pen, holeBox4.Text, hole4, slotLen4);
            OtvLuc(l4.Checked, pen[0], pen[1], numberLuc);

            BildLuc(l5.Checked, len4, len5);
            BildСutout(pen, cutoutChek5.Text, heigthToCut5, heigthCut5, depthCut5,len5);
            BuildHole(pen, holeBox5.Text, hole5, slotLen5);
            OtvLuc(l5.Checked, pen[0], pen[1], numberLuc);

            BildLuc(l6.Checked, len5, len6);
            BildСutout(pen, cutoutChek6.Text, heigthToCut6, heigthCut6, depthCut6,len6);
            BuildHole(pen, holeBox6.Text, hole6, slotLen6);
            OtvLuc(l6.Checked, pen[0], pen[1], numberLuc);

            BildLuc(l7.Checked, len6, len7);
            BildСutout(pen, cutoutChek7.Text, heigthToCut7, heigthCut7, depthCut7,len7);
            BuildHole(pen, holeBox7.Text, hole7, slotLen7);
            OtvLuc(l7.Checked, pen[0], pen[1], numberLuc);

            BildLuc(l8.Checked, len7, len8);
            BildСutout(pen, cutoutChek8.Text, heigthToCut8, heigthCut8, depthCut8,len8);
            BuildHole(pen, holeBox8.Text, hole8, slotLen8);
            OtvLuc(l8.Checked, pen[0], pen[1], numberLuc);

            // Сохранение файла


            PartDoc swPart;
            swPart = (PartDoc)swModel;

            // переменые для сохранения 

            double[] dataAlignment = new double[12];
            object varAlignment;
            string[] dataViews = new string[2];
            object varViews;
            string nameF = "C:\\Luk\\№" + nameFile.Text + ".SLDPRT";

            // Сохранение 3D

            swModel.SaveAs3(nameF, 0, 1);

            string sModelName = swModel.GetPathName();
            string sPathName = swModel.GetPathName();
            sPathName = sPathName.Substring(0, sPathName.Length - 6);
            sPathName = sPathName + "dxf";

            
            dataAlignment[0] = 0.0;
            dataAlignment[1] = 0.0;
            dataAlignment[2] = 0.0;
            dataAlignment[3] = 1.0;
            dataAlignment[4] = 0.0;
            dataAlignment[5] = 0.0;
            dataAlignment[6] = 0.0;
            dataAlignment[7] = 1.0;
            dataAlignment[8] = 0.0;
            dataAlignment[9] = 0.0;
            dataAlignment[10] = 0.0;
            dataAlignment[11] = 1.0;

            varAlignment = dataAlignment;

            dataViews[0] = "*Текущий";
            dataViews[1] = "*Спереди";
            varViews = dataViews;

           

            swModel.SketchManager.InsertSketch(true);
            swModel.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, 0);

            // Cохранение DXF

            swPart.ExportToDWG2(sPathName, sModelName, 3, true, varAlignment, false, false, 0, dataViews);

            //123123232323455
             
        }

        
    }   

}

