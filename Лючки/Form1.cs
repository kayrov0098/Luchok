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
        ModelDoc2 swModel;
        DrawingDoc swDrawing;
        SelectionMgr swSelmgr;
        SolidWorks.Interop.sldworks.View swView;

        //  ModelDoc2 swModel;

        private void button1_Click1(object sender, EventArgs e)
        { }

        private void button1_Click(object sender, EventArgs e)
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

            



            double wid = (double)width.Value / d;
            var pen = new double[2];
            bool per = false;

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
                }
            }

            // Постороение первого лючка

            swModel.SketchManager.CreateCornerRectangle(0, 0, 0, len1, wid, 0);
            BuildHole(pen, holeBox1.Text, hole1, slotLen1);

            //Построение 2-8 лючка

            BildLuc(l2.Checked, len1, len2);
            BuildHole(pen, holeBox2.Text, hole2, slotLen2);
            BildLuc(l3.Checked, len2, len3);
            BuildHole(pen, holeBox3.Text, hole3, slotLen3);
            BildLuc(l4.Checked, len3, len4);
            BuildHole(pen, holeBox4.Text, hole4, slotLen4);
            BildLuc(l5.Checked, len4, len5);
            BuildHole(pen, holeBox5.Text, hole5, slotLen5);
            BildLuc(l6.Checked, len5, len6);
            BuildHole(pen, holeBox6.Text, hole6, slotLen6);
            BildLuc(l7.Checked, len6, len7);
            BuildHole(pen, holeBox7.Text, hole7, slotLen7);
            BildLuc(l8.Checked, len7, len8);
            BuildHole(pen, holeBox8.Text, hole8, slotLen8);

            // Сохранение файла


            PartDoc swPart;
            swPart = (PartDoc)swModel;

            // переменые для сохранения 

            double[] dataAlignment = new double[12];
            object varAlignment;
            string[] dataViews = new string[2];
            object varViews;
            string nameF = "C:\\Users\\kayrov\\Desktop\\Лючки\\Лючек №" + nameFile.Text + ".SLDPRT";

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

