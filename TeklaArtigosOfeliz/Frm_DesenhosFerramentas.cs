using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using TeklaMacroBuilder;
using Point = Tekla.Structures.Geometry3d.Point;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;


namespace TeklaArtigosOfeliz
{
    public partial class Frm_DesenhosFerramentas : Form
    {
        Frm_Inico _formpai;

        private string filePath;
        private Timer checkTimer;

        public Frm_DesenhosFerramentas(Frm_Inico formpai)
        {
            InitializeComponent();
            _formpai = formpai;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button17_Click(object sender, EventArgs e)
        {



            //string V1 = "Desligado";
            //string V2 = "Desligado";
            //string V3 = "Desligado";
            //string V4 = "Desligado";

            string V5 = "0";
            string V6 = "0";
            string V7 = "0";
            string V8 = "0";


            if (checkBox2.Checked)
            {
                //V1 = "Sim";
                V5 = "2";
            }
            if (checkBox3.Checked)
            {
                //V2 = "Sim";
                V6 = "2";
            }
            if (checkBox4.Checked)
            {
                //V3 = "Sim";
                V7 = "2";
            }
            if (checkBox5.Checked)
            {
                //V4 = "Sim";
                V8 = "2";
            }


            ///////////////////old dialog box/////////////////////////////////////
            DrawingHandler dh = new DrawingHandler();
            Drawing MyCurrentDrawing = dh.GetActiveDrawing();


            if (MyCurrentDrawing.GetType().ToString().Contains("SinglePartDrawing"))
            {
                new TeklaMacroBuilder.MacroBuilder()
                .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                .PushButton("gr_wdraw_view", "wdraw_dial")
                .ValueChange("wdv_dial", "gr_dv_cre_front_sw", V5)
                .ValueChange("wdv_dial", "gr_dv_cre_top_sw", V6)
                .ValueChange("wdv_dial", "gr_dv_cre_back_sw", V8)
                .ValueChange("wdv_dial", "gr_dv_cre_bottom_sw", V7)
                .PushButton("dv_modify", "wdv_dial")
                .PushButton("dv_ok", "wdv_dial")
                .PushButton("gr_wdraw_ok", "wdraw_dial")
                .Run();

                ArrayList MyViews = new ArrayList();

                foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
                {
                    if (drawingObject is Tekla.Structures.Drawing.View)
                    {

                        Tekla.Structures.Drawing.View drawingObj = (TSD.View)drawingObject;
                        if (drawingObj.ViewType != TSD.View.ViewTypes._3DView)
                        {
                            MyViews.Add(drawingObject);
                        }
                    }
                }

                if (MyViews.Count > 0)
                {
                    dh.GetDrawingObjectSelector().SelectObjects(MyViews, false);
                }

            }
            else
            {
                new TeklaMacroBuilder.MacroBuilder()
            .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
            .PushButton("gr_adraw_view", "adraw_dial")
            .ValueChange("adv_dial", "gr_dv_cre_top_sw", V6)
            .ValueChange("adv_dial", "gr_dv_cre_front_sw", V5)
            .ValueChange("adv_dial", "gr_dv_cre_back_sw", V8)
            .ValueChange("adv_dial", "gr_dv_cre_bottom_sw", V7)
            .PushButton("dv_modify", "adv_dial")
            .PushButton("dv_ok", "adv_dial")
            .Run();
            }



            ///////////////////new dialog box/////////////////////////////////////
            //new TeklaMacroBuilder.MacroBuilder()
            //    .Callback("acmd_display_attr_dialog", "wdraw_dial", "main_frame")
            //               .TableSelect("wdraw_dial", "table_ViewsTable", new int[] { 1 })
            //               .TableValueChange("wdraw_dial", "table_ViewsTable", "col_OnOff", V1)
            //               .TableSelect("wdraw_dial", "table_ViewsTable", new int[] { 2 })
            //               .TableValueChange("wdraw_dial", "table_ViewsTable", "col_OnOff", V2)
            //               .TableSelect("wdraw_dial", "table_ViewsTable", new int[] { 3 })
            //               .TableValueChange("wdraw_dial", "table_ViewsTable", "col_OnOff", V3)
            //               .TableSelect("wdraw_dial", "table_ViewsTable", new int[] { 4 })
            //               .TableValueChange("wdraw_dial", "table_ViewsTable", "col_OnOff", V4)
            //               .Run();


        }

        private void button18_Click(object sender, EventArgs e)
        {
            DrawingHandler dh = new DrawingHandler();
            var text = string.Empty;
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_USE_OLD_DRAWING_CREATION_SETTINGS", ref text);

            if (dh.GetActiveDrawing().GetType().ToString().Contains("SinglePartDrawing"))
            {
                if (text.ToLower().Contains("true"))
                {
                    new TeklaMacroBuilder.MacroBuilder()
                    .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                    .Callback("acmd_place_drawing_views", "", "main_frame")
                    .PushButton("gr_wdraw_layout", "wdraw_dial")
                    .ValueChange("wdl_dial", "gr_size_searching_mode", "0")
                    .PushButton("dl_modify", "wdl_dial")
                    .PushButton("dl_ok", "wdl_dial")
                    .PushButton("gr_wdraw_ok", "wdraw_dial")
                    .Run();
                }
                else
                {
                    new TeklaMacroBuilder.MacroBuilder()
                   .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                   .ValueChange("wdraw_dial", "gr_size_searching_mode", "0")
                   .PushButton("gr_wdraw_modify", "wdraw_dial")
                   .PushButton("gr_wdraw_ok", "wdraw_dial")
                   .Run();


                }

            }
            else
            {
                if (text.ToLower().Contains("true"))
                {
                    new TeklaMacroBuilder.MacroBuilder()
                   .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                   .Callback("acmd_place_drawing_views", "", "main_frame")
                   .PushButton("gr_adraw_layout", "adraw_dial")
                   .ValueChange("adl_dial", "gr_size_searching_mode", "0")
                   .PushButton("dl_modify", "adl_dial")
                   .PushButton("dl_ok", "adl_dial")
                   .PushButton("gr_adraw_ok", "adraw_dial")
                   .Run();
                }
                else
                {
                    new TeklaMacroBuilder.MacroBuilder()
                  .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                  .ValueChange("adraw_dial", "gr_size_searching_mode", "0")
                  .PushButton("gr_adraw_modify", "wdraw_dial")
                  .PushButton("gr_adraw_ok", "wdraw_dial")
                  .Run();
                }
            }
        }

        private void button59_Click(object sender, EventArgs e)
        {
            DrawingHandler dh = new DrawingHandler();
            Drawing MyCurrentDrawing = dh.GetActiveDrawing();



            ArrayList MyViews = new ArrayList();

            foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
            {
                if (drawingObject is Tekla.Structures.Drawing.View)
                {

                    Tekla.Structures.Drawing.View drawingObj = (TSD.View)drawingObject;
                    if (drawingObj.ViewType != TSD.View.ViewTypes._3DView)
                    {
                        MyViews.Add(drawingObject);
                    }
                }
            }

            if (MyViews.Count > 0)
            {
                dh.GetDrawingObjectSelector().SelectObjects(MyViews, false);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder().Callback("acmd_place_drawing_views", "", "main_frame").Run();
        }

        private void button39_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder().Callback("grOpenPreviousDrawingCB", "", "main_frame")
                                                .PopupCallback("grReadyForIssueOnMcb", "", "Drawing_selection", "dia_draw_select_list")
                                                .PushButton("gr_close_save_dr_editor_yes", "gr_close_dr_editor_confirm_instance")
                                                .Run();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder().Callback("grOpenNextDrawingCB", "", "main_frame")
                                  .PopupCallback("grReadyForIssueOnMcb", "", "Drawing_selection", "dia_draw_select_list")
                                  .PushButton("gr_close_save_dr_editor_yes", "gr_close_dr_editor_confirm_instance")
                                  .Run();
        }

        private void button55_Click(object sender, EventArgs e)
        {
            try
            {
                PointList pointList = new PointList();
                ViewBase view;
                StringList stringList = new StringList();

                DrawingHandler drawingHandler = new DrawingHandler();
                Drawing drawing = drawingHandler.GetActiveDrawing();

                Tekla.Structures.Drawing.UI.Picker picker = drawingHandler.GetPicker();

                stringList.Add("SELECIONAR PONTOS DE CRIAÇAO DAS COTAS");

                //Pick the point from the drawing
                picker.PickPoints(stringList, out pointList, out view);

                //Last picked point (middle button) defines the distance where the dimansion is placed
                Tekla.Structures.Geometry3d.Point pointToSetDimension = pointList[pointList.Count - 1];
                pointList.RemoveAt(pointList.Count - 1);

                //Calculate the distance where the dimension will be set
                double distance = Distance.PointToPoint(pointToSetDimension, pointList[pointList.Count - 1]);

                // Calculate the vector for the dimension orientation, this is going to be parallel 
                // to the line defined between the first and last picked points, and the distance is
                // set by the middle button.
                Tekla.Structures.Geometry3d.Line line = new Tekla.Structures.Geometry3d.Line(pointList[pointList.Count - 1], pointList[0]);
                Tekla.Structures.Geometry3d.Point projectedPoint = Projection.PointToLine(pointToSetDimension, line);

                Vector vector = new Vector(pointToSetDimension - projectedPoint);

                // Define the dimension
                StraightDimensionSetHandler dimensionSetHandler = new StraightDimensionSetHandler();
                StraightDimensionSet.StraightDimensionSetAttributes attr = new StraightDimensionSet.StraightDimensionSetAttributes();
                StraightDimensionSet dimensionSet = dimensionSetHandler.CreateDimensionSet(view, pointList, vector, distance, attr);

                //Insert the dimension
                //dimensionSet.Insert();

                //Commit the changes to the drawing, to update database and view.
                // drawing.CommitChanges();
            }
            catch (System.Exception)
            {


            }
        }

        private void button64_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder().PushButton("dia_draw_lock_on", "Drawing_selection")
                       .PushButton("dia_draw_freeze_on", "Drawing_selection")
                       .PushButton("dia_draw_ready_for_issue_on", "Drawing_selection")
                       .PushButton("dia_draw_issue_on", "Drawing_selection").Run();

            //DrawingHandler dh = new DrawingHandler();
            //DrawingEnumerator de = dh.GetDrawingSelector().GetSelected();

            //while (de.MoveNext())
            //{
            //    Drawing currentDrawing = de.Current;

            //    string currentLockedBy = string.Empty;
            //    currentDrawing.GetUserProperty("DrawingLockedByAcc", ref currentLockedBy);

            //    string newLockedBy = "Spiderman";
            //    currentDrawing.SetUserProperty("DrawingLockedByAcc", newLockedBy);
            //    currentDrawing.IsFrozen = true;
            //    currentDrawing.IsLocked = true;
            //    currentDrawing.IsReadyForIssue = true;
            //    currentDrawing.Modify();
            //}




        }

        private void button65_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder().PushButton("dia_draw_lock_off", "Drawing_selection")
                                 .PushButton("dia_draw_freeze_off", "Drawing_selection")
                                 .PushButton("dia_draw_ready_for_issue_off", "Drawing_selection")
                                 .PushButton("dia_draw_issue_off", "Drawing_selection").Run();
        }

        private void button62_Click(object sender, EventArgs e)
        {
          new TeklaMacroBuilder.MacroBuilder().Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                                              .PushButton("gr_adraw_layout", "adraw_dial")
                                              .ValueChange("adl_dial", "gr_wdl_get_menu", "OFELIZ_A1")
                                              .PushButton("gr_wdl_get", "adl_dial")
                                              .PushButton("dl_modify", "adl_dial")
                                              .PushButton("gr_adraw_ok", "adraw_dial")
                                              .Run();
        }

        private void button63_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder().Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                                                .PushButton("gr_adraw_layout", "adraw_dial")
                                                .ValueChange("adl_dial", "gr_wdl_get_menu", "OFELIZ_A2")
                                                .PushButton("gr_wdl_get", "adl_dial")
                                                .PushButton("dl_modify", "adl_dial")
                                                .PushButton("gr_adraw_ok", "adraw_dial")
                                                .Run();
        }

        private void button60_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder().Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                                                .PushButton("gr_adraw_layout", "adraw_dial")
                                                .ValueChange("adl_dial", "gr_wdl_get_menu", "OFELIZ_A3")
                                                .PushButton("gr_wdl_get", "adl_dial")
                                                .PushButton("dl_modify", "adl_dial")
                                                .PushButton("gr_adraw_ok", "adraw_dial")
                                                .Run();
        }

        private void button61_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder().Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                                               .PushButton("gr_adraw_layout", "adraw_dial")
                                               .ValueChange("adl_dial", "gr_wdl_get_menu", "OFELIZ_A4")
                                               .PushButton("gr_wdl_get", "adl_dial")
                                               .PushButton("dl_modify", "adl_dial")
                                               .PushButton("gr_adraw_ok", "adraw_dial")
                                               .Run();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            ArrayList MyViews = new ArrayList();
            DrawingHandler MyDrawingHandler = new DrawingHandler();
            Drawing MyCurrentDrawing = MyDrawingHandler.GetActiveDrawing();

            foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
            {
                if (drawingObject is Tekla.Structures.Drawing.View)
                {
                    MyViews.Add(drawingObject);
                }
            }

            if (MyViews.Count > 0)
            {
                MyDrawingHandler.GetDrawingObjectSelector().SelectObjects(MyViews, false);
            }
            new TeklaMacroBuilder.MacroBuilder()
                   .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                   .PushButton("view_on_off", "view_dial")
                   .ValueChange("view_dial", "gr_view_scale", "5.000000000000")
                   .PushButton("view_modify", "view_dial")
                   .PushButton("view_ok", "view_dial")
                   .Run();
        }

        private void button26_Click(object sender, EventArgs e)
        {
            ArrayList MyViews = new ArrayList();
            DrawingHandler MyDrawingHandler = new DrawingHandler();
            Drawing MyCurrentDrawing = MyDrawingHandler.GetActiveDrawing();

            foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
            {
                if (drawingObject is Tekla.Structures.Drawing.View)
                {
                    MyViews.Add(drawingObject);
                }
            }

            if (MyViews.Count > 0)
            {
                MyDrawingHandler.GetDrawingObjectSelector().SelectObjects(MyViews, false);
            }
            new TeklaMacroBuilder.MacroBuilder()
                   .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                   .PushButton("view_on_off", "view_dial")
                   .ValueChange("view_dial", "gr_view_scale", "10.00000000000")
                   .PushButton("view_modify", "view_dial")
                   .PushButton("view_ok", "view_dial")
                   .Run();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            ArrayList MyViews = new ArrayList();
            DrawingHandler MyDrawingHandler = new DrawingHandler();
            Drawing MyCurrentDrawing = MyDrawingHandler.GetActiveDrawing();

            foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
            {
                if (drawingObject is Tekla.Structures.Drawing.View)
                {
                    MyViews.Add(drawingObject);
                }
            }

            if (MyViews.Count > 0)
            {
                MyDrawingHandler.GetDrawingObjectSelector().SelectObjects(MyViews, false);
            }
            new TeklaMacroBuilder.MacroBuilder()
                   .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                   .PushButton("view_on_off", "view_dial")
                   .ValueChange("view_dial", "gr_view_scale", "15.000000000000")
                   .PushButton("view_modify", "view_dial")
                   .PushButton("view_ok", "view_dial")
                   .Run();
        }

        private void button28_Click(object sender, EventArgs e)
        {
            ArrayList MyViews = new ArrayList();
            DrawingHandler MyDrawingHandler = new DrawingHandler();
            Drawing MyCurrentDrawing = MyDrawingHandler.GetActiveDrawing();

            foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
            {
                if (drawingObject is Tekla.Structures.Drawing.View)
                {
                    MyViews.Add(drawingObject);
                }
            }

            if (MyViews.Count > 0)
            {
                MyDrawingHandler.GetDrawingObjectSelector().SelectObjects(MyViews, false);
            }
            new TeklaMacroBuilder.MacroBuilder()
                   .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                   .PushButton("view_on_off", "view_dial")
                   .ValueChange("view_dial", "gr_view_scale", "20.000000000000")
                   .PushButton("view_modify", "view_dial")
                   .PushButton("view_ok", "view_dial")
                   .Run();
        }

        private void button29_Click(object sender, EventArgs e)
        {
            ArrayList MyViews = new ArrayList();
            DrawingHandler MyDrawingHandler = new DrawingHandler();
            Drawing MyCurrentDrawing = MyDrawingHandler.GetActiveDrawing();

            foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
            {
                if (drawingObject is Tekla.Structures.Drawing.View)
                {

                    Tekla.Structures.Drawing.View drawingObj = (TSD.View)drawingObject;
                    if (drawingObj.ViewType == TSD.View.ViewTypes._3DView)
                    {
                        MyViews.Add(drawingObject);
                    }
                }
            }

            if (MyViews.Count > 0)
            {
                MyDrawingHandler.GetDrawingObjectSelector().SelectObjects(MyViews, false);
            }
            new TeklaMacroBuilder.MacroBuilder()
              .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
              .PushButton("view_on_off", "view_dial")
              .ValueChange("view_dial", "gr_view_scale", textBox3.Text)
              .PushButton("view_modify", "view_dial")
              .PushButton("view_ok", "view_dial")
              .Run();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            ArrayList MyViews = new ArrayList();
            DrawingHandler MyDrawingHandler = new DrawingHandler();
            Drawing MyCurrentDrawing = MyDrawingHandler.GetActiveDrawing();

            foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
            {
                if (drawingObject is Tekla.Structures.Drawing.View)
                {

                    Tekla.Structures.Drawing.View drawingObj = (TSD.View)drawingObject;
                    if (drawingObj.ViewType != TSD.View.ViewTypes._3DView)
                    {
                        MyViews.Add(drawingObject);
                    }
                }
            }

            if (MyViews.Count > 0)
            {
                MyDrawingHandler.GetDrawingObjectSelector().SelectObjects(MyViews, false);
            }
            if (textBox4.Text != "" && textBox5.Text != "")
            {
                new TeklaMacroBuilder.MacroBuilder()
                   .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                   .TreeSelect("view_dial", "gratCastUnitDrawingAttributesMenuTree", "Attributes")
                   .PushButton("view_on_off", "view_dial")
                   //.ValueChange("view_dial", "gr_view_size_mode", "1")
                   .ValueChange("view_dial", "gr_view_scale", textBox4.Text)
                   .ValueChange("view_dial", "gr_view_cut_min_dist", textBox5.Text)
                   .PushButton("view_modify", "view_dial")
                   .PushButton("view_ok", "view_dial")
                   .PushButton("gr_adraw_ok", "adraw_dial")
                   .Run();
            }
            else if (textBox4.Text != "")
            {
                new TeklaMacroBuilder.MacroBuilder()
                   .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                   .TreeSelect("view_dial", "gratCastUnitDrawingAttributesMenuTree", "Attributes")
                   .PushButton("view_on_off", "view_dial")
                   //.ValueChange("view_dial", "gr_view_size_mode", "1")
                   .ValueChange("view_dial", "gr_view_scale", textBox4.Text)
                   .PushButton("view_modify", "view_dial")
                   .PushButton("view_ok", "view_dial")
                   .PushButton("gr_adraw_ok", "adraw_dial")
                   .Run();
            }
            else if (textBox5.Text != "")
            {
                new TeklaMacroBuilder.MacroBuilder()
                               .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                               .TreeSelect("view_dial", "gratCastUnitDrawingAttributesMenuTree", "Attributes")
                               .PushButton("view_on_off", "view_dial")
                               //.ValueChange("view_dial", "gr_view_size_mode", "1")
                               .ValueChange("view_dial", "gr_view_cut_min_dist", textBox5.Text)
                               .PushButton("view_modify", "view_dial")
                               .PushButton("view_ok", "view_dial")
                               .PushButton("gr_adraw_ok", "adraw_dial")
                               .PushButton("gr_adraw_ok", "adraw_dial")
                               .Run();
            }
        }

        private void button42_Click(object sender, EventArgs e)
        {
            DrawingHandler dh = new DrawingHandler();
            var text = string.Empty;
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_USE_OLD_DRAWING_CREATION_SETTINGS", ref text);

            if (dh.GetActiveDrawing().GetType().ToString().Contains("SinglePartDrawing"))
            {

                if (text.ToLower().Contains("true"))
                {

                    new TeklaMacroBuilder.MacroBuilder()
                .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                .PushButton("gr_adraw_view", "adraw_dial")
                .TabChange("adv_dial", "contMain", "tabAttributes")
                .ValueChange("adv_dial", "gr_dv_coord_sys_z_rotate", textBox8.Text)
                .ValueChange("adv_dial", "gr_dv_coord_sys_x_rotate", textBox7.Text)
                .PushButton("dv_modify", "adv_dial")
                .PushButton("dv_ok", "adv_dial")
                .PushButton("gr_adraw_ok", "adraw_dial")
                .PushButton("gr_adraw_ok", "adraw_dial")
                .Run();

                }
                else
                {

                    new TeklaMacroBuilder.MacroBuilder()
                    .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                    .TabChange("wdraw_dial", "contMain", "tabAttributes")
                    .ValueChange("wdraw_dial", "gr_dv_coord_sys_x_rotate", textBox7.Text)
                    .ValueChange("wdraw_dial", "gr_dv_coord_sys_z_rotate", textBox8.Text)
                    .ModalDialog(1)
                    .PushButton("gr_wdraw_modify", "wdraw_dial")
                    .PushButton("gr_wdraw_ok", "wdraw_dial")
                    .PushButton("gr_adraw_ok", "adraw_dial")
                    .Run();

                }
            }
            else
            {
                if (text.ToLower().Contains("true"))
                {
                    new TeklaMacroBuilder.MacroBuilder()
                    .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                    .PushButton("gr_adraw_view", "adraw_dial")
                    .TabChange("adv_dial", "contMain", "tabAttributes")
                    .ValueChange("adv_dial", "gr_dv_coord_sys_z_rotate", textBox8.Text)
                    .ValueChange("adv_dial", "gr_dv_coord_sys_x_rotate", textBox7.Text)
                    .PushButton("dv_modify", "adv_dial")
                    .PushButton("dv_ok", "adv_dial")
                    .PushButton("gr_adraw_ok", "adraw_dial")
                    .Run();
                }
                else
                {

                    new TeklaMacroBuilder.MacroBuilder()
                  .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                  .TabChange("adraw_dial", "contMain", "tabAttributes")
                  .ValueChange("adraw_dial", "gr_dv_coord_sys_x_rotate", textBox7.Text)
                  .ValueChange("adraw_dial", "gr_dv_coord_sys_z_rotate", textBox8.Text)
                  .ModalDialog(1)
                  .PushButton("gr_adraw_modify", "adraw_dial")
                  .PushButton("gr_adraw_ok", "adraw_dial").PushButton("gr_adraw_ok", "adraw_dial").Run();


                }
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder()
                .Callback("acmd_display_selected_drawing_object_dialog", "", "View_10 window_1")
                .ValueChange("wdraw_dial", "gr_dv_recreate_drawing", "1")
                .ModalDialog(1)
                .PushButton("gr_wdraw_modify", "wdraw_dial")
                .PushButton("gr_wdraw_ok", "wdraw_dial").PushButton("gr_adraw_ok", "adraw_dial").Run();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LBLestado.Text = "A verificar conjuntos selecionados";
            ArrayList c = ComunicaTekla.ListadePecasdoConjSelec();

            LBLestado.Text = "A criar desenhos p.f. aguarde";
            ComunicaTekla a = new ComunicaTekla();
            ComunicaTekla.selectinmodel(c);
            string fasex = null;
            List<string> l = new List<string>();
            foreach (TSM.Part item in c)
            {
                item.GetReportProperty("Fase", ref fasex);

                if (!l.Contains(fasex))
                {
                    l.Add(fasex);
                }
            }
            l.Distinct();
            string outfase = null;
            foreach (var item in l)
            {
                if (item.ToString().Trim() != "0")
                {
                    outfase += "<p>Fase " + item + "</p>";
                }

            }

            new TeklaMacroBuilder.MacroBuilder().Callback("acmd_partnumbers_selected", "", "main_frame").Run();
            bool b = a.CriaDesenhos(c);
            if (b == true)
            {
                /////////////////////////////////////////////////////////////////////////////////////////


                /////////////////////////////////////////////////////////////////////////////////////////

                LBLestado.Text = "Desenhos criados";
                MessageBox.Show(this, "Desenhos criados com sucesso", "Criação de Desenhos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LBLestado.Text = "Desenhos criados";

            }
            else
            {
                LBLestado.Text = "Erro na criação de desenhos";
                MessageBox.Show(this, "Possivel erro, o método de seleção." + Environment.NewLine + "P.F. altere o método para Selecionar conjuntos", "Criação de Desenhos", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {

            if (label11.Text == new Model().GetProjectInfo().ProjectNumber)
            {

                if (Directory.GetFiles(@"C:\R\").Length > 0)
                {
                    DialogResult DIAL = MessageBox.Show(this, @"Existe ficheiros na pasta C:\R\ deseja continuar?" + Environment.NewLine + "Se Responder sim o programa ira limpar a pasta e prosseguir", "ALERTA", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (DIAL == DialogResult.Yes)
                    {
                        foreach (string file in Directory.GetFiles(@"C:\R\"))
                        {
                            File.Delete(file);
                        }
                    }
                }
                if (Directory.GetFiles(@"C:\R\").Length == 0)
                {
                    LBLestado.Text = "A selecionar Conjuntos";
                    ArrayList peças = new ArrayList(ComunicaTekla.ListadePecasdoConjSelec());
                    ArrayList conjuntos = new ArrayList(ComunicaTekla.ListadeConjuntosSelec());
                    ArrayList objectos = new ArrayList(peças);
                    //////////////////////////////////////////////////////////soldaura//////////////////////////////////////////////////////////////////////////////////

                    //PROCURA POR DESENHOS DE SOLDADURA 
                    List<string> tudo = new List<string>();
                    foreach (TSM.Assembly ASS in conjuntos)
                    {
                        if (ASS != null)
                        {
                            string CONJ = null;
                            string soldadura = null;
                            ASS.GetReportProperty("ASSEMBLY_POS", ref CONJ);
                            ASS.GetReportProperty("Operacoes_Conj", ref soldadura);
                            if (soldadura == "Opção 2" || soldadura == "Opção 5" || soldadura == "Opção 6" || soldadura == "Opção 16")
                            {
                                tudo.Add(CONJ);
                            }
                        }
                    }
                    IEnumerable dis = tudo.Distinct();

                    // FIM DA PROCURA
                   
                    foreach (string item in dis)
                    {
                        var result = Regex.Split(item, @"\d+$")[0] + "." + Regex.Match(item, @"\d+$").Value;
                        LBLestado.Text = "A abrir desenho"+ result + " - 1";

                                  new TeklaMacroBuilder.MacroBuilder()
                          .ValueChange("Drawing_selection", "diaSearchInOptionMenu", "7")
                          .ValueChange("Drawing_selection", "diaDrawingListSearchCriteria", result)
                          .PushButton("diaDrawingListSearch", "Drawing_selection")
                          .TableSelect("Drawing_selection", "dia_draw_select_list", new int[] { 1 })
                          .PopupCallback("acmd_copy_drawing_to_new_sheet", "", "Drawing_selection", "dia_draw_select_list")
                          .ValueChange("Drawing_selection", "diaDrawingListSearchCriteria", result + " - 1")
                          .PushButton("diaDrawingListSearch", "Drawing_selection")
                          .TableSelect("Drawing_selection", "dia_draw_select_list", new int[] { 1 })
                          .Activate("Drawing_selection", "dia_draw_select_list")
                          .Run();
                        DrawingHandler drawingHandler = new DrawingHandler();
                        try
                        {

                            ContainerView sheet = drawingHandler.GetActiveDrawing().GetSheet();
                            if (drawingHandler.GetConnectionStatus())
                            {
                                LBLestado.Text = "A apagar cotas "+ result + " - 1";
                                System.Type[] Types = new System.Type[1];
                                Types.SetValue(typeof(StraightDimension), 0);

                                DrawingObjectEnumerator allDimLines = sheet.GetAllObjects(Types);
                                
                                
                                foreach (StraightDimension line in allDimLines)
                                {

                                    line.Delete();

                                }

                                Types.SetValue(typeof(AngleDimension), 0);

                                allDimLines = sheet.GetAllObjects(Types);

                                foreach (AngleDimension line in allDimLines)
                                {
                                    line.Delete();
                                }
                                Types.SetValue(typeof(RadiusDimension), 0);

                                allDimLines = sheet.GetAllObjects(Types);
                                foreach (RadiusDimension line in allDimLines)
                                {
                                    line.Delete();
                                }
                            }

                            new TeklaMacroBuilder.MacroBuilder()
                                     .ValueChange("gr_close_dr_editor_confirm_instance", "gr_close_save_dr_editor_freeze", "1")
                                     .PushButton("gr_close_save_dr_editor_yes", "gr_close_dr_editor_confirm_instance").Run();


                            DrawingHandler dh = new DrawingHandler();
                            ViewBase _sheet = dh.GetActiveDrawing().GetSheet();
                            Text text = new Text(_sheet, new Tekla.Structures.Geometry3d.Point(285, 70), "NOTA: Soldar segundo os nossos procedimentos habituais");
                            text.Attributes.LoadAttributes("soldaconfig");
                            text.Insert();
                            dh.SaveActiveDrawing();
                            dh.CloseActiveDrawing();
                        }
                        catch (System.Exception)
                        {
                            MessageBox.Show(this, "Erro por favor numerar modelo", "erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    /////////////////////////////////////////////////////////////////fim da soldadura//////////////////////////////////////////////////////////////////////////


                    LBLestado.Text = "A imprimir ...";
                    //ComunicaTekla.imprimepdf(conjuntos, peças);

                    ComunicaTekla.imprimepdf(conjuntos, peças, LBLestado);

                    LBLestado.Text = "A selecionar objectos";
                    foreach (TSM.Part item in peças)
                    {
                        foreach (BoltGroup parafuso in item.GetBolts())
                        {
                            if (parafuso != null)
                            {
                                if (parafuso.PartToBoltTo.Identifier.ID == item.Identifier.ID)
                                {
                                    objectos.Add(parafuso);
                                }
                            }
                        }
                    }
                    ComunicaTekla.selectinmodel(objectos);
                    if (!Directory.Exists(Frm_Inico.CaminhoModelo + @"\listas"))
                    {
                        Directory.CreateDirectory(Frm_Inico.CaminhoModelo + @"\listas");
                    }

                    LBLestado.Text = "A criar lista ";
                    TSM.Operations.Operation.CreateReportFromSelected("OFELIZ", @"C:\R\OFELIZ.CSV", "", "", "");
                    TSM.Operations.Operation.CreateReportFromSelected("OFELIZ.csv", @"C:\R\PEÇAS_E_CONJUNTOS.CSV", "", "", "");
                    LBLestado.Text = "A criar CNC ";
                    TSM.Operations.Operation.CreateNCFilesFromSelected("OFELIZ_chapas", @"c:\r\");
                    TSM.Operations.Operation.CreateNCFilesFromSelected("OFELIZ_perfis", @"c:\r\");
                    TSM.Operations.Operation.CreateNCFilesFromSelected("OFELIZ_madres", @"c:\r\");
                    LBLestado.Text = "A converter DXF ";

                    string[] NCfiles = Directory.GetFiles(@"c:\r", "*.nc1", SearchOption.TopDirectoryOnly);
                    List<string> myfiles = new List<string>();
                    foreach (var item in NCfiles)
                    {
                        myfiles.Add(item);
                    }
                    dstv_dxf.CRIAR(myfiles);
                    LBLestado.Text = "Convertidos os DXF's ";

                    timer1.Enabled = true;
                }
                else
                {
                    MessageBox.Show(this, "O projeto atual deste programa não é o projeto atual do Tekla.", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //// NÚMERO DE FICHEIROS NA IMPRESSORA 
            string printerName = "PDFCreator";
            PrintServer ps = new PrintServer();
            PrintQueue pq = ps.GetPrintQueue(printerName);
            if (pq.NumberOfJobs == 0)
            {
                this.Visible = false;
                timer1.Enabled = false;
                Frm_ListaOFeliz f = new Frm_ListaOFeliz(_formpai);
                LBLestado.Text = "Todos os dados criados com sucesso";
                f.ShowDialog();
                this.Visible = true;
            }
            else
            {
                LBLestado.Text = "AINDA FALTA IMPRIMIR " + pq.NumberOfJobs.ToString() + " DOCUMENTOS";
            }
        }

        private void FrmDesenhosFerramentas_Load(object sender, EventArgs e)
        {
            label11.Text = _formpai.label11.Text;
            label1.Text = "";
            foreach (var item in Frm_Inico.str)
            {
                label1.Text += item + " | ";
            }

            InitWatcher();
            this.FormClosed += Frm_DesenhosFerramentas_FormClosed;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            Model m = new Model();
            bool VERIFICACAO = false;
            if (Environment.UserName.ToLower() != "rui.ferreira")
            {
                DialogResult a = MessageBox.Show(this, "O método de seleção do tekla esta em modo de peça?", "ATENÇÃO", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (a == DialogResult.Yes)
                {
                    VERIFICACAO = true;
                }
                else
                {
                    MessageBox.Show(this, "ALTERE O MODO DE SELEÇÃO");
                }
            }
            else
            {
                   VERIFICACAO = true;
            }

            if (VERIFICACAO)
            {
                if (label11.Text == m.GetProjectInfo().ProjectNumber)
                {

                    LBLestado.Text = "A Retirar lista";
                    TSM.Operations.Operation.CreateReportFromSelected("OFELIZPARAFUSOOBRA", @"C:\R\OFELIZ.CSV", "", "", "");
                    this.Visible = false;
                    Frm_Parafusos p = new Frm_Parafusos(_formpai);
                    p.ShowDialog();
                    this.Visible = true;

                }
                else
                {
                    MessageBox.Show(this, "O projeto atual deste programa não é o projeto atual do Tekla.", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (label11.Text == new Model().GetProjectInfo().ProjectNumber)
            {

                ArrayList peças = new ArrayList(ComunicaTekla.ListadePecasdoConjSelec());
                ArrayList objectos = new ArrayList(peças);
                LBLestado.Text = "A selecionar objectos";
                foreach (TSM.Part item in peças)
                {
                    foreach (BoltGroup parafuso in item.GetBolts())
                    {
                        if (parafuso != null)
                        {
                            if (parafuso.PartToBoltTo.Identifier.ID == item.Identifier.ID)
                            {
                                objectos.Add(parafuso);
                            }
                        }


                    }
                }
                ComunicaTekla.selectinmodel(objectos);
                if (!Directory.Exists(Frm_Inico.CaminhoModelo + @"\listas"))
                {
                    Directory.CreateDirectory(Frm_Inico.CaminhoModelo + @"\listas");
                }
                TSM.Operations.Operation.CreateReportFromSelected("OFELIZ", @"C:\R\OFELIZ.CSV", "", "", "");
                this.Visible = false;
                Frm_ListaOFeliz f = new Frm_ListaOFeliz(_formpai);
                LBLestado.Text = "Lista criada com sucesso";
                f.ShowDialog(this);
                this.Visible = true;
            }
            else
            {
                MessageBox.Show(this, "O projeto atual deste programa não é o projeto atual do Tekla.", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Chb_alw_top_CheckedChanged(object sender, EventArgs e)
        {
            if (Chb_alw_top.Checked)
            {
                TopMost = true;
            }
            else
            {
                TopMost = false;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(label11.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {

            new MacroBuilder()
            .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
            .PushButton("gr_adraw_diming", "adraw_dial")
            .TabChange("adcd_dial", "Container_277", "Container_278")
            .ValueChange("adcd_dial", "adcd_main_part_extrema", "3")
            .PushButton("adcd_modify", "adcd_dial")
            .PushButton("adcd_apply", "adcd_dial")
            .PushButton("adcd_apply", "adcd_dial")
            .PushButton("adcd_ok", "adcd_dial")
            .PushButton("gr_adraw_view", "adraw_dial")
            .TabChange("adv_dial", "contMain", "tabAttributes")
            .TabChange("adv_dial", "contMain", "tabShortening")
            .PushButton("gr_adraw_layout", "adraw_dial")
            .TabChange("adl_dial", "Container_367", "tabOther")
            .ValueChange("adl_dial", "optMnuSectionViewInLineWithMain", "1")
            .ValueChange("adl_dial", "Optionmenu_1213", "1")
            .PushButton("dl_modify", "adl_dial")
            .PushButton("dl_apply", "adl_dial")
            .PushButton("dl_ok", "adl_dial")
            .PushButton("gr_adraw_apply", "adraw_dial")
            .PushButton("gr_adraw_ok", "adraw_dial")
            .Callback("acmd_place_drawing_views", "", "View_10 window_1")
            .Run();
            DrawingHandler dh = new DrawingHandler();
            Drawing MyCurrentDrawing = dh.GetActiveDrawing();
            foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
            {
                if (drawingObject is Tekla.Structures.Drawing.View)
                {

                    TSD.View drawingObj = (TSD.View)drawingObject;

                    if (drawingObj.ViewType == TSD.View.ViewTypes._3DView)
                    {
                        drawingObj.Origin = new Point(100, 20, 0) - drawingObj.FrameOrigin;
                        drawingObj.Attributes.FixedViewPlacing = true;
                        drawingObj.Modify();
                    }

                }
            }
            
            new TeklaMacroBuilder.MacroBuilder().Callback("acmd_place_drawing_views", "", "main_frame").Run();
        }
        

        public static void GetOpenDrawingInfo()
        {
            DrawingHandler drawingHandler = new DrawingHandler();

            if (drawingHandler.GetConnectionStatus())
            {
                Drawing currentDrawing = drawingHandler.GetActiveDrawing();

                if (currentDrawing != null)
                {
                    string drawingName = currentDrawing.Name;
                    string drawingMark = currentDrawing.Mark;                   

                    string drawnBy = Environment.UserName;
                    drawnBy = drawnBy.Replace('.', ' ');
                    drawnBy = string.Join(" ", drawnBy.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                    string mensagem = $"Desenho aberto: {drawingName}\n" +
                                      $"Marca: {drawingMark}\n" +
                                      $"Preparado por: {drawnBy}";

                }
                else
                {
                    MessageBox.Show("Nenhum desenho está aberto no momento.");
                }
            }
            else
            {
                MessageBox.Show("Não conectado ao ambiente de desenho.");
            }
        }



        private void InitWatcher()
        {
            Model modelo = new Model();
            string caminhoModelo = modelo.GetInfo().ModelPath;

            string nomeUsuario = Environment.UserName;
            filePath = Path.Combine(caminhoModelo, nomeUsuario + ".txt");

            checkTimer = new Timer();
            checkTimer.Interval = 5000;
            checkTimer.Tick += CheckDrawingAndUpdateFile;
            checkTimer.Start();

            CheckDrawingAndUpdateFile(null, null);
        }

        private void CheckDrawingAndUpdateFile(object sender, EventArgs e)
        {
            DrawingHandler drawingHandler = new DrawingHandler();
            if (!drawingHandler.GetConnectionStatus())
                return;

            Drawing currentDrawing = drawingHandler.GetActiveDrawing();
            if (currentDrawing == null)
                return;

            string drawingName = currentDrawing.Name;
            string drawingMark = currentDrawing.Mark;

            string nomeUsuario = Environment.UserName;
            nomeUsuario = nomeUsuario.Replace('.', ' ');
            nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

            string novaLinha = $"{nomeUsuario} -- {drawingMark} de/o {drawingName}";

            if (!File.Exists(filePath) || File.ReadAllText(filePath).Trim() != novaLinha.Trim())
            {
                File.WriteAllText(filePath, novaLinha);
            }

            MostrarTxtsDasSiglasNoWebBrowser();
        }

        private void Frm_DesenhosFerramentas_FormClosed(object sender, FormClosedEventArgs e)
        {
            checkTimer?.Stop();

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }


        private void MostrarTxtsDasSiglasNoWebBrowser()
        {
            List<string> siglas = new List<string>();
            Model modelo = new Model();
            string caminhoModelo = modelo.GetInfo().ModelPath;

            ComunicaBaseDados comunicaBD = new ComunicaBaseDados();

            try
            {
                comunicaBD.ConectarBDArtigo();

                string query = "SELECT [nome.sigla] FROM dbo.nPreparadores1";
                DataTable resultado = comunicaBD.ProcurarbdArtigo(query);

                foreach (DataRow row in resultado.Rows)
                {
                    string sigla = row[0].ToString().Trim();
                    if (!string.IsNullOrEmpty(sigla))
                    {
                        siglas.Add(sigla);
                    }                    
                }                 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao acessar banco de dados: " + ex.Message);
                return;
            }
            finally
            {
                comunicaBD.DesonectarBDArtigo();
            }

            //string html = "<html><head><style>body { font-family: Consolas; } pre { background: #eee; padding: 10px; border: 1px solid #ccc; }</style></head><body>";

            string html = @"
                            <html>
                            <head>
                            <style>
                                body {
                                    font-family: Consolas;
                                    background-color: rgb(240,240,240); /* Cor de fundo da página */
                                    margin: 20px;
                                }
                                pre {
                                    background-color: #ffffff; /* Cor de fundo das caixas de texto */
                                    padding: 10px;
                                    border: 1px solid #ccc;
                                }
                                h4 {
                                    color: #333;
                                }
                            </style>
                            </head>
                            <body>";

            bool encontrouArquivo = false;           

            foreach (string sigla in siglas)
            {
                string caminhoTxt = Path.Combine(caminhoModelo, sigla + ".txt");              

                if (File.Exists(caminhoTxt))
                {
                    string conteudo = File.ReadAllText(caminhoTxt)
                                      .Replace("[", "")
                                      .Replace("]", "");

                    string nomeUsuario = Environment.UserName;
                    nomeUsuario = nomeUsuario.Replace('.', ' ');
                    nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                    html += $"<pre><b>{System.Net.WebUtility.HtmlEncode(conteudo)}</b></pre>";
                    encontrouArquivo = true;
                }
            }

            html += "</body></html>";

            if (!encontrouArquivo)
            {
                html = "<html><body><p></p></body></html>";
            }

            webBrowser1.DocumentText = html;
        }

        }
    }


