using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using Tekla.Structures.Filtering;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;
using TSD = Tekla.Structures.Drawing;
using Tekla.Structures.Filtering.Categories;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_Soldadura : Form
    {
        public Frm_Soldadura()
        {
            InitializeComponent();
        }

        private void button41_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder()
           .Callback("acmdDisplayEnvironmentVariablesDialogFromMenu", "", "main_frame")
           .ValueChange("diaEnvironmentVariableDialog", "FindString", "​FIXED")
           .ValueChange("diaEnvironmentVariableDialog", "AllCategoriesCheckButton", "1")
           .TableSelect("diaEnvironmentVariableDialog", "tblEnvironmentVariables", new int[] { 5 })
           .TableValueChange("diaEnvironmentVariableDialog", "tblEnvironmentVariables", "colValue", "FALSE")
           .PushButton("butApply", "diaEnvironmentVariableDialog")
           .PushButton("butOk", "diaEnvironmentVariableDialog")
           .Run();

            TSM.Model MODELO = new Model();
            BinaryFilterExpression FILTRO = new BinaryFilterExpression(new TemplateFilterExpressions.CustomString("USERDEFINED.Fase"), StringOperatorType.IS_EQUAL, new StringConstantFilterExpression(textBox6.Text));
            ModelObjectEnumerator Objects = MODELO.GetModelObjectSelector().GetObjectsByFilter(FILTRO);
            List<string> tudo = new List<string>();

            while (Objects.MoveNext())
            {
                TSM.Assembly ASS = Objects.Current as TSM.Assembly;
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

            foreach (string item in dis)
            {
                var result = Regex.Split(item, @"\d+$")[0] + "." + Regex.Match(item, @"\d+$").Value;


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

                    if (checkBox6.Checked)
                    {
                        new TeklaMacroBuilder.MacroBuilder()
                                    .Callback("acmd_create_marks_all", "", "main_frame").Run();
                    }

                    new TeklaMacroBuilder.MacroBuilder()
                        .ValueChange("gr_close_dr_editor_confirm_instance", "gr_close_save_dr_editor_freeze", "1")
                        .PushButton("gr_close_save_dr_editor_yes", "gr_close_dr_editor_confirm_instance").Run();

                    ////////////////////////////////////////////////////////////////remover marcas de parafusos 

                    new TeklaMacroBuilder.MacroBuilder()
                       .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                       .PushButton("gr_adraw_smark", "adraw_dial")
                       .TableSelect("adsm_dial", "gr_mark_selected_elements", new int[] { 1, 2, 3, 4, 5, 6 })
                       .PushButton("gr_remove_element", "adsm_dial")
                       .PushButton("dsm_modify", "adsm_dial")
                       .PushButton("dsm_apply", "adsm_dial")
                       .PushButton("gr_adraw_ok", "adraw_dial").Run();

                    ArrayList MyViews = new ArrayList();
                    DrawingHandler MyDrawingHandler = new DrawingHandler();
                    Drawing MyCurrentDrawing = MyDrawingHandler.GetActiveDrawing();

                    foreach (DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
                    {
                        if (drawingObject is Tekla.Structures.Drawing.View)
                        {

                            Tekla.Structures.Drawing.View drawingObj = (TSD.View)drawingObject;
                            //if (drawingObj.ViewType != TSD.View.ViewTypes._3DView)
                            //{
                            MyViews.Add(drawingObject);
                            //}
                        }
                    }

                    if (MyViews.Count > 0)
                    {
                        MyDrawingHandler.GetDrawingObjectSelector().SelectObjects(MyViews, false);
                    }
                    new TeklaMacroBuilder.MacroBuilder()
                    .Callback("acmd_display_selected_drawing_object_dialog", "", "main_frame")
                    .PushButton("view_on_off", "view_dial")
                    .TableSelect("view_dial", "gr_vsmark_selected_elements", new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9 })
                    .PushButton("gr_vsm_remove_element", "view_dial")
                    .PushButton("view_modify", "view_dial")
                    .PushButton("view_ok", "view_dial")

                    .Run();

                    ////////////////////////////////////////////////////////////////

                    DrawingHandler dh = new DrawingHandler();
                    ViewBase _sheet = dh.GetActiveDrawing().GetSheet();
                    Text text = new Text(_sheet, new Tekla.Structures.Geometry3d.Point(285, 70), comboBox2.Text);
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
            new TeklaMacroBuilder.MacroBuilder()
             .Callback("acmdDisplayEnvironmentVariablesDialogFromMenu", "", "main_frame")
             .PushButton("gr_adraw_on_off", "adraw_dial")
             .ValueChange("diaEnvironmentVariableDialog", "FindString", "​FIXED")
             .ValueChange("diaEnvironmentVariableDialog", "AllCategoriesCheckButton", "1")
             .TableSelect("diaEnvironmentVariableDialog", "tblEnvironmentVariables", new int[] { 5 })
             .TableValueChange("diaEnvironmentVariableDialog", "tblEnvironmentVariables", "colValue", "TRUE")
             .PushButton("butApply", "diaEnvironmentVariableDialog")
             .PushButton("butOk", "diaEnvironmentVariableDialog")
             .Run();

        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                //do something
            }
            else
            {
                if (e.KeyChar == Convert.ToChar(Keys.Back))
                {

                }
                else
                {
                    e.Handled = true;
                }

            }
        }

        private void FrmSoldadura_Load(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = 0;
        }
    }
}
