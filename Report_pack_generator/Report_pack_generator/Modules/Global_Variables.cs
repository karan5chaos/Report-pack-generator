using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Report_pack_generator.Modules
{
    static class Global_Variables
    {

        public static class Status_Buttons
        {
            public static Button pdfbutton = null;
            public static Button folderbutton = null;
            public static Button rpbutton = null;
            public static Button combutton = null;
            public static Button abutton = null;
            public static Button plbutton = null;
            public static Button borbutton = null;
            public static Button arbutton = null;


            public static void Assign_buttons()
            {
                Global_Variables.Status_Buttons.pdfbutton = compact_instance.c_mode_1.pdf_btn;
                Global_Variables.Status_Buttons.folderbutton = compact_instance.c_mode_1.folder_btn;
                Global_Variables.Status_Buttons.rpbutton = compact_instance.c_mode_1.rp_btn;
                Global_Variables.Status_Buttons.combutton = compact_instance.c_mode_1.com_btn;
                Global_Variables.Status_Buttons.abutton = compact_instance.c_mode_1.a_btn;
                Global_Variables.Status_Buttons.plbutton = compact_instance.c_mode_1.pl_btn;
                Global_Variables.Status_Buttons.borbutton = compact_instance.c_mode_1.bor_btn;
                Global_Variables.Status_Buttons.arbutton = compact_instance.c_mode_1.ar_btn;
            }


        }



        public static string connectionString = "Server=gbips-i-db700;Database=FINMAR_Placement;Integrated Security=True;";

        

    }

   


}
