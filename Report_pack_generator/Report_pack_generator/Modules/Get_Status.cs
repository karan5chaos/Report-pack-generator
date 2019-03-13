using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Report_pack_generator.Modules
{
    class Get_Status
    {
        public static List<string> staus_messages = new List<string>();
        public static List<string> error_messages = new List<string>();

        public static void post_log(ListBox box)
        {
            box.Items.Clear();
            foreach (var message in staus_messages)
            {
                box.Items.Add(message);
            }
            

        }

        public static void clear_lists()
        {
            staus_messages.Clear();
            error_messages.Clear();

        }

        public static void set_error(Exception exception_message ,string filename, string reportspack_section)
        {
            error_messages.Add(reportspack_section + " > File - "+ Path.GetFileNameWithoutExtension(filename) + "/nMessage - " + exception_message.Message);
        }

        public static void post_error_log(ListBox box)
        {
            box.Items.Clear();
            foreach (var message in error_messages)
            {
                box.Items.Add(message);
            }


        }
    }
}
