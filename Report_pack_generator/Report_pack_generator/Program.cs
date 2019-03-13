using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;

namespace Report_pack_generator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            string resource1 = "Report_pack_generator.lib.itextsharp.dll";
            string resource2 = "Report_pack_generator.lib.Ookii.Dialogs.dll";
            string resource3 = "Report_pack_generator.lib.Ookii.Dialogs.resources.dll";
            string resource4 = "Report_pack_generator.lib.Transitions.dll";

            EmbeddedAssembly.Load(resource1, "itextsharp.dll");
            EmbeddedAssembly.Load(resource2, "Ookii.Dialogs.dll");
            EmbeddedAssembly.Load(resource3, "Ookii.Dialogs.resources.dll");
            EmbeddedAssembly.Load(resource4, "Transitions.dll");

            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Main_Page());
        }

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name);
        }

    }
}
