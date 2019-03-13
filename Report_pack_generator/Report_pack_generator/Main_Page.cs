using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Ookii.Dialogs;
using System.IO;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using iTextSharp;
using iTextSharp.text.pdf;
using System.Threading;
using iTextSharp.text;
using System.Reflection;
using Transitions;
using System.Data.SqlClient;
using Report_pack_generator.Settings;
using Report_pack_generator.Modules;
using System.Windows.Forms.DataVisualization.Charting;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Report_pack_generator
{
    public partial class Main_Page : Form
    {
        public Main_Page()
        {
            InitializeComponent();
            compact_instance.frm1 = this;
            Functions.get_Functions.SetDoubleBuffering(listBox1,true);
            Functions.get_Functions.SetDoubleBuffering(panel2, true);
            Functions.get_Functions.SetDoubleBuffering(panel1, true);
        }
       
        //==========================
        //REQUIRED VARIABLES.
        //==========================
        string connectionString = "Server=--server--;Database=--database--;Integrated Security=True;";

        Button pdfbutton = null;
        Button folderbutton = null;
        Button rpbutton = null;
        Button combutton = null;
        Button abutton = null;
        Button plbutton = null;
        Button borbutton = null;
        Button arbutton = null;
        Excel.Application app = new Excel.Application();

        //Common variables for processing
        //-------------------------------
        string carrier_name;  //<---- get carrier name for report pack.
        string period; //<---- Current year.
        List<string> status = new List<string>(); //<---- list for status reporting.
        Stopwatch stopWatch = new Stopwatch(); //<---- for timing process.
        List<string> fileNames; //<---- list of pages for RP compilation. (modify with CAUTION !!)
        string commentary_paths; //<---- sets current commentary path.
        string gen_paths; //<---- sets current generated reports path.
        //string region_gen_path; //<---- sets regional path. only used by vista folderbrowsers.
        string m_q; //<--- sets "Mth" & "Qtr" based on the selected dropdown.
        Compact_Mode c_mode;
        int tier;
        string region;
        //bool merged = false;
        private bool mouseDown;
        private Point lastLocation;
        const int WS_MINIMIZEBOX = 0x20000;
        const int CS_DBLCLKS = 0x8;
        List<string> BUs = new List<string>();
        List<string> Carriers = new List<string>();
        string status_lbl_text;
        Dictionary<string, double> folders_list = new Dictionary<string, double>();
        //------------------------------------

        //for check_process method only. Used for flagging different processing in visual feedback.
        //-----------------------------------------------------------------------------------------
        bool comms =false; //<--- checks and flags commentary process.
        bool rp = false; //<--- checks and flags report pack compilation process.
        bool ar = false; //<--- checks and flags additional requirements process.
        bool a = false; //<--- checks and flags analysis process.
        bool folder_ = false; //<--- checks and flags files and folder copying process.
        bool pl = false; //<--- checks and flags pipline process.
        bool bor = false; //<--- checks and flags borderaux process. this includes additional borderaux as well.
        bool pdf = false; //<--- checks and flags pdf conversion process. this gets flagged anytime the file is being converted to PDF.


        //for report sections options only ! determines if a specific section to be included in report pack or not.
        //---------------------------------------------------------------------------------------------------------
        bool commentary = true; //<--- if false, will skip commentary files checking.
        bool pipline = true; //<--- if false, pipeline files will be skipped during compilation.
        bool bordx = true; //<--- if false, borderaux files will be skipped during the compilation. this includes additional borderaux files as well.
        bool analysis = true; //<--- if false, analysis files will be skipped during compilation.
        bool add_requirements = true; //<--- if false, additional files will not be processed. this includes, CRM reports and Others not shown verbatim reports.
        bool coc = true; //<--- if false, cover for client will be skipped during compilation.
        bool dss = true; //<--- if false, will not copy data sources and service summary file.


        //iscompleted method only. iscompleted() checks if a certain process has been finished. Works in conjuntion with check_process() method.
        //--------------------------------------------------------------------------------------------------------------------------------------
        bool iscommentary = false; //<--- checks and flags as true when commentary file has been processed.
        bool ispipline = false; //<--- checks and flags as true when pipline file has been processed.
        bool isbordx = false; //<---checks and flags as true when borderaux file has been processed.
        bool isanalysis = false;  //<--- checks and flags as true when analysis file has been processed.
        bool isaddreq = false;  //<--- checks and flags as true when additonal requirments files has been processed.
        bool isrp = false; //<--- checks and flags as true when report pack compilaton has be completed.


        //This method creates a folder browser dialog for selecting both client service root folder & commentaries root folder.
        //takes descripton and path as parameters.
        VistaFolderBrowserDialog create_openDialog(string folder_path, string description)
        {
            VistaFolderBrowserDialog dlg2 = new VistaFolderBrowserDialog();
            dlg2.ShowNewFolderButton = false;
            dlg2.SelectedPath = folder_path;
            dlg2.Description = description;

            return dlg2;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.Style |= WS_MINIMIZEBOX;
                cp.ClassStyle |= CS_DBLCLKS;
                return cp;
            }
        }

        void set_buttons()
        {

                pdfbutton = compact_instance.c_mode_1.pdf_btn;
                folderbutton = compact_instance.c_mode_1.folder_btn;
                rpbutton = compact_instance.c_mode_1.rp_btn;
                combutton = compact_instance.c_mode_1.com_btn;
                abutton = compact_instance.c_mode_1.a_btn;
                plbutton = compact_instance.c_mode_1.pl_btn;
                borbutton = compact_instance.c_mode_1.bor_btn;
                arbutton = compact_instance.c_mode_1.ar_btn;
        }

        //checks and report backs on going process with color coding.
        //It is purely cometic and doesn't affects the main functionality of the RP creation.
        void check_process()
        {
           // set_buttons();

            //Resets the button colors to default colors.
            rpbutton.BackColor = SystemColors.ControlLight; 
            plbutton.BackColor = SystemColors.ControlLight;
            combutton.BackColor = SystemColors.ControlLight;
            abutton.BackColor = SystemColors.ControlLight;
            folderbutton.BackColor = SystemColors.ControlLight;
            arbutton.BackColor = SystemColors.ControlLight;
            borbutton.BackColor = SystemColors.ControlLight;
            pdfbutton.BackColor = SystemColors.ControlLight;

            rp_btn.BackColor = SystemColors.ControlLight;
            pl_btn.BackColor = SystemColors.ControlLight;
            com_btn.BackColor = SystemColors.ControlLight;
            a_btn.BackColor = SystemColors.ControlLight;
            folder_btn.BackColor = SystemColors.ControlLight;
            ar_btn.BackColor = SystemColors.ControlLight;
            bor_btn.BackColor = SystemColors.ControlLight;
            pdf_btn.BackColor = SystemColors.ControlLight;


            if (pdf)
            {
                Transition.run(pdf_btn, "BackColor", Color.Firebrick, new TransitionType_Flash(2, 500));
                Transition.run(pdfbutton, "BackColor", Color.Firebrick, new TransitionType_Flash(2, 500));
            }
            else
            {
                Transition.run(pdf_btn, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
                Transition.run(pdfbutton, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
            }

            if (rp)
            {
                Transition.run(rp_btn, "BackColor", Color.LightSteelBlue, new TransitionType_Flash(2, 500));
                Transition.run(rpbutton, "BackColor", Color.LightSteelBlue, new TransitionType_Flash(2, 500));

            }
            else
            {
                Transition.run(rp_btn, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
                Transition.run(rpbutton, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
            }
            if (comms)
            {
                Transition.run(com_btn, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
                Transition.run(combutton, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
            }
            else
            {
                Transition.run(com_btn, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
                Transition.run(combutton, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
            }
            
            if (pl)
            {
                Transition.run(pl_btn, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
                Transition.run(plbutton, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
            }
            else
            {
                Transition.run(pl_btn, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
                Transition.run(plbutton, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
            }
            if (bor)
            {
                Transition.run(bor_btn, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
                Transition.run(borbutton, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
            }
            else
            {
                Transition.run(bor_btn, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
                Transition.run(borbutton, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
            }
            if (a)
            {
                Transition.run(a_btn, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
                Transition.run(abutton, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
            }
            else
            {
                Transition.run(a_btn, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
                Transition.run(abutton, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
            }
            if (ar)
            {
                Transition.run(ar_btn, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
                Transition.run(arbutton, "BackColor", Color.LimeGreen, new TransitionType_Flash(2, 500));
            }
            else
            {
                Transition.run(ar_btn, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
                Transition.run(arbutton, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
            }
            if (folder_)
            {
                Transition.run(folder_btn, "BackColor", Color.LightSteelBlue, new TransitionType_Flash(2, 500));
                Transition.run(folderbutton, "BackColor", Color.LightSteelBlue, new TransitionType_Flash(2, 500));
            }
            else
            {
                Transition.run(folder_btn, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
                Transition.run(folderbutton, "BackColor", SystemColors.Control, new TransitionType_Acceleration(1000));
            }

            iscompleted(); //<--- call is completed method and check for the current progress.
        }

        //Reset button colors and flags progress variables to false. part of iscompleted().
        void reset_completed()
        {

           // set_buttons();

            isrp = false;
            iscommentary = false;
            ispipline = false;
            isbordx = false;
            isanalysis = false;
            isaddreq = false;

            rp_btn.BackColor = SystemColors.ControlLight;
            pl_btn.BackColor = SystemColors.ControlLight;
            com_btn.BackColor = SystemColors.ControlLight;
            a_btn.BackColor = SystemColors.ControlLight;
            folder_btn.BackColor = SystemColors.ControlLight;
            ar_btn.BackColor = SystemColors.ControlLight;
            bor_btn.BackColor = SystemColors.ControlLight;
            pdf_btn.BackColor = SystemColors.ControlLight;

            rpbutton.BackColor = SystemColors.ControlLight;
            combutton.BackColor = SystemColors.ControlLight;
            plbutton.BackColor = SystemColors.ControlLight;
            borbutton.BackColor = SystemColors.ControlLight;
            abutton.BackColor = SystemColors.ControlLight;
            arbutton.BackColor = SystemColors.ControlLight;
        
            
        }

        //checks for processes which have been completed. woks in conjuntion with check_process() method.
        void iscompleted()
        {
           // set_buttons();

            if (isrp)
            {
                Transition.run(rp_btn, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
                Transition.run(rpbutton, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
            }
            

            if (iscommentary)
            {

                Transition.run(com_btn, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
                Transition.run(combutton, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
            }


            if (ispipline)
            {
                Transition.run(pl_btn, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
                Transition.run(plbutton, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
            }

            if (isbordx)
            {
                Transition.run(bor_btn, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
                Transition.run(borbutton, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
            }

            if (isanalysis)
            {
                Transition.run(a_btn, "BackColor", Color.DarkCyan, new TransitionType_Flash(1,1000));
                Transition.run(abutton, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
            }

            if (isaddreq)
            {
                Transition.run(ar_btn, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
                Transition.run(arbutton, "BackColor", Color.DarkCyan, new TransitionType_Flash(1, 1000));
            }
        }

        public void button1_Click(object sender, EventArgs e)
        {

            bool gen = false; //<--- to check if client service root path  has been selected.
            bool comms = false; //<--- checks if commentaries root path has been selected.
            string dirName = string.Empty; //<--- gets name of the last directory of the client service selected path.


            if (comboBox1.Text == "" || comboBox3.Text == "") //<--- checks if period fields are not left blank.
            {
                listBox1.Items.Add("Period cannot be blank..");
            }
            else
            {

                set_reportsections(); //<--- check and set which report pack sections will be processed.

                //create and show dailogbox
                var generated_dialogbox = create_openDialog(report_pack.Default.generated_path, "Select generated files folder");
                //check if any path was selected in the dialog box
                if (generated_dialogbox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    gen_paths = generated_dialogbox.SelectedPath;
                    dirName = new DirectoryInfo(gen_paths).Name;
                    gen = true;
                }
                    //return false and notify user that path was not selected.
                else
                {
                    listBox1.Items.Add("Generated report pack folder not selected..");
                    gen = false;
                }

                //create and show commentary file selection dialog box.
                if (commentary)
                {
                    //only show dialog box if generated files folder is selected.
                    var commentary_dialogbox = create_openDialog(report_pack.Default.commentary_path, "Selected period : " + dirName);
                    if (gen)
                    {
                        if (commentary_dialogbox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            commentary_paths = commentary_dialogbox.SelectedPath;
                            comms = true;

                            commentary_dialogbox.Dispose();
                            generated_dialogbox.Dispose();

                        }
                            //return false if no commentary folder is selected.
                        else
                        {
                            listBox1.Items.Add("Commentary folder not selected..");
                            comms = false;

                            commentary_dialogbox.Dispose();
                            generated_dialogbox.Dispose();
                        }
                    }
                }
                else
                {
                    comms = true;
                }

                // check if both commentary and generated report folder has been selected.
                if (gen && comms)
                {

                    reset_completed(); //<--- resets changes made may iscompleted() method.

                    fileNames = new List<string>(); //<--- creates a list in memory which will hold the names of reports.
                                                    //Reports are compiled after the list is populated.
                    timer1.Start(); //<--- strats timer for monitoring progress.     
                    stopWatch.Start(); //<--- starts stopwatch for calculating process time.
                    listBox1.Items.Clear();
                    status_label_.Text = "Processing.."; //<--- adds an entry to status list.
                  //  progress_bar.Visible = true; //<--- show progress bar at bootom. Progress bar is haidden by deafult.
                    //initiate report pack background process.
                    button7.Enabled = false;  //<--- enable button. button is disabled unless all conditions are met.
                    this.Text = "RPH - " + region + " " + period + " "+ carrier_name;

                    compact_instance.c_mode_1.reg_lbl.Text = region;
                    compact_instance.c_mode_1.p_lbl.Text = period;
                    compact_instance.c_mode_1.freq_lbl.Text = m_q;
                    compact_instance.c_mode_1.c_lbl.Text = carrier_name;


                    
                    initate_Process.RunWorkerAsync(); //<--- starts background process for report pack creation.
                }
            }
            
        }

        bool coms_pdf = true;
        bool pipe_pdf = true;
        bool analysis_pdf = true;
        bool bordx_pdf = true;

        // determines and flags which report pack sections would be included in processing queue.
        void set_reportsections()
        {
            //get existing status of the options selected.
            commentary = comm_box.Checked; 
            pipline = pip_box.Checked;
            add_requirements = ar_box.Checked;
            bordx = brdx_box.Checked;
            analysis = ana_box.Checked;
            dss = dss_box.Checked;

            coms_pdf = com_pdf.Checked;
            pipe_pdf = pip_pdf.Checked;
            analysis_pdf = a_pdf.Checked;
            bordx_pdf = bord_pdf.Checked;

            //enable disabled relevant options based on the selection
            if (!pipline)
            {
                groupBox5.Enabled = false;
            }
            else
            {
                groupBox5.Enabled = true;
            }

            if (!bordx)
            {
                groupBox4.Enabled = false;
                groupBox10.Enabled = false;  
            }
            else
            {
                groupBox4.Enabled = true;
                groupBox10.Enabled = true;
            }

            if (!analysis)
            {
                groupBox6.Enabled = false;
            }
            else
            {
                groupBox6.Enabled = true;
            }

            if (!add_requirements)
            {
                groupBox11.Enabled = false;
            }
            else
            {
                groupBox11.Enabled = true;
            }
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            //opens folder browser for selecting output path.
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
                report_pack.Default.output_path = folderBrowserDialog1.SelectedPath;
                report_pack.Default.Save();

                //button7.Enabled = true;
            } 
        }

        private void generate_directory(System.Collections.Specialized.StringCollection folder_names, string path)
        {
            //creates required folders in the output path, searches for files in generated reports folder
            // and places them in the output folder
            try
            {

                foreach (string folder in folder_names)
                {
                    var combined_path = path + "/" + carrier_name + "/" + folder;
                    
                    if (!Directory.Exists(combined_path))
                    {
                        Directory.CreateDirectory(combined_path);  
                    }
                        if (place_file(folder, combined_path))
                        {
                            Get_Status.staus_messages.Add(folder + " requested to abort operation..");
                            Directory.Delete(path + "/" + carrier_name, true);
                            if (!initate_Process.IsBusy)
                            {
                                //initiate the background process for report pack creation.
                                initate_Process.CancelAsync();
                            }
                            break;
                        }
                }
            }
            catch(Exception ex)
            {
                Get_Status.staus_messages.Add("Error while creating directoreis.." + ex.Message);
            }
        }

        static string GetNameOf<T>(Expression<Func<T>> property)
        {
            return (property.Body as MemberExpression).Member.Name;
            
        }


        bool place_file(string folder_name, string path)
        {
            folder_ = false;
            comms = false;
            pl = false;
            a = false;
            ar = false;
            bor = false;   
            bool abort = false;
            
            switch (folder_name)
            {
                case "Commentaries":
                    {
                            Stopwatch com_watch = new Stopwatch();
                            com_watch.Start();
                            if (commentary)
                            {
                                try
                                {
                                    bool found = false;
                                    comms = true;

                                    Get_Status.staus_messages.Add("Copying " + folder_name + " cover page..");
                                    var s = GetNameOf(() => Properties.Resources.TOC_for_UK_Tier_2_Revised);
                                    var cover_path = path + "/" + s + ".pdf";
                                    File.WriteAllBytes(cover_path, Properties.Resources.TOC_for_UK_Tier_2_Revised);
                                    fileNames.Add(cover_path);
                                    Get_Status.staus_messages.Add(folder_name + " cover page copied..");

                                    var direct = Directory.EnumerateFiles(commentary_paths, "*.docx", SearchOption.AllDirectories);

                                    if (direct.Count() > 0)
                                    {
                                        var linqq =
                                        from file in direct
                                        where (file.Contains("Final") || file.Contains("FINAL") || file.Contains("final")) &&
                                            Path.GetFileName(file).Contains(carrier_name)
                                        select file;


                                        foreach (string file in linqq)
                                        {
                                            var copy_file = path + "/" + Path.GetFileName(file);
                                            File.Copy(file, copy_file);
                                            Get_Status.staus_messages.Add(folder_name + " file(s) copied..");

                                            if (coms_pdf)
                                            {
                                                var pdffile = path + "/" + Path.GetFileNameWithoutExtension(file) + ".pdf";

                                                Functions.get_Functions.convert_ToPDF(1, file, pdffile, folder_name,app);
                                                fileNames.Add(pdffile);
                                            }
                                            found = true;
                                        }
                                    }
                                    else
                                    {
                                        Get_Status.staus_messages.Add(folder_name + " file not found..");
                                        abort = true;
                                    }

                                }
                                catch (Exception xc)
                                {
                                    abort = true;
                                    Get_Status.staus_messages.Add("Error occured in " + folder_name + " section " + xc.Message);

                                }

                                finally
                                {
                                    if (abort)
                                        iscommentary = false;
                                    else
                                        iscommentary = true;

                                }
                            }

                            else
                            {
                                Get_Status.staus_messages.Add("Skipping " + folder_name);

                            }                    
                        com_watch.Stop();

                        double duration = Math.Round(com_watch.Elapsed.TotalMinutes, 1);
                        folders_list.Add(folder_name, duration);
                    }
                    break;

                case "Pipeline":
                    {
                        Stopwatch pipe_watch = new Stopwatch();
                        pipe_watch.Start();
                        if (pipline)
                        {
                            try
                            {
                                pl = true;

                                Get_Status.staus_messages.Add("Copying " + folder_name + " cover page..");

                                var s = GetNameOf(() => Properties.Resources.Section_3___pipeline_front_cover);
                                var cover_path = path + "/" + s + ".pdf";
                                File.WriteAllBytes(cover_path, Properties.Resources.Section_3___pipeline_front_cover);
                                fileNames.Add(cover_path);
                                Get_Status.staus_messages.Add(folder_name + " cover page copied..");

                                if (search_genereatedreports(path, folder_name))
                                {
                                    var files = Directory.EnumerateFiles(path, "*.*", SearchOption.AllDirectories)
                                                .Where(sa => sa.EndsWith(".xlsx") || sa.EndsWith(".xls"));

                                    var filterfiles = from f in files
                                                      where !f.Contains("INTERNAL") && f.Contains(carrier_name)
                                                      orderby f
                                                      select f;

                                    foreach (var xlfile in filterfiles)
                                    {
                                            Functions.get_Functions.unhide_Columns(xlfile,app);
                                            Functions.get_Functions.warp_text(xlfile,app);

                                            var pdffile = path + "/" + Path.GetFileNameWithoutExtension(xlfile) + ".pdf";

                                            if (pipe_pdf)
                                            {
                                                pdf = true;
                                                Functions.get_Functions.convert_ToPDF(0, xlfile, pdffile, folder_name,app);
                                                pdf = false;
                                                fileNames.Add(pdffile);
                                            }
                                    }
                                }
                                else
                                {
                                    abort = true;
                                    Get_Status.staus_messages.Add(folder_name + " files not found.. Aborting report pack creation..");
                                }
                            }
                            catch (Exception xc)
                            {
                                abort = true;
                                Get_Status.staus_messages.Add("Error occured in " + folder_name + " section " + xc.Message);

                            }

                            finally
                            {
                                if (abort)
                                    ispipline = false;
                                else
                                    ispipline = true;
                            }
                        }
                        else
                        {
                            Get_Status.staus_messages.Add("Skipping " + folder_name);
                        }
                        pipe_watch.Stop();
                        double duration = Math.Round(pipe_watch.Elapsed.TotalMinutes, 1);
                        folders_list.Add(folder_name, duration);
                    }
                    break;

                case "Cover for Client":
                    {
                        Stopwatch coc_watch = new Stopwatch();
                        coc_watch.Start();

                        if (coc)
                        {
                            try
                            {
                                    Get_Status.staus_messages.Add("Copying " + folder_name + " cover page..");
                                    var s = GetNameOf(() => Properties.Resources.COVERS_FOR_CLIENT);
                                    File.WriteAllBytes(path + "/" + s + ".pdf", Properties.Resources.COVERS_FOR_CLIENT);
                                    File.WriteAllBytes(path + "/" + s + ".docx", Properties.Resources.COVERS_FOR_CLIENT_);

                                    var cover_path = path + "/" + s + ".pdf";
                                    fileNames.Add(cover_path);
                                    Get_Status.staus_messages.Add(folder_name + " cover page copied..");
                            }
                            catch (Exception xc)
                            {
                                abort = true;
                                Get_Status.staus_messages.Add("Error occured in " + folder_name + " section " + xc.Message);
                            }
                        }
                        else
                        {
                            Get_Status.staus_messages.Add("Skipping " + folder_name);
                        }
                        coc_watch.Stop();
                        double duration = Math.Round(coc_watch.Elapsed.TotalMinutes, 1);
                        folders_list.Add(folder_name, duration);
                    }
                    break;

                case "Analysis":
                    {
                         Stopwatch ana_watch = new Stopwatch();
                         ana_watch.Start();
                        if (analysis)
                        {
                           
                            try
                            {
                                a = true;

                                Get_Status.staus_messages.Add("Copying " + folder_name + " cover page..");
                                var s = GetNameOf(() => Properties.Resources.Performance_and_Analysis_front_page___section_4__tier_II_);
                                var cover_path = path + "/" + s + ".pdf";
                                File.WriteAllBytes(cover_path, Properties.Resources.Performance_and_Analysis_front_page___section_4__tier_II_);
                                fileNames.Add(cover_path);
                                Get_Status.staus_messages.Add(folder_name + " cover page copied..");


                                if (search_genereatedreports(path, folder_name))
                                {
                                    //sort alphabetically.
                                    var sorted = Directory.GetFiles(path, "*.xlsx");
                                    Array.Sort(sorted, StringComparer.InvariantCulture);

                                    var analysisfile = new List<string>();
                                    
                                    //make sure others section is filtered last.
                                    //first filtering excluding "other" files
                                    foreach (var xlfile in sorted)
                                    {
                                        if (!xlfile.Contains("Other"))
                                        {
                                            var pdffile = path + "/" + Path.GetFileNameWithoutExtension(xlfile) + ".pdf";
                                            if (excel_settings.Default.trim_2_page || excel_settings.Default.trim_3_page)
                                            {
                                                Functions.get_Functions.rm_Sheets(xlfile,app);
                                            }
                                            if (analysis_pdf)
                                            {
                                                pdf = true;
                                                Functions.get_Functions.convert_ToPDF(0, xlfile, pdffile, folder_name,app);
                                                pdf = false;
                                                analysisfile.Add(pdffile);
                                            }
                                        }

                                    }
                                    //second filter for only "other" files.
                                    foreach (var xlfile in sorted)
                                    {
                                        if (xlfile.Contains("Other") || xlfile.Contains("OTHER") || xlfile.Contains("other"))
                                        {
                                            var pdffile = path + "/" + Path.GetFileNameWithoutExtension(xlfile) + ".pdf";

                                            if (excel_settings.Default.trim_2_page || excel_settings.Default.trim_3_page)
                                            {
                                                Functions.get_Functions.rm_Sheets(xlfile,app);
                                            }

                                            if (analysis_pdf)
                                            {
                                                pdf = true;
                                                Functions.get_Functions.convert_ToPDF(0, xlfile, pdffile, folder_name,app);
                                                pdf = false;
                                                analysisfile.Add(pdffile);
                                            }
                                        }
                                    }

                                    if (analysisfile.Count > 0)
                                    {
                                        fileNames.AddRange(analysisfile);
                                        analysisfile.Clear();
                                    }
                                }
                                else
                                {
                                    abort = true;
                                    Get_Status.staus_messages.Add(folder_name + " files not found.. Aborting report pack creation..");
                                }
                            }
                            catch (Exception xc)
                            {
                                abort = true;
                                Get_Status.staus_messages.Add("Error occured in "+ folder_name + " section " + xc.Message);
                            }

                            finally
                            {
                                if (abort)
                                    isanalysis = false;
                                else
                                    isanalysis = true;

                            }
                        }
                        else
                        {
                            Get_Status.staus_messages.Add("Skipping " + folder_name);
                        }
                        ana_watch.Stop();
                        double duration = Math.Round(ana_watch.Elapsed.TotalMinutes, 1);
                        folders_list.Add(folder_name, duration);

                    }
                    break;


                case "Bordereaux":
                    {
                        Stopwatch bordx_watch = new Stopwatch();
                        bordx_watch.Start();
                        if (bordx)
                        {
                            try
                            {
                                bor = true;
                                if (search_genereatedreports(path, folder_name))
                                {
                                    foreach (var xlfile in Directory.GetFiles(path, "*.xlsx"))
                                    {
                                        if (Path.GetFileName(xlfile).Contains("Additional") && excel_settings.Default.trim_additionbrdx)
                                        {
                                            Functions.get_Functions.rm_Sheet_Additional_Borderaux(xlfile,app);
                                            Get_Status.staus_messages.Add(Path.GetFileName(xlfile) + " - Not shown removed..");
                                        }
                                        else
                                        {
                                            Functions.get_Functions.unhide_Columns(xlfile,app);
                                            Functions.get_Functions.remove_NewRenewColumn(xlfile,app);
                                        }

                                        if(bordx_pdf)
                                            Functions.get_Functions.convert_ToPDF(0,xlfile, path + "/" + Path.GetFileNameWithoutExtension(xlfile) + ".pdf", folder_name,app);
                                    }
                                }
                                else
                                {
                                   // abort = true;
                                    Get_Status.staus_messages.Add(folder_name + " files not found.. Aborting report pack creation..");
                                }
                            }
                            catch (Exception xc)
                            {
                                //abort = true;
                                Get_Status.staus_messages.Add("Error occured in " + folder_name + " section " + xc.Message);
                            }

                            finally
                            {
                                //if (abort)
                                //    isbordx = false;
                                //else
                                    isbordx = true;
                            }
                        }
                        else
                        {
                            Get_Status.staus_messages.Add("Skipping " + folder_name);
                        }
                        bordx_watch.Stop();
                        double duration = Math.Round(bordx_watch.Elapsed.TotalMinutes, 1);
                        folders_list.Add(folder_name, duration);
                    }
                    break;

                case "Data Sources and Service Summary":
                    {
                        if (dss)
                        {
                            Get_Status.staus_messages.Add("Copying "+folder_name);
                            var s = GetNameOf(() => Properties.Resources.Data_Sources_and_Service_Summary);
                            File.WriteAllBytes(path + "/" + s + ".pdf", Properties.Resources.Data_Sources_and_Service_Summary);
                            Get_Status.staus_messages.Add(folder_name + " pages copied..");
                        }
                        else
                        {
                            Get_Status.staus_messages.Add("Skipping " + folder_name);
                        }
                    }
                    break;

                case "Additional Requirements":
                    {
                        Stopwatch ar_watch = new Stopwatch();
                        ar_watch.Start();
                        if (add_requirements)
                        {
                            try
                            {
                                ar = true;
                                if (search_genereatedreports(path, "CRM Reports"))
                                {
                                    foreach (var xlfile in Directory.GetFiles(path, "*.xlsx"))
                                    {
                                        if (excel_settings.Default.rm_CRM && (Path.GetFileName(xlfile).Contains("Pocket") || Path.GetFileName(xlfile).Contains("CRM")))
                                        {

                                            if (Functions.get_Functions.rm_Nodata(xlfile,app))
                                            {
                                                Get_Status.staus_messages.Add("Error occured in in processing CRM report - "+Path.GetFileName(xlfile));
                                            }
                                            else
                                            {
                                                Get_Status.staus_messages.Add(Path.GetFileName(xlfile) + " - sheets with 'No Data' removed..");
                                            }
                                        }
                                    }
                                }

                                if (search_genereatedreports(path, "Other not shown verbatim"))
                                {
                                    foreach (var xlfile in Directory.GetFiles(path, "*.xlsx"))
                                    {
                                        if (Path.GetFileName(xlfile).Contains("Not Shown") && excel_settings.Default.rm_Verbatim)
                                        {
                                            if (Functions.get_Functions.rm_Nodata(xlfile,app))
                                            {
                                                Get_Status.staus_messages.Add("Error occured in in processing verbatim report - " + Path.GetFileName(xlfile));
                                            }
                                            else
                                            {
                                                Get_Status.staus_messages.Add(Path.GetFileName(xlfile) + " - sheets with 'No Data' removed..");
                                            }

                                            Get_Status.staus_messages.Add(Path.GetFileName(xlfile) + " - no data sheet removed");
                                        }
                                    }
                                }
                                else
                                {
                                    Get_Status.staus_messages.Add(folder_name + " files not found.. ..");
                                }

                                if (search_genereatedreports(path, "Broker Decision"))
                                {
                                    foreach (var xlfile in Directory.GetFiles(path, "*.xlsx"))
                                    {
                                        if (Path.GetFileName(xlfile).Contains("Broker") && excel_settings.Default.rm_Brokerdecision)
                                        {
                                            if (Functions.get_Functions.rm_Nodata(xlfile,app))
                                            {
                                                Get_Status.staus_messages.Add("Error occured in in processing Broker decision report - " + Path.GetFileName(xlfile));
                                            }
                                            else
                                            {
                                                Get_Status.staus_messages.Add(Path.GetFileName(xlfile) + " - sheets with 'No Data' removed..");
                                            }

                                            Get_Status.staus_messages.Add(Path.GetFileName(xlfile) + " - no data sheet removed");
                                        }
                                    }
                                }
                                else
                                {
                                    Get_Status.staus_messages.Add(folder_name + " files not found.. ..");
                                }


                            }
                            catch(Exception ecx)
                            {
                                Get_Status.staus_messages.Add(folder_name + " Error occured in "+folder_name+". "+ecx.Message);
                            }

                            finally
                            {
                                    isaddreq = true;
                            }
                        }
                        else
                        {
                            Get_Status.staus_messages.Add("Skipping " + folder_name);
                        }
                        ar_watch.Stop();
                        double duration = Math.Round(ar_watch.Elapsed.TotalMinutes, 1);
                        folders_list.Add(folder_name, duration);
                    }
                    break;
            }
            return abort;
        }

        bool search_genereatedreports(string path, string folder_name)
        {
            bool found = false;
            try
            {
                if (Directory.Exists(gen_paths))
                {
                    folder_ = true;
                    var dynamicpath = gen_paths + "/" + folder_name;


                    if (m_q == "Qtr" && tier == 2)
                    {
                        if (!folder_name.Contains("Pipeline") && Directory.Exists(dynamicpath + "/Tier II Qtr"))
                        {
                            dynamicpath += "/Tier II Qtr";
                        }
                    }
                    else if (tier == 3)
                    {
                        if (!folder_name.Contains("Pipeline") && Directory.Exists(dynamicpath + "/Tier III"))
                        {
                            dynamicpath += "/Tier III";
                        }
                    }


                    var files = Directory.EnumerateFiles(dynamicpath, "*.xlsx", SearchOption.TopDirectoryOnly);



                    if (files.Count() > 0)
                    {
                        foreach (string file in files)
                        {
                            var copy_file = path + "/" + Path.GetFileName(file);

                            if (Path.GetFileName(file).Contains(carrier_name))
                            {
                                File.Copy(file, copy_file);
                            }
                        }
                        Get_Status.staus_messages.Add(folder_name + " file(s) copied..");
                        found = true;
                        folder_ = false;
                    }
                    else
                    {
                        folder_ = false;
                        Get_Status.staus_messages.Add(folder_name + " file(s) not found..");
                        found = false;
                    }
                }
            }
            catch (Exception ex)
            {
                folder_ = false;
                found = false;
                Get_Status.staus_messages.Add("Error searching " + folder_name + " reports - " + ex.Message);
            }
            return found;
        }

        void check_boxes()
        {
            gen_pdf.Checked = folder.Default.covert_pdf;
            column_hide.Checked = excel_settings.Default.unhide_column;
            radioButton2.Checked = excel_settings.Default.trim_3_page;
            new_renew_remove.Checked = excel_settings.Default.new_renew_column;
            warp_text.Checked = excel_settings.Default.wrap_text;
            radioButton1.Checked = excel_settings.Default.trim_2_page;
            verbatim_box.Checked = excel_settings.Default.rm_Verbatim;
            crm_box.Checked = excel_settings.Default.rm_CRM;
            notshown_box.Checked = excel_settings.Default.trim_additionbrdx;
            checkBox1.Checked = excel_settings.Default.rm_Brokerdecision;

            if (excel_settings.Default.trim_3_page == false && excel_settings.Default.trim_2_page == false)
            {
                radioButton3.Checked = true;
            }
            else
            {
                radioButton3.Checked = false;
            }
        
        }

      
        //List<int> folders_list = new List<int>();
        void load_chart()
        {
            chart1.Series.Clear();
            foreach (var item in folders_list)
            {

                Series series = this.chart1.Series.Add(item.Key);

                series.Points.Add(item.Value);
            
            }

        
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            run_sqlquery.RunWorkerAsync();

            c_mode = new Compact_Mode();
            compact_instance.c_mode_1 = c_mode;
            c_mode.Show();
            c_mode.Hide();

            try
            {
                period = DateTime.Now.Year.ToString();
                check_boxes();
                set_theme();
                set_buttons();

                textBox1.Text = report_pack.Default.output_path;

            }
            catch(Exception exc)
            {
                listBox1.Items.Add("Error occurred in loading settings .." + exc.Message);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            generate_directory(folder.Default.folders, textBox1.Text);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            app.Quit();
            Marshal.FinalReleaseComObject(app);

            stopWatch.Stop();
            timer1.Stop();
            gen_paths = "";
            commentary_paths = "";
            double duration = Math.Round(stopWatch.Elapsed.TotalMinutes, 1);
            status.Clear();
            label9.Text = "Process completed. Time taken - " + duration + " minutes";
            stopWatch.Reset();


            if (folder.Default.covert_pdf)
            {
                listBox1.Items.Add("Merging Files..");

                string pdfpath = textBox1.Text + "/" + carrier_name + "/" + carrier_name + " " + period + " " + DateTime.Now.Year + " FINMAR Master Report.pdf";


                if (Functions.get_Functions.MergePDFs(fileNames, pdfpath))
                {
                    File.Copy(pdfpath, textBox1.Text + "/" + carrier_name + "/" + carrier_name +" " + period + " " + DateTime.Now.Year + " FINMAR Master Report_Copy.pdf");
                    Process.Start(pdfpath);
                }
                fileNames.Clear();
            }

            initate_Process.Dispose();
            MessageBox.Show("Report pack created for " + carrier_name, "Operation success", MessageBoxButtons.OK, MessageBoxIcon.Information);

            button7.Enabled = true; 
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            carrier_name = carrier_list.SelectedItem.ToString();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                check_process();

                if (Get_Status.staus_messages.Count > 0)
                {
                    Get_Status.post_log(listBox1);
                }

                if (Get_Status.error_messages.Count > 0)
                {
                    Get_Status.post_error_log(listBox2);
                }
                
            }
            catch
            { }
        }

        private static void FindReplace(string documentLocation, string findText, string replaceText)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            var doc = app.Documents.Open(documentLocation);

            var range = doc.Range();

            range.Find.Execute(FindText: findText, Replace: WdReplace.wdReplaceOne, ReplaceWith: replaceText);

            var shapes = doc.Shapes;

            foreach (Shape shape in shapes)
            {
                if (!shape.Name.Contains("Picture"))
                {
                    
                    var initialText = shape.TextFrame.TextRange.Text;
                    var resultingText = initialText.Replace(findText, replaceText);
                    shape.TextFrame.TextRange.Text = resultingText;
                }
                
            }

            doc.Save();
            doc.Close();

            Marshal.ReleaseComObject(app);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (gen_pdf.Checked)
            {
                folder.Default.covert_pdf = true;
                groupBox14.Enabled = true;
            }
            else
            {
                folder.Default.covert_pdf = false;
                groupBox14.Enabled = false;
            }

            folder.Default.Save();
            folder.Default.Reload();
        }

        private void column_hide_CheckedChanged(object sender, EventArgs e)
        {

            if (column_hide.Checked)
            {
                excel_settings.Default.unhide_column = true;
            }
            else
            {
                excel_settings.Default.unhide_column = false;
            }
            excel_settings.Default.Save();
            excel_settings.Default.Reload();
        }

        private void page_remove_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                excel_settings.Default.trim_3_page = true;
            }
            else
            {
                excel_settings.Default.trim_3_page = false;
            }
            excel_settings.Default.Save();
            excel_settings.Default.Reload();
        }

        private void warp_text_CheckedChanged(object sender, EventArgs e)
        {
            if (warp_text.Checked)
            {
                excel_settings.Default.wrap_text = true;
            }
            else
            {
                excel_settings.Default.wrap_text = false;
            }
            excel_settings.Default.Save();
            excel_settings.Default.Reload();
        }

        private void new_renew_remove_CheckedChanged(object sender, EventArgs e)
        {
            if (new_renew_remove.Checked)
            {
                excel_settings.Default.new_renew_column = true;
            }
            else
            {
                excel_settings.Default.new_renew_column = false;
            }
            excel_settings.Default.Save();
            excel_settings.Default.Reload();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            folder.Default.Save();
            report_pack.Default.Save();
            excel_settings.Default.Save();

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
          
        }

        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            region = comboBox2.Text;
            comboBox3.ResetText();
            carrier_list.ResetText();
            carrier_list.Items.Clear();
            
        }

        private void cover_client_box_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        private void radioButton1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                excel_settings.Default.trim_3_page = false;
                excel_settings.Default.trim_2_page = true;
                excel_settings.Default.Save();
                excel_settings.Default.Reload();
            }
        }

        private void radioButton2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                excel_settings.Default.trim_2_page = false;
                excel_settings.Default.trim_3_page = true;
                excel_settings.Default.Save();
                excel_settings.Default.Reload();
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                excel_settings.Default.trim_2_page = false;
                excel_settings.Default.trim_3_page = false;
                excel_settings.Default.Save();
                excel_settings.Default.Reload();
            }
        }

        private void exportLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StringBuilder log = new StringBuilder();
            foreach (var item in listBox1.Items)
            {
                log.AppendLine(item.ToString());
            }

            if(!Directory.Exists("c://RPH/log"))
            {
                Directory.CreateDirectory("c://RPH/log");
            }

            File.WriteAllText("c://RPH/log/log.txt",log.ToString());

            Process.Start("c://RPH/log/log.txt");
        }

        private void notshown_box_CheckedChanged(object sender, EventArgs e)
        {
            if (notshown_box.Checked)
            {
                excel_settings.Default.trim_additionbrdx = true;
            }
            else
            {
                excel_settings.Default.trim_additionbrdx = false;
            }
            excel_settings.Default.Save();
            excel_settings.Default.Reload();
        }

        private void crm_box_CheckedChanged(object sender, EventArgs e)
        {
            if (crm_box.Checked)
            {
                excel_settings.Default.rm_CRM = true;
            }
            else
            {
                excel_settings.Default.rm_CRM = false;
            }
            excel_settings.Default.Save();
            excel_settings.Default.Reload();
        }

        private void verbatim_box_CheckedChanged(object sender, EventArgs e)
        {
            if (verbatim_box.Checked)
            {
                excel_settings.Default.rm_Verbatim = true;
            }
            else
            {
                excel_settings.Default.rm_Verbatim = false;
            }
            excel_settings.Default.Save();
            excel_settings.Default.Reload();
        }


        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            period = comboBox1.Text;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_q = comboBox3.Text;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (c_mode.Visible == false)
            {
                c_mode.Visible = true;
                compact_instance.is_c_mode = false;
                this.Visible = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (initate_Process.IsBusy)
            {
                if (MessageBox.Show("Report pack creation in progress..\nAbondon current process and exit application ?", "Exit application", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    initate_Process.CancelAsync();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    app.Quit();
                    Marshal.FinalReleaseComObject(app);

                    Directory.Delete(textBox1.Text + "/"+carrier_name,true);
                    this.Close();
                }

            }
            else
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                app.Quit();
                Marshal.FinalReleaseComObject(app);

                this.Close();
            }
            
        }

        private void panel2_MouseDown(object sender, MouseEventArgs e)
        {
            mouseDown = true;
            lastLocation = e.Location;
        }

        private void panel2_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void panel2_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            Random rnd = new Random();
            Color randomColor = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));

            Transition.run(panel2, "BackColor", randomColor, new TransitionType_Linear(5500));
            Transition.run(button7, "BackColor", randomColor, new TransitionType_Linear(5500));
            Transition.run(panel1, "BackColor", randomColor, new TransitionType_Linear(5500));

            Transition.run(c_mode.panel2, "BackColor", randomColor, new TransitionType_Linear(5500));
            Transition.run(label3, "ForeColor", IdealTextColor(randomColor), new TransitionType_Linear(5500));
            Transition.run(c_mode.label3, "ForeColor", IdealTextColor(randomColor), new TransitionType_Linear(5500));
            Transition.run(status_label_, "ForeColor", IdealTextColor(randomColor), new TransitionType_Linear(5500));
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (var td = new TaskDialog())
            {
                var cancelButton = new TaskDialogButton(ButtonType.Cancel);
                var iuButton = new TaskDialogButton(ButtonType.Custom);
                td.WindowTitle = "Advanced settings";
                iuButton.Text = "I understand";
                td.Buttons.Add(cancelButton);
                td.Buttons.Add(iuButton);
                td.MainIcon = TaskDialogIcon.Warning;
                td.MainInstruction = "Advanced settings";
                td.Content = "You are trying to access advanced settings.\n\nAdvanced settings directly or indirectly control core functions of this tool and may impact it if changed.\nThese settings should be changed by advanced users only.\n";
                td.ExpandedInformation = "Incorrect settings or parameters set on the following page may result in Slow-downs, Application crashes and Undesired outputs.";

                TaskDialogButton button = td.ShowDialog(this);
                if (button == iuButton)
                {
                    Advanced_Settings adv = new Advanced_Settings();
                    adv.ShowDialog(this);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new About().ShowDialog(this);
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            mouseDown = true;
            lastLocation = e.Location;
        }

        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            carrier_list.Items.Clear();
            using (SqlConnection connection = new SqlConnection("Server=gbips-i-db700;Database=FINMAR_Placement;Integrated Security=True;"))
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand("select distinct InsurerReportingName from [dbo].[InsurerSubscription];",connection);
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        carrier_list.Items.Add(rdr.GetString(0).Trim());
                    }
                }
            }
        }

        void set_theme()
        {
            change_color(color_settings.Default.colorscheme);
        }

        void reset_colors()
        {
            panel1.BackColor = SystemColors.Control;
            panel2.BackColor = SystemColors.Control;
            pictureBox1.BackColor = SystemColors.Control;
            c_mode.pictureBox1.BackColor = SystemColors.Control;

            button1.BackColor = Color.Transparent;
            button3.BackColor = Color.Transparent;
            button4.BackColor = Color.Transparent;
            button5.BackColor = Color.Transparent;
            button6.BackColor = Color.Transparent;
            button8.BackColor = Color.Transparent;

            c_mode.panel2.BackColor = SystemColors.Control;
            c_mode.button5.BackColor = SystemColors.Control;
            c_mode.pictureBox1.BackColor = SystemColors.Control;

            button7.ForeColor = IdealTextColor(SystemColors.Control);
            c_mode.label3.ForeColor = IdealTextColor(SystemColors.Control);
            label3.ForeColor = IdealTextColor(SystemColors.Control);
            status_label_.ForeColor = IdealTextColor(SystemColors.Control);
        
        }

        void change_color(Color color)
        {
            timer2.Stop();
            reset_colors();
            pictureBox2.Visible = false;
            if (color_settings.Default.theme == "Accent" || color_settings.Default.theme == "Random")
            {

                panel2.BackColor = color;
                panel1.BackColor = color;
                pictureBox1.BackColor = Color.Transparent;
                c_mode.panel2.BackColor = color;


                button1.BackColor = Color.Transparent;
                button3.BackColor = Color.Transparent;
                button4.BackColor = Color.Transparent;
                button5.BackColor = Color.Transparent;
                button6.BackColor = Color.Transparent;
                button8.BackColor = Color.Transparent;
                button7.BackColor = color;
                button7.ForeColor = IdealTextColor(color);

                c_mode.label3.ForeColor = IdealTextColor(color);
                label3.ForeColor = IdealTextColor(color);
                status_label_.ForeColor = IdealTextColor(color);
            }
            else if (color_settings.Default.theme=="Lava lamp")
            {
                
                timer2.Start();
            }

            else if (color_settings.Default.theme == "WTW")
            {
                panel1.BackColor = SystemColors.ControlLight;
                panel2.BackColor = Color.FromArgb(111,42,129);
                pictureBox1.BackColor = SystemColors.ControlLight;
                c_mode.pictureBox1.BackColor = SystemColors.ControlLight;

                button1.BackColor = SystemColors.ControlLight;
                button3.BackColor = SystemColors.ControlLight;
                button4.BackColor = SystemColors.ControlLight;
                button5.BackColor = SystemColors.ControlLight;
                button6.BackColor = SystemColors.ControlLight;
                button7.BackColor = Color.FromArgb(111, 42, 129);
                button8.BackColor = SystemColors.ControlLight;
                c_mode.panel2.BackColor = Color.FromArgb(111, 42, 129);
                c_mode.button5.BackColor = SystemColors.ControlLight;
                c_mode.pictureBox1.BackColor = SystemColors.ControlLight;
                pictureBox2.Visible = true;

                button7.ForeColor = IdealTextColor(Color.FromArgb(111, 42, 129));
                c_mode.label3.ForeColor = IdealTextColor(Color.FromArgb(111, 42, 129));
                label3.ForeColor = IdealTextColor(Color.FromArgb(111, 42, 129));

            }

            else if (color_settings.Default.theme == "Default")
            {
               
                panel1.BackColor = Color.LightSlateGray;
                panel2.BackColor = SystemColors.Control;
                pictureBox1.BackColor = Color.LightSlateGray;
                c_mode.pictureBox1.BackColor = Color.LightSlateGray;
                c_mode.panel2.BackColor = SystemColors.Control;
                button1.BackColor = Color.LightSlateGray;
                button3.BackColor = Color.LightSlateGray;
                button4.BackColor = Color.LightSlateGray;
                button5.BackColor = Color.LightSlateGray;
                button6.BackColor = Color.LightSlateGray;
                button7.BackColor = Color.LightSlateGray;
                button8.BackColor = Color.LightSlateGray;

                c_mode.button5.BackColor = Color.LightSlateGray;
                c_mode.pictureBox1.BackColor = Color.LightSlateGray;

                button7.ForeColor = Color.Black;
                c_mode.label3.ForeColor = Color.Black;
                label3.ForeColor = Color.Black;

            }

        }

        public Color IdealTextColor(Color bg)
        {
            int nThreshold = 105;
            int bgDelta = Convert.ToInt32((bg.R * 0.299) + (bg.G * 0.587) +
                                          (bg.B * 0.114));

            Color foreColor = (255 - bgDelta < nThreshold) ? Color.Black : Color.White;
            
            return foreColor;
        }

        void change_image()
        {
               Bitmap bmp = new Bitmap(button8.Image);

            //load image in picturebox1
           

            //get image dimension
            int width = bmp.Width;
            int height = bmp.Height;

            //negative
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    //get pixel value
                    Color p = bmp.GetPixel(x, y);

                    //extract ARGB value from p
                    int a = p.A;
                    int r = p.R;
                    int g = p.G;
                    int b = p.B;

                    //find negative value
                    r = 255 - r;
                    g = 255 - g;
                    b = 255 - b;

                    //set new ARGB value in pixel
                    bmp.SetPixel(x, y, Color.FromArgb(a, r, g, b));
                }
            }

            button8.Image = bmp;
        
        }

        private void comboBox3_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            tier = Convert.ToInt32(comboBox3.Text);
            carrier_list.ResetText();
            carrier_list.Items.Clear();

            if (tier == 3)
            {
                comboBox4.Enabled = false;
                comboBox4.ResetText();
                m_q = "";
            }
            else
            {
                comboBox4.Enabled = true;
            }

            if (!get_carriers.IsBusy)
            {
                get_carriers.RunWorkerAsync();
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            Basic_Settings bs = new Basic_Settings();
            bs.ShowDialog(this);
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_q = comboBox4.Text;
        }

        private void run_sqlquery_DoWork(object sender, DoWorkEventArgs e)
        {
            status_lbl_text = "Loading BUs..";
            BUs.Clear();
            string command = "select Distinct [Business Unit] from [dbo].[InsurerSubscription] ins " +
                               "left outer join [FINMAR_ODS].[dbo].[vw_DataSourceInstanceToBUMap] BUMap " +
                                "on ins.AssignedToBUHierarchyId=BUMap.BuHierarchyId " +
                                "where EndDate is NULL and InsurerReportingName<> 'Test' and IsDeleted=0;";


            BUs = Functions.get_Functions.connect_ToDatabase(connectionString, command);
        }

        private void run_sqlquery_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
            foreach (string BU in BUs)
                comboBox2.Items.Add(BU);
            status_lbl_text = "BUs Loaded";
        }

        private void get_carriers_DoWork(object sender, DoWorkEventArgs e)
        {
            status_lbl_text = "Loading carriers.."; 
            Carriers.Clear();
            
            string command = "select Distinct InsurerReportingName from [dbo].[InsurerSubscription] ins " +
                             "left outer join [FINMAR_ODS].[dbo].[vw_DataSourceInstanceToBUMap] BUMap " +
                              "on ins.AssignedToBUHierarchyId=BUMap.BuHierarchyId " +
                              "where EndDate is NULL and InsurerReportingName<> 'Test' and IsDeleted=0 and [Business Unit] = '" + region + "' and SubscriptionTier =" + tier + ";";


            Carriers = Functions.get_Functions.connect_ToDatabase(connectionString, command);
        }

        private void get_carriers_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            foreach (string Carrier in Carriers)
                carrier_list.Items.Add(Carrier);
            status_lbl_text = "Carriers Loaded.";
        }

        private void status_monitor_Tick(object sender, EventArgs e)
        {
            status_label_.Text = status_lbl_text;

            if (color_settings.Default.ThemeChanged)
            {
                set_theme();
                color_settings.Default.ThemeChanged = false;
                
            }

            load_chart();
        }


        private void button9_Click(object sender, EventArgs e)
        {

           
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                excel_settings.Default.rm_Brokerdecision = true;
            }
            else
            {
                excel_settings.Default.rm_Brokerdecision = false;
            }
            excel_settings.Default.Save();
            excel_settings.Default.Reload();
        }

        private void com_pdf_CheckedChanged(object sender, EventArgs e)
        {
            coms_pdf = com_pdf.Checked;
        }

        private void pip_pdf_CheckedChanged(object sender, EventArgs e)
        {
            pipe_pdf = pip_pdf.Checked;
        }

        private void a_pdf_CheckedChanged(object sender, EventArgs e)
        {
            analysis_pdf = a_pdf.Checked;
        }

        private void bord_pdf_CheckedChanged(object sender, EventArgs e)
        {
            bordx_pdf = bord_pdf.Checked;
        }

        private void coc_box_CheckedChanged(object sender, EventArgs e)
        {
            coc = coc_box.Checked;
        }

        private void comm_box_CheckedChanged(object sender, EventArgs e)
        {
            commentary = comm_box.Checked;
        }

        private void pip_box_CheckedChanged(object sender, EventArgs e)
        {
            pipline = pip_box.Checked;

            if (!pipline)
            {
                groupBox5.Enabled = false;
            }
            else
            {
                groupBox5.Enabled = true;
            }
        }

        private void ana_box_CheckedChanged(object sender, EventArgs e)
        {
            analysis = ana_box.Checked;

            if (!analysis)
            {
                groupBox6.Enabled = false;
            }
            else
            {
                groupBox6.Enabled = true;
            }

        }

        private void brdx_box_CheckedChanged(object sender, EventArgs e)
        {
            bordx = brdx_box.Checked;

            if (!bordx)
            {
                groupBox4.Enabled = false;
                groupBox10.Enabled = false;
            }
            else
            {
                groupBox4.Enabled = true;
                groupBox10.Enabled = true;
            }

        }

        private void ar_box_CheckedChanged(object sender, EventArgs e)
        {
            add_requirements = ar_box.Checked;

            if (!add_requirements)
            {
                groupBox11.Enabled = false;
            }
            else
            {
                groupBox11.Enabled = true;
            }
        }

        private void dss_box_CheckedChanged(object sender, EventArgs e)
        {
            dss = dss_box.Checked;
        }

        private void exportAsImageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (chart1.Series.Count() > 0)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    chart1.SaveImage(saveFileDialog1.FileName,ChartImageFormat.Png);
                } 
            
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (Directory.Exists(textBox1.Text) && textBox1.Text!="")
            {
                button7.Enabled = true;
            }
            else
            {
                button7.Enabled = false;
            }
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

    }

   
}
