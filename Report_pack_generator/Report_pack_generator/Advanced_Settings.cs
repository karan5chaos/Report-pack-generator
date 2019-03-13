using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Transitions;
using System.Globalization;
using Ookii.Dialogs;

namespace Report_pack_generator
{
    public partial class Advanced_Settings : Form
    {
        public Advanced_Settings()
        {
            InitializeComponent();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void Advanced_settings_Load(object sender, EventArgs e)
        {

            

            load_folder_settings();

               // Transition.run(checkBox1, "ForeColor", Color.Firebrick, new TransitionType_Flash(9999, 4000));

        }

        void load_folder_settings()
        {
            listBox2.Items.Clear();
            foreach (var folder in Settings.folder.Default.folders)
            {
                listBox2.Items.Insert(0, folder);
            }

        }




        private void textBox2_TextChanged(object sender, EventArgs e)
        {
           


        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                Transition.run(checkBox1, "ForeColor", Color.Firebrick, new TransitionType_Flash(9999, 2500));
               
                groupBox6.Enabled = true;
                groupBox7.Enabled = true;

            }
            else
            {
                Transition.run(checkBox1, "ForeColor", SystemColors.ControlText, new TransitionType_Deceleration(2500));
                groupBox6.Enabled = false;
                groupBox7.Enabled = false;
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                Transition.run(checkBox2, "ForeColor", Color.Firebrick, new TransitionType_Flash(9999, 2500));

                panel1.Enabled = true;
               // groupBox7.Enabled = true;

            }
            else
            {
                Transition.run(checkBox2, "ForeColor", SystemColors.ControlText, new TransitionType_Deceleration(2500));
                panel1.Enabled = false;
                //groupBox7.Enabled = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            

            
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void groupBox8_Enter(object sender, EventArgs e)
        {

        }
    }
}
