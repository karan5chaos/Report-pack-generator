using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Transitions;
using Report_pack_generator.Settings;


namespace Report_pack_generator
{
    public partial class Basic_Settings : Form
    {
        public Basic_Settings()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
               var color = colorDialog1.Color;
               button1.BackColor = color;
                color_settings.Default.colorscheme = color;
                color_settings.Default.ThemeChanged = true;
                color_settings.Default.Save();
                color_settings.Default.Reload();
                
            } 
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

      


        private void timer1_Tick(object sender, EventArgs e)
        {
        }

        public static void SetDoubleBuffering(System.Windows.Forms.Control control, bool value)
        {
            System.Reflection.PropertyInfo controlProperty = typeof(System.Windows.Forms.Control)
                .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            controlProperty.SetValue(control, value, null);
        }

        bool onload = false;
        private void basic_settings_Load(object sender, EventArgs e)
        {
            onload = true;
            

            if (color_settings.Default.theme == "Accent")
            {
                button1.Enabled = true;
                button1.BackColor = color_settings.Default.colorscheme;
            }
            comboBox1.Text = color_settings.Default.theme;
        }

        void load_color_settings()
        {
            comboBox1.SelectedText = color_settings.Default.theme;
            button1.BackColor = color_settings.Default.colorscheme;
 
        
        }



        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!onload)
            {
                switch (comboBox1.Text.ToString())
                {
                    case "Default":
                        {
                            color_settings.Default.theme = "Default";

                        }
                        break;

                    case "Lava lamp":
                        {
                            color_settings.Default.theme = "Lava lamp";
                        }
                        break;

                    case "Random":
                        {
                            Random rnd = new Random();
                            Color randomColor = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                            color_settings.Default.theme = "Random";
                            color_settings.Default.colorscheme = randomColor;

                        }
                        break;

                    case "Accent":
                        {
                            color_settings.Default.theme = "Accent";
                            button1.Enabled = true;
                            button1.BackColor = color_settings.Default.colorscheme;

                        }
                        break;

                    case "WTW":
                        {
                            color_settings.Default.theme = "WTW";

                        }
                        break;

                    default:
                        {

                        }
                        break;


                }
                color_settings.Default.ThemeChanged = true;
                color_settings.Default.Save();
                color_settings.Default.Reload();
            }
            onload = false;

        }
    }
}
