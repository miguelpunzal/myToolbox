using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    
    public partial class RALLY_NOTIFYER : Form
    {
        public static String textnotification { get; set; }
    public RALLY_NOTIFYER()
        {
            InitializeComponent();
            label2.Text = textnotification;
            Rectangle workingArea = Screen.GetWorkingArea(this);
            this.Location = new Point(workingArea.Right - Size.Width,workingArea.Bottom - Size.Height);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void RALLY_NOTIFYER_Load(object sender, EventArgs e)
        {
            textBox1.Text = label2.Text;
        }
    }
}
