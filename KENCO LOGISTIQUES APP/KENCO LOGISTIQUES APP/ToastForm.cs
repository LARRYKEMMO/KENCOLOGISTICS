using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KENCO_LOGISTIQUES_APP
{
    public partial class ToastForm : Form
    {
        public ToastForm(string type, string message)
        {
            InitializeComponent();
            lblType.Text = type;
            lblMessage.Text = message;

            switch (type)
            {
                case "SUCCESS":
                    ImageBox.Image = Properties.Resources.green_check_mark_clipart_8;
                    panel1.BackColor = Color.FromArgb(0, 192, 0);
                    break;

                case "ERROR":
                    ImageBox.Image = Properties.Resources.failed_icon_7;
                    panel1.BackColor = Color.Red;
                    break;

                case "INFO":
                    ImageBox.Image = Properties.Resources.image_info;
                    panel1.BackColor = Color.Blue;
                    break;

                case "WARNING":
                    ImageBox.Image = Properties.Resources.Warning;
                    panel1.BackColor = Color.FromArgb(255, 128, 0);
                    break;

                case "REMINDER":
                    ImageBox.Image = Properties.Resources.failed_icon_7;
                    panel1.BackColor = Color.Red;
                    break;
            }
            
        }

        private void Position()
        {
            int ScreenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int ScreenHeight = Screen.PrimaryScreen.WorkingArea.Height;

            toastX = ScreenWidth - this.Width;
            toastY = ScreenHeight - this.Height;

            this.Location = new Point(toastX, toastY);
        }

        private void ToastForm_Load(object sender, EventArgs e)
        {
            Position();
        }

        private void toastTimer_Tick(object sender, EventArgs e)
        {
            toastY -= 10;
            this.Location = new Point(toastX, toastY);
            if(toastY <= 1000)
            {
                toastTimer.Stop();
                toastHide.Start();
            }
        }
        int y = 300;
        private void toastHide_Tick(object sender, EventArgs e)
        {
            y--;
            if(y <= 0)
            {
                toastY += 1;
                this.Location = new Point(toastX, toastY += 10);
                if(toastY > 1040)
                {
                    toastHide.Stop();
                    y = 100;
                    this.Close();
                }
            }
        }
    }

}
