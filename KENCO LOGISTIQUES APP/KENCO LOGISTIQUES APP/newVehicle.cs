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
    public partial class newVehicle : Form
    {
        public newVehicle()
        {
            InitializeComponent();
            int newX = 10;
            int newY = 10;

            // Set the new location of the form
            this.Location = new System.Drawing.Point(newX, newY);
        }

        private void MainMenu_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.ShowDialog();
            this.Close();
        }
    }
}
