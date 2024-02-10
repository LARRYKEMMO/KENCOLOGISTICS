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
    public partial class Menu : Form
    {
        public Menu()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Owner form3 = new Owner();
            form3.ShowDialog();
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            this.Hide();
            newVehicle newVehicles = new newVehicle();
            newVehicles.ShowDialog();
            this.Close();

            //this.Hide();

        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            newDriver NewDriver = new newDriver();
            NewDriver.ShowDialog();
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            Reminder reminder = new Reminder();
            reminder.ShowDialog();
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Accounting accounting = new Accounting();
            accounting.ShowDialog();
            this.Close();
        }
    }
}
