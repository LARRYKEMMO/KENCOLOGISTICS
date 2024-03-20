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
            reminder.reminderCheck();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }


        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to exit?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Close the application
                Application.Exit();
            }
        }

        private void ownersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Owner form3 = new Owner();
            form3.ShowDialog();
            this.Close();
        }

        private void vehiclesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            newVehicle newVehicles = new newVehicle();
            newVehicles.ShowDialog();
            this.Close();
        }

        private void employeesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            newDriver NewDriver = new newDriver();
            NewDriver.ShowDialog();
            this.Close();
        }

        private void remindersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Accounting accounting = new Accounting();
            accounting.ShowDialog();
            this.Close();
        }

        private void accountingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void incomeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void expensesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void analyticsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void cashFLowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void actionsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void incomeToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            IncomeAnalytics IA = new IncomeAnalytics();
            IA.ShowDialog();
            this.Close();
        }

        private void expensesToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            ExpenseAnalyticscs EA = new ExpenseAnalyticscs();
            EA.ShowDialog();
            this.Close();
        }

        private void cashFlowToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            CashFlowAnalytics CFA = new CashFlowAnalytics();
            CFA.ShowDialog();
            this.Close();
        }

        private void remindersToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Reminder reminder = new Reminder();
            reminder.ShowDialog();
            this.Close();
        }

        private void invoicesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Invoice invoice = new Invoice();
            invoice.ShowDialog();
            this.Close();
        }

        private void accidentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Accidents accident = new Accidents();
            accident.ShowDialog();
            this.Close();
        }
    }
}
