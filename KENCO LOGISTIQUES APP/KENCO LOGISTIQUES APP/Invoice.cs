﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KENCO_LOGISTIQUES_APP
{
    public partial class Invoice : Form
    {
        public Invoice()
        {
            InitializeComponent();
            lblDateToday.Text = DateTime.Now.ToLongDateString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoadForm(new InvoiceCopy());
        }

        private void btnMenu_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.ShowDialog();
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Print(this.panelPrint);
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Print(Panel pnl)
        {
            PrinterSettings ps = new PrinterSettings();
            panelPrint = pnl;
            getprintarea(pnl);
            printPreviewDialog1.Document = printDocument1;
            printDocument1.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);
            printPreviewDialog1.ShowDialog();
        }

        private Bitmap memorying;

        private void getprintarea(Panel pnl)
        {
            memorying = new Bitmap(pnl.Width, pnl.Height);
            pnl.DrawToBitmap(memorying, new Rectangle(0, 0, pnl.Width, pnl.Height));
        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(btnPrint, "Print");
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            Rectangle pagearea = e.PageBounds;
            e.Graphics.DrawImage(memorying, pagearea);
        }

        public void LoadForm(object form)
        {
            if (this.panelPrint.Controls.Count > 0)
            {
                this.panelPrint.Controls.Clear();
            }
            Form f = form as Form;
            f.TopLevel = false;
            f.Dock = DockStyle.Fill;
            this.panelPrint.Controls.Add(f);
            this.panelPrint.Tag = f;
            f.Show();
        }

        private void btnReceipts_Click(object sender, EventArgs e)
        {
            LoadForm(new Receipt());
        }
    }
}
