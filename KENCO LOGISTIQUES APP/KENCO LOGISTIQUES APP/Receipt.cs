﻿using System;
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
    public partial class Receipt : Form
    {
        public Receipt()
        {
            InitializeComponent();
            lblDateToday.Text = DateTime.Now.ToLongDateString();
        }

        

        private void Receipt_Load(object sender, EventArgs e)
        {

        }
    }
}
