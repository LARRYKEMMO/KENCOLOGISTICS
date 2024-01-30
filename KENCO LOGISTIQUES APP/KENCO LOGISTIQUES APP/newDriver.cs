using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KENCO_LOGISTIQUES_APP
{
    public partial class newDriver : Form
    {
        public newDriver()
        {
            InitializeComponent();
            OpenXLFile();
            
        }

        private void MainMenu_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.ShowDialog();
            this.Close();
        }

        private void AddNew_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(FirstNameBox.Text, LastNameBox.Text, DOBString + MOBString + YOBString, POBBox.Text, IDBox.Text, IDEDateString + IDEMonthString + IDEYearString, DriverLicencseBox.Text, LicenseCategory.Text, LicenseDateString + LicenseMonthString + LicenseYearString, AddressBox.Text, TelNumberBox.Text, EmailBox.Text);
        }

        private void DOBBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            DOBString = DOBBox.Text;
        }

        private void MOBBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            MOBString = MOBBox.Text;
        }

        private void YOBBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            YOBString = YOBBox.Text;
        }

        private void VehicleList_SelectedIndexChanged(object sender, EventArgs e)
        {
            AssignedVehicleString = VehicleList.Text;
        }

        private void IDEDateBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            IDEDateString = IDEDateBox.Text;
        }

        private void IDEMonthBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            IDEMonthString = IDEMonthBox.Text;
        }

        private void IDEYearBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            IDEYearString = IDEYearBox.Text;
        }

        private void LicenseDateBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            LicenseDateString = LicenseDateBox.Text;
        }

        private void LicenseMonthBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            LicenseMonthString = LicenseMonthBox.Text;
        }

        private void LicenseYearBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            LicenseYearString = LicenseYearBox.Text;
        }

        private void Delete_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
            }
        }

        private void Reset_Click(object sender, EventArgs e)
        {
            // Clears Input
            FirstNameBox.Text = null;
            LastNameBox.Text = null;
            DOBBox.Text = null;
            MOBBox.Text = null; 
            YOBBox.Text = null;
            POBBox.Text = null; 
            IDBox.Text = null;
            IDEDateBox.Text = null;
            IDEMonthBox.Text = null;
            IDEYearBox.Text = null;
            DriverLicencseBox.Text = null;
            LicenseCategory.Text = null;
            LicenseDateBox.Text = null;
            LicenseMonthBox.Text = null;
            LicenseYearBox.Text = null;
            AddressBox.Text = null;
            TelNumberBox.Text = null;
            EmailBox.Text = null;
            VehicleList.Text = null;


            // Clears Data Grid View
            int numRows = dataGridView1.Rows.Count;
            for(int i = 1; i < numRows; i++)
            {
                try
                {
                    int max = dataGridView1.Rows.Count - 1;
                    dataGridView1.Rows.Remove(dataGridView1.Rows[max]);
                }
                catch(Exception exe)
                {
                    MessageBox.Show("All rows are to be deleted " + exe, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if(i == numRows - 1)
                {
                    dataGridView1.Rows.Remove(dataGridView1.Rows[0]);
                }
            }
        }

      

        private void OpenXLFile()
        {
            dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
            dynamic xlworksheet = xlworkbook.Worksheets["Sheet1"];
            dynamic xlrange = xlworksheet.UsedRange;

            dataGridView1.ColumnCount = xlrange.Columns.Count;

            for (int xlrow = 2; xlrow <= xlrange.Rows.Count; xlrow++)
            {

                dataGridView1.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text,
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text,
                xlrange.Cells[xlrow, 6].Text, xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text,
                xlrange.Cells[xlrow, 9].Text, xlrange.Cells[xlrow, 10].Text, xlrange.Cells[xlrow, 11].Text,
                xlrange.Cells[xlrow, 12].Text);

            }

            xlworkbook.Close();
            xlapp.Quit();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            //dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            //dynamic xlworkbook = xlapp.Workbooks.Open(excelFilePath);
            //dynamic xlworksheet = xlworkbook.Worksheets["Sheet1"];
            //dynamic xlrange = xlworksheet.UsedRange;

            //// Clear existing data in the Excel worksheet
            //xlrange.Clear();

            //// Write data from the DataGridView to the Excel worksheet
            //for (int row = 0; row < dataGridView1.Rows.Count; row++)
            //{
            //    for (int col = 0; col < dataGridView1.Columns.Count; col++)
            //    {
            //        xlworksheet.Cells[row + 1, col + 1].Value = dataGridView1.Rows[row].Cells[col].Value;
            //    }
            //}

            //// Save changes to the Excel file
            //xlworkbook.Save();

            //// Close workbook and quit Excel
            ////xlworkbook.Close();
            ////xlapp.Quit();
            ///

            if (File.Exists(newFilePath))
            {
                // Open Excel and the workbook
                dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
                dynamic xlworksheet = xlworkbook.Worksheets["Sheet1"];

                // Clear existing data in the Excel worksheet
                xlworksheet.Cells.ClearContents();

                for (int col = 1; col <= dataGridView1.ColumnCount; col++)
                {
                    xlworksheet.Cells[1, col].Value = dataGridView1.Columns[col - 1].HeaderText;
                }

                // Write data from the DataGridView to the Excel worksheet
                for (int row = 0; row < dataGridView1.Rows.Count; row++)
                {
                    for (int col = 0; col < dataGridView1.Columns.Count; col++)
                    {
                        xlworksheet.Cells[row + 2, col + 1].Value = dataGridView1.Rows[row].Cells[col].Value;
                    }
                }

                // Save changes to the Excel file
                xlworkbook.Save();

                // Close workbook and quit Excel
                xlworkbook.Close(true);
                xlapp.Quit();

                MessageBox.Show("Data saved successfully to Excel file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("The specified Excel file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
