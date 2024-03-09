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
            if (string.IsNullOrEmpty(FirstNameBox.Text) ||
                string.IsNullOrEmpty(LastNameBox.Text) ||
                string.IsNullOrEmpty(DOBString) ||
                string.IsNullOrEmpty(MOBString) ||
                string.IsNullOrEmpty(YOBString) ||
                string.IsNullOrEmpty(POBBox.Text) ||
                string.IsNullOrEmpty(IDBox.Text) ||
                string.IsNullOrEmpty(IDEDateString) ||
                string.IsNullOrEmpty(IDEMonthString) ||
                string.IsNullOrEmpty(IDEYearString) ||
                string.IsNullOrEmpty(DriverLicencseBox.Text) ||
                string.IsNullOrEmpty(LicenseCategory.Text) ||
                string.IsNullOrEmpty(LicenseDateString) ||
                string.IsNullOrEmpty(LicenseMonthString) ||
                string.IsNullOrEmpty(LicenseYearString) ||
                string.IsNullOrEmpty(AddressBox.Text) ||
                string.IsNullOrEmpty(TelNumberBox.Text) ||
                string.IsNullOrEmpty(EmailBox.Text))
            {
                //MessageBox.Show("All fields must be filled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowToast("ERROR", "All fields must be filled.");
            }
            else
            {
                dataGridView1.Rows.Add(FirstNameBox.Text, LastNameBox.Text, DOBString + Slash + MOBString + Slash + YOBString, POBBox.Text, IDBox.Text, IDEDateString + Slash + IDEMonthString + Slash + IDEYearString, DriverLicencseBox.Text, LicenseCategory.Text, LicenseDateString + Slash + LicenseMonthString + Slash + LicenseYearString, AddressBox.Text, TelNumberBox.Text, EmailBox.Text);

            }

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
                    //MessageBox.Show("All rows are to be deleted " + exe, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ShowToast("ERROR", "Rows are not able to be deleted");
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
            dynamic xlworksheet = xlworkbook.Worksheets["NewEmployee"];
            dynamic xlrange = xlworksheet.UsedRange;

            //dataGridView1.ColumnCount = xlrange.Columns.Count;
            dataGridView1.Rows.Clear();

            for (int xlrow = 2; xlrow <= xlrange.Rows.Count; xlrow++)
            {

                dataGridView1.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text,
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text,
                xlrange.Cells[xlrow, 6].Text, xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text,
                xlrange.Cells[xlrow, 9].Text, xlrange.Cells[xlrow, 10].Text, xlrange.Cells[xlrow, 11].Text,
                xlrange.Cells[xlrow, 12].Text);

            }

            DeleteEmptyRows(dataGridView1);

            xlworkbook.Close();
            xlapp.Quit();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            if (File.Exists(newFilePath))
            {
                // Open Excel and the workbook
                dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
                dynamic xlworksheet = xlworkbook.Worksheets["NewEmployee"];

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

                //MessageBox.Show("Data saved successfully to Excel file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ShowToast("SUCCESS", "Data saved successfully to Excel file.");
            }
            else
            {
                //MessageBox.Show("The specified Excel file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowToast("ERROR", "The specified Excel file does not exist.");
            }
        }

        private void DeleteEmptyRows(DataGridView dataGridView)
        {
            // Iterate through the rows in reverse order to avoid issues with indices
            for (int i = dataGridView.Rows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = dataGridView.Rows[i];
                if (IsRowEmpty(row))
                {
                    dataGridView.Rows.RemoveAt(i);
                }
            }
        }

        private bool IsRowEmpty(DataGridViewRow row)
        {
            foreach (DataGridViewCell cell in row.Cells)
            {
                if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                {
                    return false; // At least one cell has a non-null, non-empty value
                }
            }
            return true; // All cells are either null or empty
        }

        public void ShowToast(string Type, string Message)
        {
            ToastForm toastForm = new ToastForm(Type, Message);
            toastForm.ShowDialog();
        }
    }
}
