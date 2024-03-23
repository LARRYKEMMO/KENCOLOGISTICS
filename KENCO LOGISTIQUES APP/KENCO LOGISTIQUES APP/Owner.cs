using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KENCO_LOGISTIQUES_APP
{
    public partial class Owner : Form
    {
        public Owner()
        {
            InitializeComponent();
            OpenXLFile();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void MainMenu_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.ShowDialog();
            this.Close();
        }

        private void OpenXLFile()
        {
            dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
            dynamic xlworksheet = xlworkbook.Worksheets["Owner"];
            dynamic xlrange = xlworksheet.UsedRange;

            dataGridView1.ColumnCount = xlrange.Columns.Count;

            for (int xlrow = 2; xlrow <= xlrange.Rows.Count; xlrow++)
            {

                dataGridView1.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text,
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text,
                xlrange.Cells[xlrow, 6].Text, xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text,
                xlrange.Cells[xlrow, 9].Text, xlrange.Cells[xlrow, 10].Text);

            }

            DeleteEmptyRows(dataGridView1);

            xlworkbook.Close();
            xlapp.Quit();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (File.Exists(newFilePath))
            {
                // Open Excel and the workbook
                dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
                dynamic xlworksheet = xlworkbook.Worksheets["Owner"];

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

        private void AddNew_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(FirstNameBox.Text) ||
                string.IsNullOrEmpty(LastNameBox.Text) ||
                string.IsNullOrEmpty(EmailBox.Text) ||
                string.IsNullOrEmpty(PhoneNumberBox.Text) ||
                string.IsNullOrEmpty(AddressBox.Text) ||
                string.IsNullOrEmpty(BANameBox.Text) ||
                string.IsNullOrEmpty(BANumberBox.Text))
            {
                //MessageBox.Show("All fields must be filled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowToast("ERROR", "All fields must be filled.");
            }
            else
            {
                if(IntegerValidity(PhoneNumberBox.Text).Equals(true) && IntegerValidity(BANumberBox.Text).Equals(true))
                {
                    dataGridView1.Rows.Add(FirstNameBox.Text, LastNameBox.Text, EmailBox.Text, AddressBox.Text, PhoneNumberBox.Text, BANameBox.Text, BANumberBox.Text);
                }
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

        private void Delete_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
            }
        }

        private void Reset_Click(object sender, EventArgs e)
        {
            FirstNameBox = null;
            LastNameBox = null;
            EmailBox = null;
            PhoneNumberBox = null;
            AddressBox = null;
            BANameBox = null;
            BANumberBox = null;

            // Clears Data Grid View
            int numRows = dataGridView1.Rows.Count;
            for (int i = 0; i < numRows; i++)
            {
                try
                {
                    int max = dataGridView1.Rows.Count - 1;
                    dataGridView1.Rows.Remove(dataGridView1.Rows[max]);
                }
                catch (Exception exe)
                {
                    //MessageBox.Show("All rows are to be deleted " + exe, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ShowToast("ERROR", "Rows are not able to be deleted.");
                }
            }
        }

        private void EmailBox_TextChanged(object sender, EventArgs e)
        {

        }

        private bool IntegerValidity(string TempText)
        {
            int number;
            bool isInteger;

            try
            {
                number = int.Parse(TempText);
                isInteger = true;

            }
            catch (Exception e)
            {
                ShowToast("ERROR", "There seems to be a problem with your input (" + TempText + ")");
                isInteger = false;
            }

            return isInteger;
        }

        private bool CarPlateValidity(string TempText)
        {
            string pattern = @"^[A-Z]{2}\d{3}[A-Z]{2}$";
            bool isMatch = Regex.IsMatch(TempText, pattern);

            if (!isMatch)
            {
                ShowToast("ERROR", "The Vehicle Plate Number is not in the format AB123CD.");
            }

            return isMatch;
        }
    
    }
}
