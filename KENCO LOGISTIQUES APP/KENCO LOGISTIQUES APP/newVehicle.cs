using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KENCO_LOGISTIQUES_APP
{
    public partial class newVehicle : Form
    {
        public newVehicle()
        {
            InitializeComponent();
            //int newX = 10;
            //int newY = 10;

            //// Set the new location of the form
            //this.Location = new System.Drawing.Point(newX, newY);

            OpenXLFile();
        }

        private void MainMenu_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.ShowDialog();
            this.Close();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void MakeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            MakeBoxString = MakeBox.Text;
        }

        private void TypeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            TypeBoxString = TypeBox.Text;
        }

        private void FUDateBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FUDateBoxString = FUDateBox.Text;
        }

        private void FUMonthBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FUMonthBoxString = FUMonthBox.Text;
        }

        private void FUYearBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FUYearBoxString = FUYearBox.Text;
        }

        private void CapacityUnityBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            CapacityUnityBoxString = CapacityUnityBox.Text;
        }

        private void AddNew_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(PlateNumberBox.Text) ||
                string.IsNullOrEmpty(MakeBoxString) ||
                string.IsNullOrEmpty(TypeBoxString) ||
                string.IsNullOrEmpty(FUDateBoxString) ||
                string.IsNullOrEmpty(FUMonthBoxString) ||
                string.IsNullOrEmpty(FUYearBoxString) ||
                string.IsNullOrEmpty(CapacityBox.Text) ||
                string.IsNullOrEmpty(CapacityUnityBoxString) ||
                string.IsNullOrEmpty(ChassisNumberBox.Text) ||
                string.IsNullOrEmpty(DriverNameBox.Text) ||
                string.IsNullOrEmpty(CDBox.Text) ||
                string.IsNullOrEmpty(DTelephoneBox.Text) ||
                string.IsNullOrEmpty(CDTelephoneBox.Text))
            {
                //MessageBox.Show("All fields must be filled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowToast("ERROR", "All fields must be filled.");
            }
            else
            {
                if (CarPlateValidity(PlateNumberBox.Text).Equals(true) && IntegerValidity(CapacityBox.Text).Equals(true) && IntegerValidity(ChassisNumberBox.Text).Equals(true) && IntegerValidity(DTelephoneBox.Text).Equals(true) && IntegerValidity(CDTelephoneBox.Text).Equals(true))
                {
                    dataGridView1.Rows.Add(PlateNumberBox.Text, MakeBoxString, TypeBoxString, FUDateBoxString + Slash + FUMonthBoxString + Slash + FUYearBoxString, CapacityBox.Text + CapacityUnityBoxString, ChassisNumberBox.Text, DriverNameBox.Text, DTelephoneBox.Text, CDBox.Text, CDTelephoneBox.Text);

                }

            }

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
            CapacityUnityBox.Text = null;
            MakeBox.Text = null;
            TypeBox.Text = null;
            FUYearBox.Text = null;
            FUMonthBox.Text = null;
            FUDateBox.Text = null;
            CapacityBox.Text = null;
            ChassisNumberBox.Text = null;
            PlateNumberBox.Text = null;
            CDBox.Text = null;
            CDTelephoneBox.Text = null;
            DriverNameBox.Text = null;
            DTelephoneBox.Text = null;


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

        private void OpenXLFile()
        {
            dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
            dynamic xlworksheet = xlworkbook.Worksheets["NewVehicle"];
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

        private void SaveButton_Click(object sender, EventArgs e)
        {
            if (File.Exists(newFilePath))
            {
                // Open Excel and the workbook
                dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
                dynamic xlworksheet = xlworkbook.Worksheets["NewVehicle"];

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
