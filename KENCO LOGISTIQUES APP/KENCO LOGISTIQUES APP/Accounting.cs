using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KENCO_LOGISTIQUES_APP
{
    public partial class Accounting : Form
    {
        public Accounting()
        {
            InitializeComponent();
            OpenXLFile();
            //OpenIncomeXLFile();
            //OpenExpensesXLFile();
        }

        private void ResetIncomeClick()
        {
            IncomeAmountBox.Text = null;
            IncomeDescription.Text = null;
            IncomeDateBox.Text = null;
            IncomeMonthBox.Text = null;
            IncomeYearBox.Text = null;
            IVBox.Text = null;
            NicknameBox.Text = null;

            VehicleIncomeList.Clear();
            NicknameIncomeList.Clear();
            IncomeList.Clear();
            DescriptionIncomeList.Clear();
            DateIncomeList.Clear();

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
                    ShowToast("ERROR", "Rows are not to be deleted");
                }
            }
        }

        private void ResetExpenseClick()
        {
            ExpenseAmountBox.Text = null;
            ExpenseDecriptionBox.Text = null;
            ExpenseDayBox.Text = null;
            ExpenseMonthBox.Text = null;
            ExpenseYearBox.Text = null;
            EVBox.Text = null;
            NicknameBox2.Text = null;

            VehicleExpenseList.Clear();
            NicknameExpenseList.Clear();
            ExpenseList.Clear();
            DescriptionExpenseList.Clear();
            DateExpenseList.Clear();

            // Clears Data Grid View
            int numRows = dataGridView2.Rows.Count;
            for (int i = 0; i < numRows; i++)
            {
                try
                {
                    int max = dataGridView2.Rows.Count - 1;
                    dataGridView2.Rows.Remove(dataGridView2.Rows[max]);
                }
                catch (Exception exe)
                {
                    //MessageBox.Show("All rows are to be deleted " + exe, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ShowToast("ERROR", "Rows are not to be deleted");
                }
            }
        }


        private void GetVehicle(DataGridView dataGridView)
        {
            VehicleNumber = "Null";
            string? digit;


            digit = dataGridView.Rows[0].Cells[0].Value?.ToString();
            int CorrectCounter = 0;

            if (digit != null)
            {
                VehicleNumber = digit;
            }
            else
            {
                VehicleNumber = "None";
            }

            for (int counter = 0; counter < dataGridView.Rows.Count; counter++)
            {
                digit = dataGridView.Rows[counter].Cells[0].Value?.ToString();

                if (digit != null)
                {
                    if (VehicleNumber.Equals(digit))
                    {
                        CorrectCounter++;
                    }
                    else
                    {
                        VehicleNumber = "None";
                    }

                }

            }

            if (CorrectCounter == 3)
            {
                //MessageBox.Show("Date: " + VehicleNumber, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void GetNickName(DataGridView dataGridView)
        {
            NickName = "Null";
            string? digit;


            digit = dataGridView.Rows[0].Cells[1].Value?.ToString();
            int CorrectCounter = 0;

            if (digit != null)
            {
                NickName = digit;
            }
            else
            {
                NickName = "None";
            }

            for (int counter = 0; counter < dataGridView.Rows.Count; counter++)
            {
                digit = dataGridView.Rows[counter].Cells[1].Value?.ToString();

                if (digit != null)
                {
                    if (VehicleNumber.Equals(digit))
                    {
                        CorrectCounter++;
                    }
                    else
                    {
                        NickName = "None";
                    }

                }

            }

            if (CorrectCounter == 3)
            {
                //MessageBox.Show("Date: " + NickName, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void GetCash(DataGridView dataGridView)
        {
            income = 0;
            float cellValue = 0;
            string? digit;

            for (int counter = 0; counter < dataGridView.Rows.Count; counter++)
            {
                digit = dataGridView.Rows[counter].Cells[2].Value?.ToString();

                if (digit != null)
                {
                    cellValue = float.Parse(digit);
                }
                else
                {
                    cellValue = 1;
                }
                income += cellValue;
            }

            //MessageBox.Show("Income: " + income, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetDescription(DataGridView dataGridView)
        {
            Description = "";
           string? digit;
           

            for (int counter = 0; counter < dataGridView.Rows.Count; counter++)
            {
                digit = dataGridView.Rows[counter].Cells[3].Value?.ToString();

                if (digit != null)
                {
                    if(counter >= dataGridView.Rows.Count - 1)
                    {
                        Description += digit;
                    }
                    else
                    {
                        Description += digit + ", ";

                    }
                }
                else
                {
                    Description = "None";
                }
            }

            //MessageBox.Show("Description: " + Description, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetDate(DataGridView dataGridView)
        {
            date = "Null";
            string? digit;
            

            digit = dataGridView.Rows[0].Cells[4].Value?.ToString();
            int CorrectCounter = 0;

            if (digit != null)
            {
                date = digit;
            }
            else
            {
                date = "None";
            }

            for (int counter = 0; counter < dataGridView.Rows.Count; counter++)
            {
                digit = dataGridView.Rows[counter].Cells[4].Value?.ToString();

                if (digit != null)
                {
                    if (date.Equals(digit))
                    {
                        CorrectCounter ++;
                    }
                    else
                    {
                        date = "None";
                    }
                    
                }
                
            }

            if (CorrectCounter == 3)
            {
                //MessageBox.Show("Date: " + date, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }
        private void AddNew_Click(object sender, EventArgs e)
        {
            float NetCash;
            int counter = 0;

            if(dataGridView1.Rows.Count > 0 && dataGridView2.Rows.Count > 0)
            {
                GetCash(dataGridView1);
                Income = income;
                GetDescription(dataGridView1);
                DescriptionIncome = Description;
                GetDate(dataGridView1);
                DateIncome = date;
                GetVehicle(dataGridView1);
                GetNickName(dataGridView1);
                VIString = VehicleNumber;
                GetCash(dataGridView2);
                Expenses = income;
                GetDescription(dataGridView2);
                DescriptionExpenses = Description;
                GetDate(dataGridView2);
                DateExpenses = date;
                GetVehicle(dataGridView2);
                GetNickName(dataGridView2);
                VEString = VehicleNumber;

                NetCash = Income - Expenses;
                Profit_Loss = NetCash.ToString();

                if(DateIncome.Equals(DateExpenses))
                {
                    Date = DateIncome;
                    counter++;
                }

                if (VIString.Equals(VEString))
                {
                    VehicleNumber = VIString;
                    counter++;
                }

                if (NicknameBox.Text.Equals(NicknameBox2.Text))
                {
                    NickName = NicknameBox.Text;
                    counter++;
                }

                if(counter.Equals(3))
                {
                    if(CarPlateValidity(VehicleNumber).Equals(true) && IntegerValidity(Income.ToString()).Equals(true) && IntegerValidity(Expenses.ToString()).Equals(true) && IntegerValidity(Profit_Loss).Equals(true))
                    {
                        dataGridView3.Rows.Add(VehicleNumber, NickName, Income.ToString(), DescriptionIncome, Expenses.ToString(), DescriptionExpenses, Profit_Loss, Date);

                    }

                }
                else
                {
                    //MessageBox.Show("Error - Check your Income and Expense entries to find the problem");
                    ShowToast("ERROR", "Check your Income and Expense entries to find the problem");
                }

            }
            else
            {
                ShowToast("ERROR", "Cannot Calculate with Empty Income and Expenses Fields");
            }
           
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Delete_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView3.SelectedRows)
            {
                dataGridView3.Rows.RemoveAt(item.Index);
            }
        }

        private void Reset_Click(object sender, EventArgs e)
        {
            // Clears Data Grid View
            int numRows = dataGridView3.Rows.Count;
            for (int i = 0; i < numRows; i++)
            {
                try
                {
                    int max = dataGridView3.Rows.Count - 1;
                    dataGridView3.Rows.Remove(dataGridView3.Rows[max]);
                }
                catch (Exception exe)
                {
                    //MessageBox.Show("All rows are to be deleted " + exe, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ShowToast("ERROR", "Rows are not to be deleted");
                }
            }

            ResetExpenseClick();
            ResetIncomeClick();

        }

        private void OpenXLFile()
        {
            dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
            dynamic xlworksheet = xlworkbook.Worksheets["Accounting"];
            dynamic xlrange = xlworksheet.UsedRange;

            //dataGridView3.ColumnCount = xlrange.Columns.Count;
            dataGridView3.Rows.Clear();

            for (int xlrow = 2; xlrow <= xlrange.Rows.Count; xlrow++)
            {

                dataGridView3.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text,
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text,
                xlrange.Cells[xlrow, 6].Text, xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text);

            }

            DeleteEmptyRows(dataGridView3);

            xlworkbook.Close();
            xlapp.Quit();
        }

        private void OpenIncomeXLFile()
        {
            dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
            dynamic xlworksheet = xlworkbook.Worksheets["Income"];
            dynamic xlrange = xlworksheet.UsedRange;

            //dataGridView3.ColumnCount = xlrange.Columns.Count;
            dataGridView1.Rows.Clear();

            for (int xlrow = 2; xlrow <= xlrange.Rows.Count; xlrow++)
            {

                dataGridView1.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text,
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text);

            }

            DeleteEmptyRows(dataGridView1);

            xlworkbook.Close();
            xlapp.Quit();
        }

        private void OpenExpensesXLFile()
        {
            dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
            dynamic xlworksheet = xlworkbook.Worksheets["Expenses"];
            dynamic xlrange = xlworksheet.UsedRange;

            //dataGridView3.ColumnCount = xlrange.Columns.Count;
            dataGridView2.Rows.Clear();

            for (int xlrow = 2; xlrow <= xlrange.Rows.Count; xlrow++)
            {

                dataGridView2.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text,
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text);

            }

            DeleteEmptyRows(dataGridView2);

            xlworkbook.Close();
            xlapp.Quit();
        }

        private void SaveIncome()
        {
            if (File.Exists(newFilePath))
            {
                // Open Excel and the workbook
                dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
                dynamic xlworksheet = xlworkbook.Worksheets["Income"];

                // Clear existing data in the Excel worksheet
                xlworksheet.Cells.ClearContents();

                for (int col = 1; col <= dataGridView1.ColumnCount; col++)
                {
                    xlworksheet.Cells[1, col].Value = dataGridView1.Columns[col - 1].HeaderText;
                }

                if (dataGridView3.RowCount > 0)
                {
                    OpenIncomeXLFile();
                }


                for(int i = 0; i < IncomeList.Count; i++)
                {
                    dataGridView1.Rows.Add(VehicleIncomeList[i], NicknameIncomeList[i], IncomeList[i], DescriptionIncomeList[i], DateIncomeList[i]);
                }

                DeleteEmptyRows(dataGridView1);

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


        private void SaveExpenses()
        {
            if (File.Exists(newFilePath))
            {
                // Open Excel and the workbook
                dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
                dynamic xlworksheet = xlworkbook.Worksheets["Expenses"];

                // Clear existing data in the Excel worksheet
                xlworksheet.Cells.ClearContents();

                for (int col = 1; col <= dataGridView2.ColumnCount; col++)
                {
                    xlworksheet.Cells[1, col].Value = dataGridView2.Columns[col - 1].HeaderText;
                }

                if(dataGridView3.RowCount > 0)
                {
                    OpenExpensesXLFile();

                }

                for (int i = 0; i < ExpenseList.Count; i++)
                {
                    dataGridView2.Rows.Add(VehicleExpenseList[i], NicknameExpenseList[i], ExpenseList[i], DescriptionExpenseList[i], DateExpenseList[i]);
                }

                DeleteEmptyRows(dataGridView2);

                // Write data from the DataGridView to the Excel worksheet
                for (int row = 0; row < dataGridView2.Rows.Count; row++)
                {
                    for (int col = 0; col < dataGridView2.Columns.Count; col++)
                    {
                        xlworksheet.Cells[row + 2, col + 1].Value = dataGridView2.Rows[row].Cells[col].Value;
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

        private void SaveButton_Click(object sender, EventArgs e)
        {
            if (File.Exists(newFilePath))
            {
                // Open Excel and the workbook
                dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
                dynamic xlworksheet = xlworkbook.Worksheets["Accounting"];

                // Clear existing data in the Excel worksheet
                xlworksheet.Cells.ClearContents();

                for (int col = 1; col <= dataGridView3.ColumnCount; col++)
                {
                    xlworksheet.Cells[1, col].Value = dataGridView3.Columns[col - 1].HeaderText;
                }

                // Write data from the DataGridView to the Excel worksheet
                for (int row = 0; row < dataGridView3.Rows.Count; row++)
                {
                    for (int col = 0; col < dataGridView3.Columns.Count; col++)
                    {
                        xlworksheet.Cells[row + 2, col + 1].Value = dataGridView3.Rows[row].Cells[col].Value;
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

            SaveIncome();
            SaveExpenses();
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

        private void MainMenu_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.ShowDialog();
            this.Close();
        }

        public void ShowToast(string Type, string Message)
        {
            ToastForm toastForm = new ToastForm(Type, Message);
            toastForm.ShowDialog();
        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void EVBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void ExpenseAmountBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void AddIncome_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(IVBox.Text) ||
                string.IsNullOrEmpty(NicknameBox.Text) ||
                string.IsNullOrEmpty(IncomeAmountBox.Text) ||
                string.IsNullOrEmpty(IncomeDescription.Text) ||
                string.IsNullOrEmpty(IncomeDateBox.Text) ||
                string.IsNullOrEmpty(IncomeMonthBox.Text) ||
                string.IsNullOrEmpty(IncomeYearBox.Text))
            {
                //MessageBox.Show("All fields must be filled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowToast("ERROR", "All Income fields must be filled.");
            }
            else
            {
                if (CarPlateValidity(IVBox.Text).Equals(true) && IntegerValidity(IncomeAmountBox.Text.ToString()).Equals(true))
                {
                    dataGridView1.Rows.Add(IVBox.Text, NicknameBox.Text, IncomeAmountBox.Text, IncomeDescription.Text, IncomeDateBox.Text + slash + IncomeMonthBox.Text + slash + IncomeYearBox.Text);
                    VehicleIncomeList.Add(IVBox.Text);
                    NicknameIncomeList.Add(NicknameBox.Text);
                    IncomeList.Add(IncomeAmountBox.Text);
                    DescriptionIncomeList.Add(IncomeDescription.Text);
                    DateIncomeList.Add(IncomeDateBox.Text + slash + IncomeMonthBox.Text + slash + IncomeYearBox.Text);

                }

            }
        }

        private void DeleteIncome_Click_1(object sender, EventArgs e)
        {
            List<int> indicesToRemove = new List<int>();

            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                indicesToRemove.Add(item.Index);

            }

            foreach (int index in indicesToRemove.OrderByDescending(i => i))
            {
                dataGridView1.Rows.RemoveAt(index);

                if (index < VehicleIncomeList.Count && index < IncomeList.Count &&
                    index < DescriptionIncomeList.Count && index < DateIncomeList.Count)
                {
                    VehicleIncomeList.RemoveAt(index);
                    NicknameIncomeList.RemoveAt(index);
                    IncomeList.RemoveAt(index);
                    DescriptionIncomeList.RemoveAt(index);
                    DateIncomeList.RemoveAt(index);
                }
            }
        }

        private void ResetIncome_Click_1(object sender, EventArgs e)
        {
            ResetIncomeClick();
        }

        private void AddExpense_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(EVBox.Text) ||
                string.IsNullOrEmpty(NicknameBox2.Text) ||
                string.IsNullOrEmpty(ExpenseAmountBox.Text) ||
                string.IsNullOrEmpty(ExpenseDecriptionBox.Text) ||
                string.IsNullOrEmpty(ExpenseDayBox.Text) ||
                string.IsNullOrEmpty(ExpenseMonthBox.Text) ||
                string.IsNullOrEmpty(ExpenseYearBox.Text))
            {
                //MessageBox.Show("All fields must be filled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowToast("ERROR", "All Expense fields must be filled.");
            }
            else
            {
                if (CarPlateValidity(EVBox.Text).Equals(true) && IntegerValidity(ExpenseAmountBox.Text.ToString()).Equals(true))
                dataGridView2.Rows.Add(EVBox.Text, NicknameBox2.Text, ExpenseAmountBox.Text, ExpenseDecriptionBox.Text, ExpenseDayBox.Text + slash + ExpenseMonthBox.Text + slash + ExpenseYearBox.Text);
                VehicleExpenseList.Add(EVBox.Text);
                NicknameExpenseList.Add(NicknameBox2.Text);
                ExpenseList.Add(ExpenseAmountBox.Text);
                DescriptionExpenseList.Add(ExpenseDecriptionBox.Text);
                DateExpenseList.Add(ExpenseDayBox.Text + slash + ExpenseMonthBox.Text + slash + ExpenseYearBox.Text);

            }
        }

        private void DeleteExpense_Click_1(object sender, EventArgs e)
        {
            List<int> indicesToRemove = new List<int>();

            foreach (DataGridViewRow item in this.dataGridView2.SelectedRows)
            {
                indicesToRemove.Add(item.Index);

            }

            foreach (int index in indicesToRemove.OrderByDescending(i => i))
            {
                dataGridView2.Rows.RemoveAt(index);

                if (index < VehicleExpenseList.Count && index < ExpenseList.Count &&
                    index < DescriptionExpenseList.Count && index < DateExpenseList.Count)
                {
                    VehicleExpenseList.RemoveAt(index);
                    NicknameExpenseList.RemoveAt(index);
                    ExpenseList.RemoveAt(index);
                    DescriptionExpenseList.RemoveAt(index);
                    DateExpenseList.RemoveAt(index);
                }
            }
        }

        private void ResetExpense_Click_1(object sender, EventArgs e)
        {
            ResetExpenseClick();
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
