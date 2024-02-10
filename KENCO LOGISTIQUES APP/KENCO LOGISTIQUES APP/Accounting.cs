using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.Linq;
using System.Text;
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
        }

        private void AddIncome_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(IncomeAmountBox.Text, IncomeDescription.Text, IncomeDateBox.Text + slash + IncomeMonthBox.Text + slash + IncomeYearBox.Text);
            //DGVSize++;
        }

        private void DeleteIncome_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
            }
        }

        private void ResetIncome_Click(object sender, EventArgs e)
        {
            IncomeAmountBox.Text = null;
            IncomeDescription.Text = null;
            IncomeDateBox.Text = null;
            IncomeMonthBox.Text = null;
            IncomeYearBox.Text = null;

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
                    MessageBox.Show("All rows are to be deleted " + exe, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void AddExpense_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Add(ExpenseAmountBox.Text, ExpenseDecriptionBox.Text, ExpenseDayBox.Text + slash + ExpenseMonthBox.Text + slash + ExpenseYearBox.Text);
        }

        private void DeleteExpense_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView2.SelectedRows)
            {
                dataGridView2.Rows.RemoveAt(item.Index);
            }
        }

        private void ResetExpense_Click(object sender, EventArgs e)
        {
            ExpenseAmountBox.Text = null;
            ExpenseDecriptionBox.Text = null;
            ExpenseDayBox.Text = null;
            ExpenseMonthBox.Text = null;
            ExpenseYearBox.Text = null;

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
                    MessageBox.Show("All rows are to be deleted " + exe, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }


        private void GetCash(DataGridView dataGridView)
        {
            income = 0;
            float cellValue = 0;
            string? digit;

            for (int counter = 0; counter < dataGridView.Rows.Count; counter++)
            {
                digit = dataGridView.Rows[counter].Cells[0].Value?.ToString();

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

            MessageBox.Show("Income: " + income, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetDescription(DataGridView dataGridView)
        {
            Description = "";
           string? digit;
           

            for (int counter = 0; counter < dataGridView.Rows.Count; counter++)
            {
                digit = dataGridView.Rows[counter].Cells[1].Value?.ToString();

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

            MessageBox.Show("Description: " + Description, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetDate(DataGridView dataGridView)
        {
            date = "Null";
            string? digit;
            

            digit = dataGridView.Rows[0].Cells[2].Value?.ToString();
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
                digit = dataGridView.Rows[counter].Cells[2].Value?.ToString();

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
                MessageBox.Show("Date: " + date, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }
        private void AddNew_Click(object sender, EventArgs e)
        {
            float NetCash;


            GetCash(dataGridView1);
            Income = income;
            GetDescription(dataGridView1);
            DescriptionIncome = Description;
            GetDate(dataGridView1);
            DateIncome = date;
            GetCash(dataGridView2);
            Expenses = income;
            GetDescription(dataGridView2);
            DescriptionExpenses = Description;
            GetDate(dataGridView2);
            DateExpenses = date;

            NetCash = Income - Expenses;
            Profit_Loss = NetCash.ToString();

            if(DateIncome.Equals(DateExpenses))
            {
                Date = DateIncome;
            }

            dataGridView3.Rows.Add(Income.ToString(), DescriptionIncome, Expenses.ToString(), DescriptionExpenses, Profit_Loss, Date);
           
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
                    MessageBox.Show("All rows are to be deleted " + exe, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void OpenXLFile()
        {
            dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
            dynamic xlworksheet = xlworkbook.Worksheets["Sheet4"];
            dynamic xlrange = xlworksheet.UsedRange;

            dataGridView3.ColumnCount = xlrange.Columns.Count;

            for (int xlrow = 2; xlrow <= xlrange.Rows.Count; xlrow++)
            {

                dataGridView3.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text,
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text,
                xlrange.Cells[xlrow, 6].Text);

            }

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
                dynamic xlworksheet = xlworkbook.Worksheets["Sheet4"];

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

                MessageBox.Show("Data saved successfully to Excel file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("The specified Excel file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
