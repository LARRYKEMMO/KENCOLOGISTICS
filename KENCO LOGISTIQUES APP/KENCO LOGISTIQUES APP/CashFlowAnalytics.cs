﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ScottPlot;
using ScottPlot.Drawing.Colormaps;

namespace KENCO_LOGISTIQUES_APP
{
    public partial class CashFlowAnalytics : Form
    {
        public CashFlowAnalytics()
        {
            InitializeComponent();
            OpenXLFile();
            GetVehicles();
            SearchMechanics("");
        }

        private void OpenXLFile()
        {
            dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
            dynamic xlworksheet = xlworkbook.Worksheets["Accounting"];
            dynamic xlrange = xlworksheet.UsedRange;

            //dataGridView3.ColumnCount = xlrange.Columns.Count;
            dataGridView1.Rows.Clear();

            for (int xlrow = 2; xlrow <= xlrange.Rows.Count; xlrow++)
            {

                dataGridView1.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text,
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text,
                xlrange.Cells[xlrow, 6].Text, xlrange.Cells[xlrow, 7].Text, xlrange.Cells[xlrow, 8].Text);

            }

            DeleteEmptyRows(dataGridView1);

            xlworkbook.Close();
            xlapp.Quit();
        }


        private void Search_Click_Month()
        {
            income = 0;
            GetCashMonth();
            GetVehicleMonth();
            SearchMechanics3();
        }

        private void GetCashMonth()
        {
            string? digit;
            DateTime Date;
            string? dateConvert;


            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow && row.Visible)
                {
                    // Assuming the column contains numeric values as strings
                    string cellValue = Convert.ToString(row.Cells["Column8"].Value);

                    // Parse the string value to double
                    double value;
                    if (double.TryParse(cellValue, out value))
                    {
                        // Add the parsed value to the sum
                        //MessageBox.Show("Income: " + income, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        digit = dataGridView1.Rows[row.Index].Cells[7]?.Value.ToString();

                        if (digit != null)
                        {
                            Date = DateTime.Parse(digit);
                            dateConvert = Date.ToString("yyyy-MM");
                            //MessageBox.Show(digit, "Message");
                            if (dateConvert.Equals(MonthSearch))
                            {
                                income += value;
                            }

                        }
                    }
                    else
                    {
                        // Handle cases where parsing fails (invalid format)
                        // For example, you may log an error or skip the value
                    }
                }
            }
        }

        private void DetermineDate(string Month)
        {
            string Year;
            DateTime currentDate;

            currentDate = DateTime.Today;
            Year = currentDate.ToString("yyyy");
            
            if (Month.Equals("January"))
            {
                MonthSearch = Year + "-01";
            }
            else if (Month.Equals("February"))
            {
                MonthSearch = Year + "-02";
            }
            else if (Month.Equals("March"))
            {
                MonthSearch = Year + "-03";
            }
            else if (Month.Equals("April"))
            {
                MonthSearch = Year + "-04";
            }
            else if (Month.Equals("May"))
            {
                MonthSearch = Year + "-05";
            }
            else if (Month.Equals("June"))
            {
                MonthSearch = Year + "-06";
            }
            else if (Month.Equals("July"))
            {
                MonthSearch = Year + "-07";
            }
            else if (Month.Equals("August"))
            {
                MonthSearch = Year + "-08";
            }
            else if (Month.Equals("September"))
            {
                MonthSearch = Year + "-09";
            }
            else if (Month.Equals("October"))
            {
                MonthSearch = Year + "-10";
            }
            else if (Month.Equals("November"))
            {
                MonthSearch = Year + "-11";
            }
            else if (Month.Equals("December"))
            {
                MonthSearch = Year + "-12";
            }
        }

        private void GetVehicleMonth()
        {
            string? digit;
            string? NDigit;
            Color[] colors = { Color.Orange, Color.DarkGreen, Color.Indigo, Color.Violet, Color.Pink, Color.Blue, Color.Red, Color.Yellow, Color.DeepSkyBlue, Color.Cyan, Color.Magenta, Color.Gray, Color.Magenta, Color.DeepPink, Color.Purple, Color.Brown };
            var plt = new ScottPlot.Plot(620, 368);
            var plt2 = new ScottPlot.Plot(678, 367);
            double[] dataX = new double[0];
            double[] dataY = new double[0];
            string[] vehicles = new string[0];
            List<string> labels = new List<string>();

            plt.Palette = Palette.Amber;
            plt.Title("Cash Flow Chart");
            plt.XLabel("Vehicles");
            plt.YLabel("Profit / Loss");

            plt2.Palette = Palette.Amber;
            plt2.Title("Cash Flow Pie Chart");

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                // Skip the header row
                if (!dataGridView1.Rows[i].IsNewRow)
                {
                    // Access the current row using dataGridView1.Rows[i]
                    // Your code to process the row goes here
                    if (VisibleVehicles.Contains(dataGridView1.Rows[i].Cells[0].Value.ToString()))
                    {

                    }
                    else
                    {
                        VisibleVehicles.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                        VisibleVehicle.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                        VList.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    }
                }
            }

            VisibleVehicle.Sort();

            for (int i = 0; i < VisibleVehicle.Count; i++)
            {
                income = 0;
                digit = VisibleVehicle[i]?.ToString();
                NDigit = VList[i]?.ToString();
                if (digit != null)
                {
                    vehicles = vehicles.Append(NDigit).ToArray();
                    SearchMechanics(digit);
                    GetCashMonth();
                    GetDigit(digit);
                    labels.Add(digit);
                    dataX = dataX.Append(doubleX).ToArray();
                    dataY = dataY.Append(income).ToArray();
                    plt.PlotBar(dataX, dataY, fillColor: Color.Purple, barWidth: 20.0);
                    plt.XTicks(dataX, vehicles);
                    plt2.PlotPie(dataY, labels.ToArray(), colors, showPercentages: true);
                }


            }

            plt.SaveFig("CashFlowChart.png");
            pictureBox1.ImageLocation = "CashFlowChart.png";

            plt2.SaveFig("CashFlowParetto.png");
            pictureBox2.ImageLocation = "CashFlowParetto.png";
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

        private void Search_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Check if the row is not a new row
                if (!row.IsNewRow)
                {
                    // Access the cell value for the 'Amount (FCFA)' column
                    string amountCellValue = Convert.ToString(row.Cells["Column1"].Value);

                    // Check if the cell value contains the search text
                    if (amountCellValue.Contains(SearchBox.Text))
                    {
                        // Show the row if the cell value contains the search text
                        row.Visible = true;
                    }
                    else
                    {
                        // Hide the row if the cell value does not contain the search text
                        row.Visible = false;
                    }
                }
            }
        }

        private void SearchMechanics(string Text)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Check if the row is not a new row
                if (!row.IsNewRow)
                {
                    // Access the cell value for the 'Amount (FCFA)' column
                    string amountCellValue = Convert.ToString(row.Cells["Column4"].Value);

                    // Check if the cell value contains the search text
                    if (amountCellValue.Contains(Text))
                    {
                        // Show the row if the cell value contains the search text
                        row.Visible = true;
                    }
                    else
                    {
                        // Hide the row if the cell value does not contain the search text
                        row.Visible = false;
                    }
                }
            }
        }

        private void SearchMechanics3()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Check if the row is not a new row
                if (!row.IsNewRow)
                {
                    // Access the cell value for the 'Amount (FCFA)' column
                    string amountCellValue = Convert.ToString(row.Cells["Column9"].Value);

                    // Check if the cell value contains the search text
                    if (amountCellValue.Contains("2024-03"))
                    {
                        // Show the row if the cell value contains the search text
                        row.Visible = true;
                    }
                    else
                    {
                        // Hide the row if the cell value does not contain the search text
                        row.Visible = false;
                    }
                }
            }
        }

        private void GetVehicles()
        {
            string? digit;
            string? NDigit;
            Color[] colors = { Color.Orange, Color.DarkGreen, Color.Indigo, Color.Violet, Color.Pink, Color.Blue, Color.Red, Color.Yellow, Color.DeepSkyBlue, Color.Cyan, Color.Magenta, Color.Gray, Color.Magenta, Color.DeepPink, Color.Purple, Color.Brown};
            var plt = new ScottPlot.Plot(620, 368);
            var plt2 = new ScottPlot.Plot(678, 367);
            double[] dataX = new double[0];
            double[] dataY = new double[0];
            string[] vehicles = new string[0];
            List<string> labels = new List<string>();

            plt.Palette = Palette.Amber;
            plt.Title("Cash Flow Chart");
            plt.XLabel("Vehicles");
            plt.YLabel("Profit / Loss");

            plt2.Palette = Palette.Amber;
            plt2.Title("Cash Flow Pie Chart");

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                // Skip the header row
                if (!dataGridView1.Rows[i].IsNewRow)
                {
                    // Access the current row using dataGridView1.Rows[i]
                    // Your code to process the row goes here
                    if (Vehicles.Contains(dataGridView1.Rows[i].Cells[0].Value.ToString()))
                    {

                    }
                    else
                    {
                        Vehicles.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                        VehiclesList.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                        VList.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                        //MessageBox.Show("Income: " + dataGridView1.Rows[i].Cells[0].Value.ToString(), "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }

            VehiclesList.Sort();

            //MessageBox.Show("Income: " + VehiclesList.Count, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);

            for (int i = 0; i < VehiclesList.Count; i++)
            {
                income = 0;
                digit = VehiclesList[i]?.ToString();
                NDigit = VList[i]?.ToString();
                if (digit != null)
                {
                    vehicles = vehicles.Append(NDigit).ToArray();
                    SearchMechanics(digit);
                    GetCash(dataGridView1);
                    GetDigit(digit);
                    labels.Add(digit);
                    dataX = dataX.Append(doubleX).ToArray();
                    dataY = dataY.Append(income).ToArray();
                    plt.PlotBar(dataX, dataY, fillColor: Color.Blue, barWidth: 20.0);
                    plt.XTicks(dataX, vehicles);
                    plt2.PlotPie(dataY, labels.ToArray(), colors, showPercentages: true);
                }


            }

            plt.SaveFig("CashFlowChart.png");
            pictureBox1.ImageLocation = "CashFlowChart.png";

            plt2.SaveFig("CashFlowParetto.png");
            pictureBox2.ImageLocation = "CashFlowParetto.png";
        }

        private void GetDigit(string temp)
        {
            string Number;

            Number = temp.Substring(2, 3);
            //MessageBox.Show("Income: " + Number, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            doubleX = double.Parse(Number);
        }

        private void GetCash(DataGridView dataGridView)
        {


            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (!row.IsNewRow && row.Visible)
                {
                    // Assuming the column contains numeric values as strings
                    string cellValue = Convert.ToString(row.Cells["Column8"].Value);

                    // Parse the string value to double
                    double value;
                    if (double.TryParse(cellValue, out value))
                    {
                        // Add the parsed value to the sum
                        //MessageBox.Show("Income: " + income, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        income += value;
                    }
                    else
                    {
                        // Handle cases where parsing fails (invalid format)
                        // For example, you may log an error or skip the value
                    }
                }
            }
        }

        private void GetCash2(DataGridView dataGridView)
        {


            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (!row.IsNewRow && row.Visible)
                {
                    // Assuming the column contains numeric values as strings
                    string cellValue = Convert.ToString(row.Cells["Column10"].Value);

                    // Parse the string value to double
                    double value;
                    if (double.TryParse(cellValue, out value))
                    {
                        // Add the parsed value to the sum
                        //MessageBox.Show("Income: " + income, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        income2 += value;
                    }
                    else
                    {
                        // Handle cases where parsing fails (invalid format)
                        // For example, you may log an error or skip the value
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SearchMechanics("");
            income = 0;
            GetCash(dataGridView1);
            GetVehicles();
            SearchMechanics("");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Search_Click_Month();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.ShowDialog();
            this.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string month;

            month = comboBox1.Text;
            income = 0;
            DetermineDate(month);
            GetCashMonth();
            GetVehicleMonth();
            SearchMechanics3();

        }
    }
}
