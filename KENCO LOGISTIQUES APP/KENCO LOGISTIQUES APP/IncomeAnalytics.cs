using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ScottPlot;




namespace KENCO_LOGISTIQUES_APP
{
    public partial class IncomeAnalytics : Form
    {
        public IncomeAnalytics()
        {
            InitializeComponent();
            OpenXLFile();
        }

        private void OpenXLFile()
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
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text);

            }

            DeleteEmptyRows(dataGridView1);

            xlworkbook.Close();
            xlapp.Quit();
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

        private void DrawGraph()
        {
            var plt = new ScottPlot.Plot(596, 165);
            int pointcount = 60;
            double[] x = new double[] { 1, 2, 3, 4, 5 };
            double[] Sin = new double[] { 1, 4, 9, 16, 25 };
            //double[] Cos = DataGen.Cos(pointcount);

            plt.Palette = Palette.Amber;
            plt.Title("My chart");
            plt.XLabel("x-axis");
            plt.YLabel("y-axis");
            plt.AddScatter(x, Sin);
            //plt.AddScatter(x, Cos);
            plt.SaveFig("scater.png");
            pictureBox1.ImageLocation = "scater.png";
        }


        private void IncomeAnalytics_Load(object sender, EventArgs e)
        {
            //Search1 = "CE182EP";
            //Search2 = "transportation services";
            //CreatePieChart();
            GetVehicles();
            //SearchMechanics2(Search2);
            GetDescription();
        }

        private void SearchMechanics(string Text)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Check if the row is not a new row
                if (!row.IsNewRow)
                {
                    // Access the cell value for the 'Amount (FCFA)' column
                    string amountCellValue = Convert.ToString(row.Cells["Column1"].Value);

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


        private void SearchMechanics2(string Text2)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Check if the row is not a new row
                if (!row.IsNewRow)
                {
                    // Access the cell value for the 'Amount (FCFA)' column
                    //string amountCellValue = Convert.ToString(row.Cells["Column1"].Value);
                    string amountCellValue2 = Convert.ToString(row.Cells["Column2"].Value);

                    // Check if the cell value contains the search text
                    //if (amountCellValue.Contains(Text))
                    //{
                    //    // Show the row if the cell value contains the search text
                    //    row.Visible = true;
                    if(amountCellValue2.Contains(Text2))
                    {
                        row.Visible = true;
                    }
                    else
                    {
                        row.Visible = false;
                    }
                    //}
                    //else
                    //{
                    //    // Hide the row if the cell value does not contain the search text
                    //    row.Visible = false;
                    //}
                }
            }
        }

        private void GetVehicles()
        {
            string? digit;
            var plt = new ScottPlot.Plot(596, 368);
            double[] dataX = new double[0];
            double[] dataY = new double[0];

            plt.Palette = Palette.Amber;
            plt.Title("Income Chart");
            plt.XLabel("Vehicles");
            plt.YLabel("Income");

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
                        //MessageBox.Show("Income: " + dataGridView1.Rows[i].Cells[0].Value.ToString(), "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }

            VehiclesList.Sort();

            //MessageBox.Show("Income: " + VehiclesList.Count, "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);

            for (int i = 0; i < VehiclesList.Count;i++)
            {
                income = 0;
                digit = VehiclesList[i]?.ToString();
                if (digit != null)
                {
                    SearchMechanics(digit);
                    GetCash(dataGridView1);
                    GetDigit(digit);
                    dataX = dataX.Append(doubleX).ToArray();
                    dataY = dataY.Append(income).ToArray();
                    plt.AddScatter(dataX, dataY);
                }
                
            }

            plt.SaveFig("IncomeChart.png");
            pictureBox1.ImageLocation = "IncomeChart.png";
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
                    string cellValue = Convert.ToString(row.Cells["Column10"].Value);

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

        private void CreatePieChart()
        {
            var plt = new ScottPlot.Plot(400, 357);
            // Define data
            double[] values = { 35, 25, 20, 10, 5 };
            string[] labels = { "Apple", "Banana", "Orange", "Grapes", "Watermelon" };

            // Create ScottPlot pie chart
            plt.PlotPie(values, labels);
            plt.Title("Fruit Distribution");

            // Render the plot
            plt.SaveFig("Fruit.png");
            pictureBox2.ImageLocation = "Fruit.png";
        }

        private void GetDescription()
        {
            string? digit;
            var plt = new ScottPlot.Plot(391, 367);
            double[] values = new double[0];
            List<string> labels = new List<string>();


            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                // Skip the header row
                if (!dataGridView1.Rows[i].IsNewRow)
                {
                    // Access the current row using dataGridView1.Rows[i]
                    // Your code to process the row goes here
                    if (Description.Contains(dataGridView1.Rows[i].Cells[2].Value.ToString()))
                    {

                    }
                    else
                    {
                        Description.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());
                        DescriptionList.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());
                        //MessageBox.Show("Income: " + dataGridView1.Rows[i].Cells[0].Value.ToString(), "DataGridView Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }

            for (int i = 0; i < DescriptionList.Count; i++)
            {
                income2 = 0;
                digit = DescriptionList[i]?.ToString();
                if (digit != null)
                {
                    SearchMechanics2(digit);
                    GetCash2(dataGridView1);
                    values = values.Append(income2).ToArray();
                    labels.Add(digit);
                    plt.PlotPie(values, labels.ToArray(), showPercentages: true, showLabels: false);

                }

            }


            plt.Title("Paretto Chart");
            plt.SaveFig("Paretto.png");
            pictureBox2.ImageLocation = "Paretto.png";

        }
    }
}
