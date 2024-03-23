using ScottPlot;
using ScottPlot.Drawing.Colormaps;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace KENCO_LOGISTIQUES_APP
{
    public partial class Reminder : Form
    {
        public Reminder()
        {
            InitializeComponent();
            OpenXLFile();
        }

        private void Reminder_Load(object sender, EventArgs e)
        {

        }

        private void DateGenerator(DateTime Itommorrow, DateTime TVtommorrow, DateTime RCtommorrow)
        {
            DateTime tomorrowI;
            DateTime tomorrowTV;
            DateTime tomorrowRC;


            for (int i = 0; i < 31; i++)
            {
                tomorrowI = Itommorrow.AddDays(i);
                IDateList.Add(tomorrowI);
            }

            for (int i = 0; i < 31; i++)
            {
                tomorrowTV = TVtommorrow.AddDays(i);
                TVDateList.Add(tomorrowTV);
            }

            for (int i = 0; i < 31; i++)
            {
                tomorrowRC = RCtommorrow.AddDays(i);
                RCDateList.Add(tomorrowRC);
            }
        }

        public async void reminderCheck()
        {
            await Task.Delay(1000);
            

            bool Expired1 = false;
            bool Expired2 = false;
            bool Expired3 = false;

            string? digit1;
            string? digit2;
            string? digit3;
            string? digit4;
            string? digit5;
            string? digit6;

            DateTime dateTime1;
            DateTime dateTime2;
            DateTime dateTime3;
            DateTime dateTime4;
            DateTime dateTime5;
            DateTime dateTime6;

            // Get today's date
            DateTime currentDate = DateTime.Now.Date;


            for(int j = 0; j < 31; j++)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if(IDateList != null && TVDateList != null && RCDateList != null)
                    {
                        IDateList.Clear();
                        TVDateList.Clear();
                        RCDateList.Clear();
                    }

                    digit1 = dataGridView1.Rows[i].Cells[2].Value?.ToString();
                    digit2 = dataGridView1.Rows[i].Cells[4].Value?.ToString();
                    digit3 = dataGridView1.Rows[i].Cells[6].Value?.ToString();
                    digit4 = dataGridView1.Rows[i].Cells[1].Value?.ToString();
                    digit5 = dataGridView1.Rows[i].Cells[3].Value?.ToString();
                    digit6 = dataGridView1.Rows[i].Cells[5].Value?.ToString();

                    if (digit1 != null && digit2 != null && digit3 != null && digit4 != null && digit5 != null && digit6 != null)
                    {

                        DateTime.TryParseExact(digit1, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTime1);
                        DateTime.TryParseExact(digit2, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTime2);
                        DateTime.TryParseExact(digit3, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTime3);
                        DateTime.TryParseExact(digit4, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTime4);
                        DateTime.TryParseExact(digit5, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTime5);
                        DateTime.TryParseExact(digit6, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTime6);

                        DateGenerator(dateTime1, dateTime2, dateTime3);


                        if (dateTime1.Equals(IDateList[j]))
                        {
                               
                            if (dateTime4.Equals(currentDate))
                            {
                                //MessageBox.Show("Reminder: Insurance Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Today", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                ShowToast("REMINDER", "Reminder: Insurance Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Today");
                                Expired1 = true;
                                RowList.Add(i);
                                ColList.Add(1);
                            }

                            if (Expired1 == false)
                            {
                                //MessageBox.Show("Reminder: Insurance Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Soon", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                ShowToast("WARNING", "Reminder: Insurance Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Soon");
                                RowList2.Add(i);
                                ColList2.Add(2);
                            }
                        }


                        
                        if (dateTime2.Equals(TVDateList[j]))
                        {

                            if (dateTime5.Equals(currentDate))
                            {
                                //MessageBox.Show("Reminder: Technical Visit Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Today", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                ShowToast("REMINDER", "Reminder: Technical Visit Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Today");
                                Expired2 = true;
                                RowList.Add(i);
                                ColList.Add(3);
                            }

                            if (Expired2 == false)
                            {
                                //MessageBox.Show("Reminder: Technical Visit Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Soon", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                ShowToast("WARNING", "Reminder: Technical Visit Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Soon");
                                RowList2.Add(i);
                                ColList2.Add(4);
                            }
                        
                        }


                        
                        if (dateTime3.Equals(RCDateList[j]))
                        {

                            if (dateTime6.Equals(currentDate))
                            {
                                //MessageBox.Show("Reminder: Registration Card Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Today", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                ShowToast("REMINDER", "Reminder: Registration Card Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Today");
                                Expired3 = true;
                                RowList.Add(i);
                                ColList.Add(5);
                            }

                            if (Expired3 == false)
                            {
                                //MessageBox.Show("Reminder: Registration Card Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Soon", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                ShowToast("WARNING", "Reminder: Registration Card Expiry Date for " + dataGridView1.Rows[i].Cells[0].Value?.ToString() + " is Due Soon");
                                RowList2.Add(i);
                                ColList2.Add(6);
                            }
                      
                        }

                    }


                }
            }

            
        }

        public void ShowToast(string Type, string Message)
        {
            ToastForm toastForm = new ToastForm(Type, Message);
            toastForm.ShowDialog();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void label37_Click(object sender, EventArgs e)
        {

        }

        private void label34_Click(object sender, EventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }

        private void label30_Click(object sender, EventArgs e)
        {

        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

      

        

        private void OpenXLFile()
        {
            dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
            dynamic xlworksheet = xlworkbook.Worksheets["Reminders"];
            dynamic xlrange = xlworksheet.UsedRange;

            //dataGridView3.ColumnCount = xlrange.Columns.Count;
            dataGridView1.Rows.Clear();

            for (int xlrow = 2; xlrow <= xlrange.Rows.Count; xlrow++)
            {

                dataGridView1.Rows.Add(xlrange.Cells[xlrow, 1].Text, xlrange.Cells[xlrow, 2].Text,
                xlrange.Cells[xlrow, 3].Text, xlrange.Cells[xlrow, 4].Text, xlrange.Cells[xlrow, 5].Text,
                xlrange.Cells[xlrow, 6].Text, xlrange.Cells[xlrow, 7].Text);

            }

            DeleteEmptyRows(dataGridView1);

            xlworkbook.Close();
            xlapp.Quit();
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

        private void MainMenu_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.ShowDialog();
            this.Close();
        }

        private void SaveButton_Click_1(object sender, EventArgs e)
        {
            if (File.Exists(newFilePath))
            {
                // Open Excel and the workbook
                dynamic xlapp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                dynamic xlworkbook = xlapp.Workbooks.Open(newFilePath);
                dynamic xlworksheet = xlworkbook.Worksheets["Reminders"];

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

        private void Reset_Click_1(object sender, EventArgs e)
        {
            VehiclePlateNumberBox.Text = null;
            IEDDateBox.Text = null;
            IEDMonthBox.Text = null;
            IEDYearBox.Text = null;
            IEDReminderDateBox.Text = null;
            IEDReminderMonthBox.Text = null;
            IEDReminderYearBox.Text = null;
            TVEDDateBox.Text = null;
            TVEDMonthBox.Text = null;
            TVEDYearBox.Text = null;
            TVEDReminderDateBox.Text = null;
            TVEDReminderMonthBox.Text = null;
            TVEDReminderYearBox.Text = null;
            RCEDateBox.Text = null;
            RCEMonthBox.Text = null;
            RCEYearBox.Text = null;
            RCEReminderDateBox.Text = null;
            RCEReminderMonthBox.Text = null;
            RCEReminderYearBox.Text = null;

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
                    ShowToast("ERROR", "Rows are not able to be deleted.");
                }
            }
        }

        private void Delete_Click_1(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
            }
        }

        private void AddNew_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(VehiclePlateNumberBox.Text) ||
                string.IsNullOrEmpty(IEDDateBox.Text) ||
                string.IsNullOrEmpty(IEDMonthBox.Text) ||
                string.IsNullOrEmpty(IEDYearBox.Text) ||
                string.IsNullOrEmpty(IEDReminderDateBox.Text) ||
                string.IsNullOrEmpty(IEDReminderMonthBox.Text) ||
                string.IsNullOrEmpty(IEDReminderYearBox.Text) ||
                string.IsNullOrEmpty(TVEDDateBox.Text) ||
                string.IsNullOrEmpty(TVEDMonthBox.Text) ||
                string.IsNullOrEmpty(TVEDYearBox.Text) ||
                string.IsNullOrEmpty(TVEDReminderDateBox.Text) ||
                string.IsNullOrEmpty(TVEDReminderMonthBox.Text) ||
                string.IsNullOrEmpty(TVEDReminderYearBox.Text) ||
                string.IsNullOrEmpty(RCEDateBox.Text) ||
                string.IsNullOrEmpty(RCEMonthBox.Text) ||
                string.IsNullOrEmpty(RCEYearBox.Text) ||
                string.IsNullOrEmpty(RCEReminderDateBox.Text) ||
                string.IsNullOrEmpty(RCEReminderMonthBox.Text) ||
                string.IsNullOrEmpty(RCEReminderYearBox.Text))
            {
                //MessageBox.Show("All fields must be filled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowToast("ERROR", "All fields must be filled.");
            }
            else
            {
                if(CarPlateValidity(VehiclePlateNumberBox.Text) == true)
                {
                    dataGridView1.Rows.Add(VehiclePlateNumberBox.Text, IEDDateBox.Text + Slash + IEDMonthBox.Text + Slash + IEDYearBox.Text, IEDReminderDateBox.Text + Slash + IEDReminderMonthBox.Text + Slash + IEDReminderYearBox.Text, TVEDDateBox.Text + Slash + TVEDMonthBox.Text + Slash + TVEDYearBox.Text, TVEDReminderDateBox.Text + Slash + TVEDReminderMonthBox.Text + Slash + TVEDReminderYearBox.Text, RCEDateBox.Text + Slash + RCEMonthBox.Text + Slash + RCEYearBox.Text, RCEReminderDateBox.Text + Slash + RCEReminderMonthBox.Text + Slash + RCEReminderYearBox.Text);

                }

            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            for (int i = 0; i < RowList.Count; i++)
            {
                //MessageBox.Show("Row: " + RowList[i].ToString(), "Colored Rows");
                dataGridView1.Rows[RowList[i]].Cells[ColList[i]].Style.BackColor = Color.Red;
            }

            for (int i = 0; i < RowList2.Count; i++)
            {
                dataGridView1.Rows[RowList2[i]].Cells[ColList2[i]].Style.BackColor = Color.Orange;
            }
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
            catch(Exception e)
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
