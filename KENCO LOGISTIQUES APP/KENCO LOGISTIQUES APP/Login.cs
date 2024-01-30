namespace KENCO_LOGISTIQUES_APP
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int CorrectCounter = 0;

            if(textBox1.Text != "ADMIN")
            {
                label4.Visible = true;
            }
            else
            {
                label4.Visible = false;
                CorrectCounter++;
            }

            if (maskedTextBox1.Text != "KENCOL123")
            {
                label5.Visible = true;
            }
            else
            {
                label5.Visible = false;
                CorrectCounter++;
            }

            if(CorrectCounter == 2)
            {
                this.Hide();
                Menu form2 = new Menu();
                form2.ShowDialog();
                this.Close();
            }
        }
    }
}