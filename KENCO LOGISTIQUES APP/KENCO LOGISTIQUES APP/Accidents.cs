using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KENCO_LOGISTIQUES_APP
{
    public partial class Accidents : Form
    {
        public Accidents()
        {
            InitializeComponent();
            lblDateToday.Text = DateTime.Now.ToLongDateString();
        }


        private void Print(Panel pnl)
        {
            PrinterSettings ps = new PrinterSettings();
            panelPrint = pnl;
            getprintarea(pnl);
            printPreviewDialog1.Document = printDocument1;
            printDocument1.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);
            printPreviewDialog1.ShowDialog();
        }

        private Bitmap memorying;

        private void getprintarea(Panel pnl)
        {
            memorying = new Bitmap(pnl.Width, pnl.Height);
            pnl.DrawToBitmap(memorying, new Rectangle(0, 0, pnl.Width, pnl.Height));
        }

        private Image ResizeImage(Image image, int width, int height)
        {
            // Create a new Bitmap with the desired width and height
            Bitmap resizedBitmap = new Bitmap(width, height);

            // Create a Graphics object from the new Bitmap
            using (Graphics graphics = Graphics.FromImage(resizedBitmap))
            {
                // Set the interpolation mode to high quality
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                // Draw the original image onto the new Bitmap with the desired size
                graphics.DrawImage(image, 0, 0, width, height);
            }

            // Return the resized image
            return resizedBitmap;
        }

        private void label17_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Image Files (*.bmp;*.jpg;*.jpeg;*.png;*.gif)|*.bmp;*.jpg;*.jpeg;*.png;*.gif";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedImagePath = openFileDialog.FileName;

                // Load the original image
                Image originalImage = Image.FromFile(selectedImagePath);

                // Resize the image
                int desiredWidth = pictureBox1.Width; // Set the desired width to match PictureBox width
                int desiredHeight = pictureBox1.Height; // Set the desired height to match PictureBox height
                Image resizedImage = ResizeImage(originalImage, desiredWidth, desiredHeight);

                // Display the resized image in the PictureBox
                pictureBox1.Image = resizedImage;
            }
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            Rectangle pagearea = e.PageBounds;
            e.Graphics.DrawImage(memorying, pagearea);
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            Print(panelPrint);
        }

        private void MainMenu_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.ShowDialog();
            this.Close();
        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void panelPrint_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
