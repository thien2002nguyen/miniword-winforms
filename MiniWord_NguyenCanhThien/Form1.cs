using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MiniWord_NguyenCanhThien
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            richTextBox1.ContextMenuStrip = contextMenuStrip1;
            this.Text = "Untitled";
            string[] listFont = System.Drawing.FontFamily.Families.Select(f => f.Name).ToArray();
            toolStripComboBox2.Items.AddRange(listFont);
            toolStripComboBox1.SelectedItem = "12";
            toolStripComboBox2.SelectedItem = "Arial";
            richTextBox1.Font = new Font("Arial", 12);
        }

        private void newFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void newWindowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form1 newForm = new Form1();
            newForm.Show();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Open";
            openFileDialog.Filter = "Text txtument(*.txt)|*.txt|All Files(*.*)|*.*";
            if(openFileDialog.ShowDialog() == DialogResult.OK)
            { 
                richTextBox1.LoadFile(openFileDialog.FileName, RichTextBoxStreamType.PlainText);
            }
            this.Text = openFileDialog.FileName;
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.Text) || !File.Exists(this.Text))
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "Save";
                saveFileDialog.Filter = "Text txtument(*.txt)|*.txt|All Files(*.*)|*.*";
                saveFileDialog.FileName = this.Text;
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    richTextBox1.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText);
                    this.Text = saveFileDialog.FileName;
                }
            }
            else
            {
                richTextBox1.SaveFile(this.Text, RichTextBoxStreamType.PlainText);
            }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog openFileDialog = new SaveFileDialog();
            openFileDialog.Title = "Save";
            openFileDialog.Filter = "Text txtument(*.txt)|*.txt|All Files(*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SaveFile(openFileDialog.FileName, RichTextBoxStreamType.PlainText);
            }
            this.Text = openFileDialog.FileName;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void dateTimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = System.DateTime.Now.ToString();
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        private void paseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectedText = string.Empty;
        }

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void replaceToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectAll();
        }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FontDialog fnt = new FontDialog();
            if(fnt.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SelectionFont = fnt.Font;
            }
        }

        private void colorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog fnt = new ColorDialog();
            if (fnt.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SelectionColor = fnt.Color;
            }
        }

        private void zoomInToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float newZoomFactor = richTextBox1.ZoomFactor * 1.2f;
            richTextBox1.ZoomFactor = newZoomFactor;
        }

        private void zoomOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float newZoomFactor = richTextBox1.ZoomFactor / 1.2f;
            richTextBox1.ZoomFactor = newZoomFactor;
        }

        private void restoreDefaultZoomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.ZoomFactor = 1.0f;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ToolStripComboBox comboBox = toolStripComboBox1;
            string selectedSize = comboBox.SelectedItem as string;
            if (!string.IsNullOrEmpty(selectedSize))
            {
                float fontSize;
                if (float.TryParse(selectedSize, out fontSize))
                {
                    Font currentFont = richTextBox1.Font;
                    Font newFont = new Font(currentFont.FontFamily, fontSize, currentFont.Style);
                    richTextBox1.Font = newFont;
                }
            }
        }

        private void toolStripComboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                ToolStripComboBox toolStripComboBox = (ToolStripComboBox)sender;
                string enteredSizeText = toolStripComboBox.Text;
                if (float.TryParse(enteredSizeText, out float fontSize))
                {
                    Font currentFont = richTextBox1.Font;
                    richTextBox1.Font = new Font(currentFont.FontFamily, fontSize, currentFont.Style);
                    richTextBox1.Focus();
                }
            }
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            RichTextBox richTextBox = richTextBox1;
            Font currentFont = richTextBox1.Font;
            string selectedFontFamily = toolStripComboBox2.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(selectedFontFamily))
            {
                Font newFont = new Font(selectedFontFamily, currentFont.Size, richTextBox.Font.Style);
                richTextBox.Font = newFont;
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp|All Files|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string imagePath = openFileDialog.FileName;
                    InsertImage(imagePath, new Size(200, 200));
                }
            }
        }

        private void InsertImage(string imagePath, Size newSize)
        {
            if (File.Exists(imagePath))
            {
                Image originalImage = Image.FromFile(imagePath);
                Image resizedImage = ResizeImage(originalImage, newSize);
                Clipboard.SetImage(resizedImage);
                DataFormats.Format format = DataFormats.GetFormat(DataFormats.Bitmap);
                richTextBox1.Paste(format);
                Clipboard.Clear();
            }
        }

        private Image ResizeImage(Image image, Size newSize)
        {
            Bitmap resizedBitmap = new Bitmap(newSize.Width, newSize.Height);
            using (Graphics g = Graphics.FromImage(resizedBitmap))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(image, new Rectangle(Point.Empty, newSize));
            }
            return resizedBitmap;
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            UpdateFontStyle(FontStyle.Bold, toolStripButton8);
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            UpdateFontStyle(FontStyle.Italic, toolStripButton7);
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            UpdateFontStyle(FontStyle.Underline, toolStripButton6);
        }
        private void UpdateFontStyle(FontStyle style, ToolStripButton button)
        {
            Font currentFont = richTextBox1.SelectionFont;
            bool isApplied = (currentFont.Style & style) == style;
            FontStyle newFontStyle = isApplied ? currentFont.Style & ~style : currentFont.Style | style;
            Font newFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
            richTextBox1.SelectionFont = newFont;
            button.Checked = !isApplied;
        }
    }
}
