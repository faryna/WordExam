using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Windows.Media;
using System.Drawing.Text;

namespace Word
{
    public partial class Form1 : Form
    {
        
        FileStream sr;
        string str;
        string fn;
        public Form1()
        {
            InitializeComponent();
            richTextBox1.AllowDrop = true;

            richTextBox1.DragDrop += RichTextBox_Drop;
            this.BackColor = Color.CornflowerBlue;
            fn = "";
            for (int i = 8; i < 48; i++)
            {
                if (i > 20)
                    i++;
                toolStripComboBox1.Items.Add(i); 
            }
            toolStripComboBox1.SelectedIndex = 0;
            InstalledFontCollection ifc = new InstalledFontCollection();
            FontFamily[] ff = ifc.Families;

            foreach (FontFamily fontFamily in ff)
            {
                toolStripComboBox2.Items.Add(fontFamily.Name);
            }
            toolStripComboBox2.SelectedItem = "Arial";
            toolStripButton1.BackColor = Color.Black;
            richTextBox1.HideSelection = false;
            richTextBox1.DragEnter += new DragEventHandler(Form1_DragEnter);
            toolStripStatusLabel1.Text = "Ln: 0 Col: 0";
            
            
        }
        public Form1(string FileName)
        {
            InitializeComponent();
            fn = FileName;
            this.BackColor = Color.CornflowerBlue;
            richTextBox1.Rtf = File.ReadAllText(FileName);
            this.Visible = true;
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.Filter = "rtf files (*.rtf)|*.rtf|TXT files (*.txt)|*.txt";
                openFileDialog1.ShowDialog();
                string str = openFileDialog1.FileName;
                Form1 nf = new Form1(str);
            }
            catch (Exception ex)
            {

            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (fn != "")
                {
                    File.WriteAllText(fn, richTextBox1.Text);
                }
                else
                {
                    saveFileDialog1.Filter = "rtf files (*.rtf)|*.rtf|TXT files (*.txt)|*.txt";
                    saveFileDialog1.ShowDialog();
                    File.WriteAllText(saveFileDialog1.FileName, richTextBox1.Rtf);
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ToolStripComboBox cb = sender as ToolStripComboBox;
            if (toolStripComboBox2.SelectedItem != null)
                richTextBox1.SelectionFont = new System.Drawing.Font(toolStripComboBox2.SelectedItem.ToString(), Convert.ToInt16(cb.SelectedItem));
            else
                richTextBox1.SelectionFont = new System.Drawing.Font("Arial", Convert.ToInt16(cb.SelectedItem));
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ToolStripComboBox cb = sender as ToolStripComboBox;
            richTextBox1.SelectionFont = new System.Drawing.Font(cb.SelectedItem.ToString(), this.Font.Size);
        }

        private void toolStripUndo_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void toolStripRedo_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        private void pasteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            //richTextBox1.ForeColor = colorDialog1.Color;
            toolStripButton1.BackColor = colorDialog1.Color;
            richTextBox1.SelectionColor = colorDialog1.Color;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Find f = new Find(ref richTextBox1);
            f.Show();
        }

        private void richTextBox1_SelectionChanged(object sender, EventArgs e)
        {
            System.Drawing.Point pntCursorPosition = new System.Drawing.Point();
            pntCursorPosition.X = richTextBox1.GetLineFromCharIndex(richTextBox1.SelectionStart);
            pntCursorPosition.Y = richTextBox1.SelectionStart - richTextBox1.GetFirstCharIndexOfCurrentLine();
            toolStripStatusLabel1.Text = "Ln: " + pntCursorPosition.X + " Col: " + pntCursorPosition.Y;
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            var lvi = sender as ListViewItem;
            MessageBox.Show(sender.ToString());
        }

        private void RichTextBox_Drop(object sender, DragEventArgs e)
        {
            string[] filenames = e.Data.GetData(DataFormats.FileDrop) as string[];

            if (filenames != null)
            {
                foreach (string name in filenames)
                {
                    try
                    {
                        richTextBox1.AppendText(File.ReadAllText(name));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            Color c = colorDialog1.Color;
            richTextBox1.SelectionColor = c;
        }

        private void toolStripComboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            richTextBox1.SelectionFont = new System.Drawing.Font(toolStripComboBox2.Selected.ToString(), Convert.ToInt32(toolStripComboBox1.Selected));
        }

    }
}
