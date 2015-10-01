using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Word
{
    public partial class Find : Form
    {
        private RichTextBox richTextBox;
        private int index;
        public Find()
        {
            InitializeComponent();

            index = 0;
        }
        public Find(ref RichTextBox rtb)
        {
            InitializeComponent();
            richTextBox = rtb;
            index = 0;
        }
        private void buttonFind_Click(object sender, EventArgs e)
        {
           index = richTextBox.Find(textBox1.Text, index, RichTextBoxFinds.None) + textBox1.Text.Length;
           MessageBox.Show(index.ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            index = richTextBox.Find(textBox2.Text, index, RichTextBoxFinds.None) + textBox2.Text.Length;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (richTextBox.SelectedText != "")
            {
                richTextBox.SelectedText = textBox3.Text;
            }
            else
            {
                index = richTextBox.Find(textBox2.Text, index, RichTextBoxFinds.None) + textBox2.Text.Length;
                if (richTextBox.SelectedText != "")
                {
                    richTextBox.SelectedText = textBox3.Text;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            do
            {
                index = richTextBox.Find(textBox2.Text, index, RichTextBoxFinds.None) + textBox2.Text.Length;
                if (richTextBox.SelectedText != "")
                {
                    richTextBox.SelectedText = textBox3.Text;
                    index = richTextBox.Find(textBox2.Text, index, RichTextBoxFinds.None) + textBox2.Text.Length;
                }
            } while (richTextBox.SelectedText != "");
        }
    }
}
