﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DesignManager
{
    public partial class FormRename : Form
    {
        public FormRename()
        {
            InitializeComponent();
        }
        public bool CreateCopy { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox1.Text))
            {
                if (CreateCopy)
                {
                    Helpers.DesignManagerHelpers.CreateCopy(textBox1.Text);
                }
                else
                {
                    Helpers.DesignManagerHelpers.Rename(textBox1.Text);
                }
                
                Close();
            }
            
        }
    }
}
