using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SExtensions
{
    public partial class ExcelWaitingForm : Form
    {
        //public Action Worker { get; set; }

        public bool ExportRutas { get; set; }
        public bool Utillaje { get; set; }
        public bool GetPwdFiles { get; set; }
        public ExcelWaitingForm(/*Action worker*/)
        {
            
            InitializeComponent();

            

            //if (worker == null)
            //{
            //    throw new ArgumentNullException();
            //}

            //Worker = worker;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

        }

        private void StartAction()
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                return;
            }

            Task.Factory.StartNew(() => 
            {
                Helpers.FindOccurrencesAndExport(GetPwdFiles, ExportRutas, Utillaje, textBox1.Text);

            }).ContinueWith(t => { this.Close(); }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void Button_export_Click(object sender, EventArgs e)
        {

            StartAction();
        }

        private void Button_output_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
