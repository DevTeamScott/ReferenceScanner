using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace ReferenceDataInput
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void CSVOutput_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            string outputLocation = Properties.Settings.Default.OutputLocation;

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ReferenceScanner.ReferenceCheckedData ScanforRefs = new ReferenceScanner.ReferenceCheckedData();

                string filePath = openFileDialog.ToString();
                string fileName = Path.GetFileName(filePath);
                string fileLocation = outputLocation + "\\" + fileName;

                List<string> headers = new List<string>();

                headers.Add("Reference");
                headers.Add("name 1");
                headers.Add("name 2");

                string cleanedFile = "cleaned_" + fileName;

                ScanforRefs.ReturnReferenceScannedData(fileLocation, ReferenceScanner.OutputFormat.CSV, headers, outputLocation, cleanedFile);

            }
        }


        private void ExlOutput_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            string outputLocation = Properties.Settings.Default.OutputLocation;

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                ReferenceScanner.ReferenceCheckedData ScanforRefs = new ReferenceScanner.ReferenceCheckedData();

                string filePath = openFileDialog.ToString();
                string fileName = Path.GetFileName(filePath);
                string fileLocation = outputLocation + "\\" + fileName;

                List<string> headers = new List<string>();

                headers.Add("Reference");
                headers.Add("name 1");
                headers.Add("name 2");

                string sheetName = "EBU Daily";

                ScanforRefs.ReturnReferenceScannedData(fileLocation, ReferenceScanner.OutputFormat.Excel, headers, outputLocation, fileName, sheetName);

            }
        }

        private void DtOutput_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            string outputLocation = Properties.Settings.Default.OutputLocation;

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string filePath = openFileDialog.ToString();
                string fileName = Path.GetFileName(filePath);
                string fileLocation = outputLocation + "\\" + fileName;

                List<string> headers = new List<string>();

                headers.Add("Reference");
                headers.Add("name 1");
                headers.Add("name 2");

                ReferenceScanner.ReferenceCheckedData speedCheck = new ReferenceScanner.ReferenceCheckedData();

                DataTable checkedData = speedCheck.ReturnReferenceScannedDataTable(fileLocation, ReferenceScanner.OutputFormat.DataTable, ReferenceScanner.FileType.csv, headers);

                dataGridView1.DataSource = checkedData;

            }
        }
    }
}

