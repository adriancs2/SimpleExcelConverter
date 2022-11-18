using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace ExcelDocConverter
{
    public partial class Form2 : Form
    {
        ExcelHelper.ExcelDocConverter excel = null;
        BackgroundWorker bw = new BackgroundWorker();
        Dictionary<int, CurFile> dic = new Dictionary<int, CurFile>();
        bool cancelProcess = false;
        ExcelHelper.ExcelDocConverter.FormatType formatType = ExcelHelper.ExcelDocConverter.FormatType.XLS;

        public Form2()
        {
            InitializeComponent();
            bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
            bw.WorkerReportsProgress = true;
            bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            dataGridView1.CellContentClick += new DataGridViewCellEventHandler(dataGridView1_CellContentClick);
            dataGridView1.CellDoubleClick += new DataGridViewCellEventHandler(dataGridView1_CellDoubleClick);
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void btOutputFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog f = new FolderBrowserDialog();
            f.ShowNewFolderButton = true;
            f.Description = "Select a folder to save the converted files..";
            if (f.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtOutput.Text = f.SelectedPath;
            }
        }

        private void btSelectFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Excel|*.xls;*.xlsx";
            of.Multiselect = true;
            if (of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string s in of.FileNames)
                {
                    bool added = false;
                    foreach (DataGridViewRow r in dataGridView1.Rows)
                    {
                        if (r.Cells[colnFiles.Index].Value + "" == s)
                        {
                            added = true;
                            break;
                        }
                    }

                    if (!added)
                    {
                        DataGridViewRow dgvr = dataGridView1.Rows[dataGridView1.Rows.Add()];
                        dgvr.Cells[colnSelect.Index].Value = "Remove";
                        dgvr.Cells[colnFiles.Index].Value = s;
                        dgvr.Cells[colnStatus.Index].Value = "";
                    }
                }
            }
        }

        private void btStartConvert_Click(object sender, EventArgs e)
        {
            if (btStartConvert.Text == "Cancel")
            {
                cancelProcess = true;
                pictureBox1.Visible = false;
                btStartConvert.Text = "Start Convert";
                return;
            }
            else
            {
                if (cbFileType.SelectedIndex < 0)
                {
                    MessageBox.Show("Select a target file type that you wish to convert to.");
                    return;
                }

                if (!Directory.Exists(txtOutput.Text))
                {
                    try
                    {
                        Directory.CreateDirectory(txtOutput.Text);
                    }
                    catch
                    { }
                    if (!Directory.Exists(txtOutput.Text))
                    {
                        MessageBox.Show("The output folder is not existed. Please select a folder to save the converted files.");
                        return;
                    }
                    else
                    {
                        DirectoryInfo di = new DirectoryInfo(txtOutput.Text);
                        txtOutput.Text = di.FullName;
                    }
                }
                btStartConvert.Text = "Cancel";
            }
            
            lbError.Visible = false;
            string fileExtension = GetFileExtension();
            string outputfolder = txtOutput.Text;
            dic = new Dictionary<int, CurFile>();
            formatType = GetFileType();

            foreach (DataGridViewRow dgvr in dataGridView1.Rows)
            {
                excel = new ExcelHelper.ExcelDocConverter();
                string file = dgvr.Cells[colnFiles.Index].Value + "";

                FileInfo fi = new FileInfo(file);

                string targetFile = txtOutput.Text + "\\" + fi.Name.Replace(fi.Extension, string.Empty) + fileExtension;
                CurFile cf = new CurFile();
                cf.Index = dgvr.Index;
                cf.OriFile = file;
                cf.TargetFile = targetFile;
                dic[dgvr.Index] = cf;

                dgvr.Cells[colnStatus.Index].Value = "";
            }
            //curIndex = 0;
            cancelProcess = false;
            pictureBox1.Visible = true;
            bw.RunWorkerAsync();
        }

        string GetFileExtension()
        {
            switch (cbFileType.SelectedIndex)
            {
                case 0:
                    return ".xls";
                case 1:
                    return ".xlsx";
                case 2:
                    return ".pdf";
                case 3:
                    return ".xps";
                case 4:
                    return ".csv";
                default:
                    throw new Exception("Unknown File Type");
            }
        }

        ExcelHelper.ExcelDocConverter.FormatType GetFileType()
        {
            switch (cbFileType.SelectedIndex)
            {
                case 0:
                    return ExcelHelper.ExcelDocConverter.FormatType.XLS;
                case 1:
                    return ExcelHelper.ExcelDocConverter.FormatType.XLSX;
                case 2:
                    return ExcelHelper.ExcelDocConverter.FormatType.PDF;
                case 3:
                    return ExcelHelper.ExcelDocConverter.FormatType.XPS;
                case 4:
                    return ExcelHelper.ExcelDocConverter.FormatType.CSV;
                default:
                    throw new Exception("Unknown File Type");
            }
        }

        private void btRemoveAll_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you to remove all selected files?", "Remove All", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                dataGridView1.Rows.Clear();
            }
        }

        void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            foreach (KeyValuePair<int, CurFile> kv in dic)
            {
                if (cancelProcess)
                    return;

                CurFile cf = kv.Value;
                try
                {
                    excel = new ExcelHelper.ExcelDocConverter();
                    excel.Convert(formatType, cf.OriFile, cf.TargetFile);
                    excel = null;
                    cf.SuccessConvert = true;
                }
                catch (Exception ex)
                {
                    cf.SuccessConvert = false;
                    cf.StatusMessage = "Error";
                    cf.ErrorMessage = ex.Message + "\r\n\r\nClick [More Info] for further information.";
                }
                bw.ReportProgress(cf.Index);
            }
        }

        void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (cancelProcess)
            {
                foreach (DataGridViewRow dgvr in dataGridView1.Rows)
                {
                    if (dgvr.Cells[colnStatus.Index].Value + "" == "")
                    {
                        dgvr.Cells[colnStatus.Index].Value = "Canceled";
                    }
                }
                pictureBox1.Visible = false;
                this.Refresh();
                MessageBox.Show("Canceled", "Cancel");
            }
            else
            {
                pictureBox1.Visible = false;
                this.Refresh();
                MessageBox.Show("Finished.");
            }
            btStartConvert.Text = "Start Convert";
            bool hasError = false;
            foreach (DataGridViewRow dgvr in dataGridView1.Rows)
            {
                dgvr.Tag = dic[dgvr.Index];
                if (!dic[dgvr.Index].SuccessConvert)
                {
                    hasError = true;
                }
            }
            if (hasError)
            {
                lbError.Visible = true;
            }
            else
                lbError.Visible = false;
        }

        void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            CurFile cf = dic[e.ProgressPercentage];
            dataGridView1.Rows[e.ProgressPercentage].Cells[colnStatus.Index].Value = dic[e.ProgressPercentage].StatusMessage;
            if (cf.SuccessConvert)
            {
                dataGridView1.Rows[e.ProgressPercentage].Cells[colnStatus.Index].Value = "Completed";
                dataGridView1.Rows[e.ProgressPercentage].Cells[colnStatus.Index].Style.ForeColor = Color.DarkGreen;
                dataGridView1.Rows[e.ProgressPercentage].Cells[colnStatus.Index].Style.SelectionForeColor = Color.DarkGreen;
            }
            else
            {
                dataGridView1.Rows[e.ProgressPercentage].Cells[colnStatus.Index].Value = "Error";
                dataGridView1.Rows[e.ProgressPercentage].Cells[colnStatus.Index].Style.ForeColor = Color.Red;
                dataGridView1.Rows[e.ProgressPercentage].Cells[colnStatus.Index].Style.SelectionForeColor = Color.Red;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string msg = "Simple Microsoft Excel Documents Converter 2.1\r\nhttps://github.com/adriancs2/SimpleExcelConverter\r\nFreeware.\r\n\r\nThe computer must have Excel 2007 or newer installed for this software to work.\r\n\r\nExcel 2007 support conversion to:\r\n- XLS\r\n- XLSX\r\n- CSV\r\n\r\nExcel 2010 / 2012 / 2013 supports conversion to:\r\n- XLS\r\n- XLSX\r\n- PDF\r\n- XPS\r\n- CSV";
            Form1 f = new Form1(msg, "About");
            f.Height = 400;
            f.ShowDialog();
        }

        void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != colnSelect.Index)
            { }
            else
            {
                dataGridView1.Rows.RemoveAt(e.RowIndex);
            }
        }

        void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            if (e.ColumnIndex == colnStatus.Index)
            {
                if (dataGridView1.Rows[e.RowIndex].Tag == null)
                    return;

                if (dataGridView1.Rows[e.RowIndex].Tag.GetType() != typeof(CurFile))
                    return;

                CurFile cf = (CurFile)dataGridView1.Rows[e.RowIndex].Tag;
                Form1 f =null;
                if (cf.SuccessConvert)
                {
                    f = new Form1("File converted successfully. No error found.", "No Error");
                }
                else
                {
                    if (cf.ErrorMessage.ToLower().Contains("value does not fall within the expected range"))
                    {
                        f = new Form1("This computer does not has Excel 2010 or newer installed. It is required for converting to selected file format.", "Error Message");
                    }
                    else if (cf.ErrorMessage.Contains("Could not load file or assembly 'Microsoft.Office.Interop.Excel, Version=14.0.0.0"))
                    {
                        f = new Form1("This computer does not has Excel 2007 or newer installed.", "Error Message");
                    }
                    else
                    {
                        f = new Form1(cf.ErrorMessage, "Error Message");
                    }
                }
                f.ShowDialog();
            }
        }

    }

    public class CurFile
    {
        public int Index = -1;
        public string OriFile = "";
        public string TargetFile = "";
        public bool SuccessConvert = false;
        public string ErrorMessage = "";
        public string StatusMessage = "";
    }
}
