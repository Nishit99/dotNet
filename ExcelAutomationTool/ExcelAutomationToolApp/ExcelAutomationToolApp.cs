using System;
using System.Drawing;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelAutomationToolLibrary;

namespace ExcelAutomationToolApp
{
    public partial class Frm_AutomationTool : Form
    {
        readonly ToolHelperClass toolHelperClass = ToolHelperClass.GetWorkBookInstance;
        private Logger logger = null;

        /// <summary>
        /// Constructor used to initialize objects.
        /// </summary>
        public Frm_AutomationTool()
        {
            InitializeComponent();
            logger = Logger.Instance;

            txt_SourceFileName.Text = @"C:\Nishit\Temp\RA_Tool_Reports\DummyApplicationByReadinessReport07Nov22v1.0.xlsx";
            txt_ReferenceFilePath_URL.Text = "https://eu001-sp.shell.com/:x:/r/sites/UGGFAppsProgramsProjectsCollaboration/Shared%20Documents/UEM%20Application%20Testing%20and%20Readiness/App%20List/App%20Baseline/PowerBI_Dashboard/Baseline_Reference_Sheet.xlsx?d=wc747dbb2afb04b8e85ac623a2a0a8438&csf=1&web=1&e=ayhads";
        }

        #region "Events"

        /// <summary>
        /// Occurs when the control is clicked.
        /// 1. Validate all the required fields.
        /// 2. Call the RunExcelAutomationTool method from helper class and paas the required paramters.
        /// 3. Show to success / failed message.
        /// 4. log the error in error log.
        /// </summary>
        /// <param name="sender">Object type. The source of the event.</param>
        /// <param name="e">EventArgs, An object that contains no event data</param>
        private void btnRun_Click(object sender, EventArgs e)
        {
            //Validate all the required fields.
            //if (!ValidateChildren(ValidationConstraints.Enabled))
            //{
            //    return;
            //}

            try
            {
                //Call the RunExcelAutomationTool method from helper class and paas the required paramters.
                toolHelperClass.RunExcelAutomationTool(txt_SourceFileName.Text, txt_ReferenceFilePath_URL.Text, lstBox_Worksheets.SelectedItem.ToString());

                //Show to success message.
                toolStripStatusLabel.ForeColor = Color.DarkGreen;
                toolStripStatusLabel.Text = "Completed!!";
            }
            catch (Exception ex)
            {
                //Show to failed message.
                toolStripStatusLabel.ForeColor = Color.Red;
                toolStripStatusLabel.Text = "Failed....";

                //log the error in error log.
                logger.LogMessage(ex.Message);
            }
        }

        /// <summary>
        /// Occurs when the control is clicked.
        /// 1. Open the file Browser. to select file to be updated.
        /// </summary>
        /// <param name="sender">Object type. The source of the event.</param>
        /// <param name="e">EventArgs, An object that contains no event data</param>
        private void btn_BrowseFile_Click(object sender, EventArgs e)
        {
            //Open the file Browser. to select file to be updated.
            oFD_SourceFile.RestoreDirectory = true;
            oFD_SourceFile.ShowDialog();
            txt_SourceFileName.Text = oFD_SourceFile.FileName;
        }

        /// <summary>
        /// Occurs when the control is clicked.
        /// 1. Close the Excel Automation Tool app.
        /// </summary>
        /// <param name="sender">Object type. The source of the event.</param>
        /// <param name="e">EventArgs, An object that contains no event data</param>
        private void btn_Close_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// Occurs when the Text property value changes
        /// 1. source (a file from where the valued will be comared and picked.) file/URL path.
        /// 2. Add all the list of worksheet from the select source file to ListBox.
        /// </summary>
        /// <param name="sender">Object type. The source of the event.</param>
        /// <param name="e">EventArgs, An object that contains no event data</param>
        private void txt_ReferenceFilePath_URL_TextChanged(object sender, EventArgs e)
        {
            lstBox_Worksheets.Items.Clear();

            //Source file (a file from where the values will be compared and picked.) file/URL path.
            List<string> lstWorksheets = toolHelperClass.GetAllWorksheets(txt_ReferenceFilePath_URL.Text);

            //Add all the list of worksheet from the select source file to ListBox.
            lstWorksheets.ForEach(Sheet => lstBox_Worksheets.Items.Add(Sheet));
        }

        /// <summary>
        /// Occurs when the SelectedIndex property has changed.
        /// 1. Show/hide Xlookup and IF panel base on selection.
        /// </summary>
        /// <param name="sender">Object type. The source of the event.</param>
        /// <param name="e">EventArgs, An object that contains no event data</param>
        private void cmb_SelectFormula_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Show/hide Xlookup and IF panel base on selection.
            switch (cmb_SelectFormula.SelectedIndex)
            {
                case 0:
                    pnl_XLookup.Visible = true;
                    pnl_If.Visible = false;
                    break;
                case 1:
                    pnl_XLookup.Visible = false;
                    pnl_If.Location = pnl_XLookup.Location;
                    pnl_If.Visible = true;
                    break;
                default:
                    pnl_XLookup.Visible = false;
                    pnl_If.Visible = false;
                    break;
            }

            if (cmb_SelectFormula.SelectedIndex == 0)
            {

            }
        }

        /// <summary>
        /// Occurs when the control is validating.
        /// 1. Validate ( check for null or empty) all the required form control.
        /// 2. Set foucus on currect control and show the error icon if validation fails.
        /// </summary>
        /// <param name="sender">Object type. The source of the event.</param>
        /// <param name="e">EventArgs, An object that contains no event data</param>
        private void FormControl_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            dynamic ControlComponent = null;

            if (sender is TextBox)
            {
                ControlComponent = (TextBox)sender;
            }
            else if (sender is ComboBox)
            {
                ControlComponent = (ComboBox)sender;
            }

            if (ControlComponent != null)
            {
                if (string.IsNullOrWhiteSpace(ControlComponent.Text))
                {
                    //Set foucus on currect control and show the error icon if validation fails.
                    e.Cancel = true;
                    ControlComponent.Focus();
                    errorProvider.SetError(ControlComponent, "Message");
                }
                else
                {
                    e.Cancel = false;
                    errorProvider.SetError(ControlComponent, "");
                }
            }
        }

        #endregion


    }
}
