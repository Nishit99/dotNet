
namespace ExcelAutomationToolApp
{
    partial class Frm_AutomationTool
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.oFD_SourceFile = new System.Windows.Forms.OpenFileDialog();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.grpBoxCoulmn = new System.Windows.Forms.GroupBox();
            this.btn_BrowseFile = new System.Windows.Forms.Button();
            this.txt_SourceFileName = new System.Windows.Forms.TextBox();
            this.lbl_SourceFileName = new System.Windows.Forms.Label();
            this.lbl_SelectWorksheet = new System.Windows.Forms.Label();
            this.lstBox_Worksheets = new System.Windows.Forms.ListBox();
            this.btn_Close = new System.Windows.Forms.Button();
            this.txt_ReferenceFilePath_URL = new System.Windows.Forms.TextBox();
            this.lbl_ReferenceFilePath_URL = new System.Windows.Forms.Label();
            this.btnRun = new System.Windows.Forms.Button();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.chkBox_BetweenColumns = new System.Windows.Forms.CheckBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.lblCoulmnName = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lbl_before = new System.Windows.Forms.Label();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.chkBoxAddFormulaInCell = new System.Windows.Forms.CheckBox();
            this.cmb_SelectFormula = new System.Windows.Forms.ComboBox();
            this.lbl_SelectFromula = new System.Windows.Forms.Label();
            this.pnl_XLookup = new System.Windows.Forms.Panel();
            this.cmb_Search_Mode = new System.Windows.Forms.ComboBox();
            this.cmb_Match_Mode = new System.Windows.Forms.ComboBox();
            this.txt_IfNotFound = new System.Windows.Forms.TextBox();
            this.txt_ReturnArray = new System.Windows.Forms.TextBox();
            this.txt_LookupArray = new System.Windows.Forms.TextBox();
            this.txt_LookupValue = new System.Windows.Forms.TextBox();
            this.lbl_Search_Mode = new System.Windows.Forms.Label();
            this.lbl_Match_Mode = new System.Windows.Forms.Label();
            this.lbl_if_not_found = new System.Windows.Forms.Label();
            this.lbl_ReturnArray = new System.Windows.Forms.Label();
            this.lbl_LookupArray = new System.Windows.Forms.Label();
            this.lbl_LookupValue = new System.Windows.Forms.Label();
            this.pnl_If = new System.Windows.Forms.Panel();
            this.txtFalseValue = new System.Windows.Forms.TextBox();
            this.txtTrueValue = new System.Windows.Forms.TextBox();
            this.txt_LogicValue = new System.Windows.Forms.TextBox();
            this.lbl_FalseValue = new System.Windows.Forms.Label();
            this.lbl_TrueValue = new System.Windows.Forms.Label();
            this.lbl_LogicValue = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.Button();
            this.listBox = new System.Windows.Forms.ListBox();
            this.statusStrip1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.grpBoxCoulmn.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.pnl_XLookup.SuspendLayout();
            this.pnl_If.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar1,
            this.toolStripStatusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 648);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1140, 23);
            this.statusStrip1.SizingGrip = false;
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 15);
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.toolStripStatusLabel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(0, 17);
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.grpBoxCoulmn);
            this.panel2.Controls.Add(this.btn_BrowseFile);
            this.panel2.Controls.Add(this.txt_SourceFileName);
            this.panel2.Controls.Add(this.lbl_SourceFileName);
            this.panel2.Controls.Add(this.lbl_SelectWorksheet);
            this.panel2.Controls.Add(this.lstBox_Worksheets);
            this.panel2.Controls.Add(this.btn_Close);
            this.panel2.Controls.Add(this.txt_ReferenceFilePath_URL);
            this.panel2.Controls.Add(this.lbl_ReferenceFilePath_URL);
            this.panel2.Controls.Add(this.btnRun);
            this.panel2.Location = new System.Drawing.Point(7, 8);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1127, 658);
            this.panel2.TabIndex = 11;
            // 
            // grpBoxCoulmn
            // 
            this.grpBoxCoulmn.Controls.Add(this.listBox);
            this.grpBoxCoulmn.Controls.Add(this.btnAdd);
            this.grpBoxCoulmn.Controls.Add(this.splitContainer2);
            this.grpBoxCoulmn.Controls.Add(this.splitContainer1);
            this.grpBoxCoulmn.Location = new System.Drawing.Point(6, 318);
            this.grpBoxCoulmn.Name = "grpBoxCoulmn";
            this.grpBoxCoulmn.Size = new System.Drawing.Size(1109, 293);
            this.grpBoxCoulmn.TabIndex = 27;
            this.grpBoxCoulmn.TabStop = false;
            this.grpBoxCoulmn.Text = "Insert Column and Formula";
            // 
            // btn_BrowseFile
            // 
            this.btn_BrowseFile.Location = new System.Drawing.Point(993, 19);
            this.btn_BrowseFile.Margin = new System.Windows.Forms.Padding(0);
            this.btn_BrowseFile.Name = "btn_BrowseFile";
            this.btn_BrowseFile.Size = new System.Drawing.Size(110, 34);
            this.btn_BrowseFile.TabIndex = 2;
            this.btn_BrowseFile.Text = "Browse File";
            this.btn_BrowseFile.UseVisualStyleBackColor = true;
            this.btn_BrowseFile.Click += new System.EventHandler(this.btn_BrowseFile_Click);
            // 
            // txt_SourceFileName
            // 
            this.txt_SourceFileName.Location = new System.Drawing.Point(113, 19);
            this.txt_SourceFileName.Name = "txt_SourceFileName";
            this.txt_SourceFileName.Size = new System.Drawing.Size(861, 22);
            this.txt_SourceFileName.TabIndex = 1;
            this.txt_SourceFileName.WordWrap = false;
            // 
            // lbl_SourceFileName
            // 
            this.lbl_SourceFileName.Location = new System.Drawing.Point(3, 19);
            this.lbl_SourceFileName.Name = "lbl_SourceFileName";
            this.lbl_SourceFileName.Size = new System.Drawing.Size(85, 34);
            this.lbl_SourceFileName.TabIndex = 26;
            this.lbl_SourceFileName.Text = "Source File:";
            // 
            // lbl_SelectWorksheet
            // 
            this.lbl_SelectWorksheet.Location = new System.Drawing.Point(3, 182);
            this.lbl_SelectWorksheet.Name = "lbl_SelectWorksheet";
            this.lbl_SelectWorksheet.Size = new System.Drawing.Size(200, 16);
            this.lbl_SelectWorksheet.TabIndex = 15;
            this.lbl_SelectWorksheet.Text = "Select Worksheet:";
            // 
            // lstBox_Worksheets
            // 
            this.lstBox_Worksheets.FormattingEnabled = true;
            this.lstBox_Worksheets.ItemHeight = 16;
            this.lstBox_Worksheets.Location = new System.Drawing.Point(221, 180);
            this.lstBox_Worksheets.Name = "lstBox_Worksheets";
            this.lstBox_Worksheets.Size = new System.Drawing.Size(293, 132);
            this.lstBox_Worksheets.TabIndex = 4;
            this.lstBox_Worksheets.Validating += new System.ComponentModel.CancelEventHandler(this.FormControl_Validating);
            // 
            // btn_Close
            // 
            this.btn_Close.Location = new System.Drawing.Point(993, 619);
            this.btn_Close.Name = "btn_Close";
            this.btn_Close.Size = new System.Drawing.Size(110, 34);
            this.btn_Close.TabIndex = 13;
            this.btn_Close.Text = "Close";
            this.btn_Close.UseVisualStyleBackColor = true;
            this.btn_Close.Click += new System.EventHandler(this.btn_Close_Click);
            // 
            // txt_ReferenceFilePath_URL
            // 
            this.txt_ReferenceFilePath_URL.Location = new System.Drawing.Point(221, 69);
            this.txt_ReferenceFilePath_URL.Multiline = true;
            this.txt_ReferenceFilePath_URL.Name = "txt_ReferenceFilePath_URL";
            this.txt_ReferenceFilePath_URL.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.txt_ReferenceFilePath_URL.Size = new System.Drawing.Size(469, 101);
            this.txt_ReferenceFilePath_URL.TabIndex = 3;
            this.txt_ReferenceFilePath_URL.TextChanged += new System.EventHandler(this.txt_ReferenceFilePath_URL_TextChanged);
            this.txt_ReferenceFilePath_URL.Validating += new System.ComponentModel.CancelEventHandler(this.FormControl_Validating);
            // 
            // lbl_ReferenceFilePath_URL
            // 
            this.lbl_ReferenceFilePath_URL.AutoSize = true;
            this.lbl_ReferenceFilePath_URL.Location = new System.Drawing.Point(3, 69);
            this.lbl_ReferenceFilePath_URL.Name = "lbl_ReferenceFilePath_URL";
            this.lbl_ReferenceFilePath_URL.Size = new System.Drawing.Size(200, 16);
            this.lbl_ReferenceFilePath_URL.TabIndex = 10;
            this.lbl_ReferenceFilePath_URL.Text = "Reference Workbook Path/URL:";
            this.lbl_ReferenceFilePath_URL.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(860, 619);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(110, 34);
            this.btnRun.TabIndex = 12;
            this.btnRun.Text = "Run";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // splitContainer1
            // 
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(18, 31);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.chkBox_BetweenColumns);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.textBox2);
            this.splitContainer1.Panel2.Controls.Add(this.lblCoulmnName);
            this.splitContainer1.Panel2.Controls.Add(this.textBox1);
            this.splitContainer1.Panel2.Controls.Add(this.lbl_before);
            this.splitContainer1.Size = new System.Drawing.Size(599, 75);
            this.splitContainer1.SplitterDistance = 171;
            this.splitContainer1.TabIndex = 25;
            // 
            // chkBox_BetweenColumns
            // 
            this.chkBox_BetweenColumns.AutoSize = true;
            this.chkBox_BetweenColumns.Location = new System.Drawing.Point(12, 9);
            this.chkBox_BetweenColumns.Name = "chkBox_BetweenColumns";
            this.chkBox_BetweenColumns.Size = new System.Drawing.Size(136, 20);
            this.chkBox_BetweenColumns.TabIndex = 21;
            this.chkBox_BetweenColumns.Text = "Between Columns";
            this.chkBox_BetweenColumns.UseVisualStyleBackColor = true;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(114, 39);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(238, 22);
            this.textBox2.TabIndex = 23;
            // 
            // lblCoulmnName
            // 
            this.lblCoulmnName.Location = new System.Drawing.Point(7, 39);
            this.lblCoulmnName.Name = "lblCoulmnName";
            this.lblCoulmnName.Size = new System.Drawing.Size(101, 21);
            this.lblCoulmnName.TabIndex = 22;
            this.lblCoulmnName.Text = "Coulmn Name";
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(114, 7);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(59, 22);
            this.textBox1.TabIndex = 21;
            // 
            // lbl_before
            // 
            this.lbl_before.Location = new System.Drawing.Point(7, 7);
            this.lbl_before.Name = "lbl_before";
            this.lbl_before.Size = new System.Drawing.Size(59, 21);
            this.lbl_before.TabIndex = 20;
            this.lbl_before.Text = "Before";
            // 
            // splitContainer2
            // 
            this.splitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer2.Location = new System.Drawing.Point(18, 112);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.chkBoxAddFormulaInCell);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.pnl_XLookup);
            this.splitContainer2.Panel2.Controls.Add(this.cmb_SelectFormula);
            this.splitContainer2.Panel2.Controls.Add(this.lbl_SelectFromula);
            this.splitContainer2.Size = new System.Drawing.Size(599, 173);
            this.splitContainer2.SplitterDistance = 170;
            this.splitContainer2.TabIndex = 26;
            // 
            // chkBoxAddFormulaInCell
            // 
            this.chkBoxAddFormulaInCell.AutoSize = true;
            this.chkBoxAddFormulaInCell.Location = new System.Drawing.Point(12, 19);
            this.chkBoxAddFormulaInCell.Name = "chkBoxAddFormulaInCell";
            this.chkBoxAddFormulaInCell.Size = new System.Drawing.Size(138, 20);
            this.chkBoxAddFormulaInCell.TabIndex = 22;
            this.chkBoxAddFormulaInCell.Text = "Add formula in cell";
            this.chkBoxAddFormulaInCell.UseVisualStyleBackColor = true;
            // 
            // cmb_SelectFormula
            // 
            this.cmb_SelectFormula.FormattingEnabled = true;
            this.cmb_SelectFormula.Items.AddRange(new object[] {
            "XLookup",
            "IF"});
            this.cmb_SelectFormula.Location = new System.Drawing.Point(126, 21);
            this.cmb_SelectFormula.Name = "cmb_SelectFormula";
            this.cmb_SelectFormula.Size = new System.Drawing.Size(238, 24);
            this.cmb_SelectFormula.TabIndex = 24;
            // 
            // lbl_SelectFromula
            // 
            this.lbl_SelectFromula.AutoSize = true;
            this.lbl_SelectFromula.Location = new System.Drawing.Point(11, 23);
            this.lbl_SelectFromula.Name = "lbl_SelectFromula";
            this.lbl_SelectFromula.Size = new System.Drawing.Size(97, 16);
            this.lbl_SelectFromula.TabIndex = 25;
            this.lbl_SelectFromula.Text = "Select Formula";
            // 
            // pnl_XLookup
            // 
            this.pnl_XLookup.Controls.Add(this.pnl_If);
            this.pnl_XLookup.Controls.Add(this.cmb_Search_Mode);
            this.pnl_XLookup.Controls.Add(this.cmb_Match_Mode);
            this.pnl_XLookup.Controls.Add(this.txt_IfNotFound);
            this.pnl_XLookup.Controls.Add(this.txt_ReturnArray);
            this.pnl_XLookup.Controls.Add(this.txt_LookupArray);
            this.pnl_XLookup.Controls.Add(this.txt_LookupValue);
            this.pnl_XLookup.Controls.Add(this.lbl_Search_Mode);
            this.pnl_XLookup.Controls.Add(this.lbl_Match_Mode);
            this.pnl_XLookup.Controls.Add(this.lbl_if_not_found);
            this.pnl_XLookup.Controls.Add(this.lbl_ReturnArray);
            this.pnl_XLookup.Controls.Add(this.lbl_LookupArray);
            this.pnl_XLookup.Controls.Add(this.lbl_LookupValue);
            this.pnl_XLookup.Location = new System.Drawing.Point(14, 60);
            this.pnl_XLookup.Name = "pnl_XLookup";
            this.pnl_XLookup.Size = new System.Drawing.Size(405, 101);
            this.pnl_XLookup.TabIndex = 26;
            this.pnl_XLookup.Visible = false;
            // 
            // cmb_Search_Mode
            // 
            this.cmb_Search_Mode.FormattingEnabled = true;
            this.cmb_Search_Mode.Items.AddRange(new object[] {
            "0 - Search Mode1",
            "1 - Search Mode2",
            "3 - Search Mode3"});
            this.cmb_Search_Mode.Location = new System.Drawing.Point(269, 67);
            this.cmb_Search_Mode.Name = "cmb_Search_Mode";
            this.cmb_Search_Mode.Size = new System.Drawing.Size(105, 24);
            this.cmb_Search_Mode.TabIndex = 11;
            // 
            // cmb_Match_Mode
            // 
            this.cmb_Match_Mode.FormattingEnabled = true;
            this.cmb_Match_Mode.Items.AddRange(new object[] {
            "0 - Test1",
            "1 - Test2",
            "3 - Test3"});
            this.cmb_Match_Mode.Location = new System.Drawing.Point(269, 38);
            this.cmb_Match_Mode.Name = "cmb_Match_Mode";
            this.cmb_Match_Mode.Size = new System.Drawing.Size(105, 24);
            this.cmb_Match_Mode.TabIndex = 10;
            // 
            // txt_IfNotFound
            // 
            this.txt_IfNotFound.Location = new System.Drawing.Point(269, 9);
            this.txt_IfNotFound.Name = "txt_IfNotFound";
            this.txt_IfNotFound.Size = new System.Drawing.Size(105, 22);
            this.txt_IfNotFound.TabIndex = 9;
            // 
            // txt_ReturnArray
            // 
            this.txt_ReturnArray.Location = new System.Drawing.Point(102, 67);
            this.txt_ReturnArray.Name = "txt_ReturnArray";
            this.txt_ReturnArray.Size = new System.Drawing.Size(44, 22);
            this.txt_ReturnArray.TabIndex = 8;
            // 
            // txt_LookupArray
            // 
            this.txt_LookupArray.Location = new System.Drawing.Point(102, 38);
            this.txt_LookupArray.Name = "txt_LookupArray";
            this.txt_LookupArray.Size = new System.Drawing.Size(44, 22);
            this.txt_LookupArray.TabIndex = 7;
            // 
            // txt_LookupValue
            // 
            this.txt_LookupValue.Location = new System.Drawing.Point(102, 9);
            this.txt_LookupValue.Name = "txt_LookupValue";
            this.txt_LookupValue.Size = new System.Drawing.Size(44, 22);
            this.txt_LookupValue.TabIndex = 6;
            // 
            // lbl_Search_Mode
            // 
            this.lbl_Search_Mode.AutoSize = true;
            this.lbl_Search_Mode.Location = new System.Drawing.Point(171, 67);
            this.lbl_Search_Mode.Name = "lbl_Search_Mode";
            this.lbl_Search_Mode.Size = new System.Drawing.Size(91, 16);
            this.lbl_Search_Mode.TabIndex = 22;
            this.lbl_Search_Mode.Text = "Search Mode:";
            // 
            // lbl_Match_Mode
            // 
            this.lbl_Match_Mode.AutoSize = true;
            this.lbl_Match_Mode.Location = new System.Drawing.Point(171, 38);
            this.lbl_Match_Mode.Name = "lbl_Match_Mode";
            this.lbl_Match_Mode.Size = new System.Drawing.Size(84, 16);
            this.lbl_Match_Mode.TabIndex = 21;
            this.lbl_Match_Mode.Text = "Match Mode:";
            // 
            // lbl_if_not_found
            // 
            this.lbl_if_not_found.AutoSize = true;
            this.lbl_if_not_found.Location = new System.Drawing.Point(171, 9);
            this.lbl_if_not_found.Name = "lbl_if_not_found";
            this.lbl_if_not_found.Size = new System.Drawing.Size(81, 16);
            this.lbl_if_not_found.TabIndex = 20;
            this.lbl_if_not_found.Text = "If Not Found:";
            // 
            // lbl_ReturnArray
            // 
            this.lbl_ReturnArray.AutoSize = true;
            this.lbl_ReturnArray.Location = new System.Drawing.Point(3, 67);
            this.lbl_ReturnArray.Name = "lbl_ReturnArray";
            this.lbl_ReturnArray.Size = new System.Drawing.Size(84, 16);
            this.lbl_ReturnArray.TabIndex = 19;
            this.lbl_ReturnArray.Text = "Return Array:";
            // 
            // lbl_LookupArray
            // 
            this.lbl_LookupArray.AutoSize = true;
            this.lbl_LookupArray.Location = new System.Drawing.Point(3, 38);
            this.lbl_LookupArray.Name = "lbl_LookupArray";
            this.lbl_LookupArray.Size = new System.Drawing.Size(90, 16);
            this.lbl_LookupArray.TabIndex = 18;
            this.lbl_LookupArray.Text = "Lookup Array:";
            // 
            // lbl_LookupValue
            // 
            this.lbl_LookupValue.AutoSize = true;
            this.lbl_LookupValue.Location = new System.Drawing.Point(3, 9);
            this.lbl_LookupValue.Name = "lbl_LookupValue";
            this.lbl_LookupValue.Size = new System.Drawing.Size(93, 16);
            this.lbl_LookupValue.TabIndex = 17;
            this.lbl_LookupValue.Text = "Lookup Value:";
            // 
            // pnl_If
            // 
            this.pnl_If.Controls.Add(this.txtFalseValue);
            this.pnl_If.Controls.Add(this.txtTrueValue);
            this.pnl_If.Controls.Add(this.txt_LogicValue);
            this.pnl_If.Controls.Add(this.lbl_FalseValue);
            this.pnl_If.Controls.Add(this.lbl_TrueValue);
            this.pnl_If.Controls.Add(this.lbl_LogicValue);
            this.pnl_If.Location = new System.Drawing.Point(112, 68);
            this.pnl_If.Name = "pnl_If";
            this.pnl_If.Size = new System.Drawing.Size(250, 101);
            this.pnl_If.TabIndex = 27;
            this.pnl_If.Visible = false;
            // 
            // txtFalseValue
            // 
            this.txtFalseValue.Location = new System.Drawing.Point(95, 67);
            this.txtFalseValue.Name = "txtFalseValue";
            this.txtFalseValue.Size = new System.Drawing.Size(105, 22);
            this.txtFalseValue.TabIndex = 14;
            // 
            // txtTrueValue
            // 
            this.txtTrueValue.Location = new System.Drawing.Point(95, 38);
            this.txtTrueValue.Name = "txtTrueValue";
            this.txtTrueValue.Size = new System.Drawing.Size(105, 22);
            this.txtTrueValue.TabIndex = 13;
            // 
            // txt_LogicValue
            // 
            this.txt_LogicValue.Location = new System.Drawing.Point(95, 9);
            this.txt_LogicValue.Name = "txt_LogicValue";
            this.txt_LogicValue.Size = new System.Drawing.Size(57, 22);
            this.txt_LogicValue.TabIndex = 12;
            // 
            // lbl_FalseValue
            // 
            this.lbl_FalseValue.AutoSize = true;
            this.lbl_FalseValue.Location = new System.Drawing.Point(3, 67);
            this.lbl_FalseValue.Name = "lbl_FalseValue";
            this.lbl_FalseValue.Size = new System.Drawing.Size(82, 16);
            this.lbl_FalseValue.TabIndex = 19;
            this.lbl_FalseValue.Text = "False Value:";
            // 
            // lbl_TrueValue
            // 
            this.lbl_TrueValue.AutoSize = true;
            this.lbl_TrueValue.Location = new System.Drawing.Point(3, 38);
            this.lbl_TrueValue.Name = "lbl_TrueValue";
            this.lbl_TrueValue.Size = new System.Drawing.Size(76, 16);
            this.lbl_TrueValue.TabIndex = 18;
            this.lbl_TrueValue.Text = "True Value:";
            // 
            // lbl_LogicValue
            // 
            this.lbl_LogicValue.AutoSize = true;
            this.lbl_LogicValue.Location = new System.Drawing.Point(3, 9);
            this.lbl_LogicValue.Name = "lbl_LogicValue";
            this.lbl_LogicValue.Size = new System.Drawing.Size(75, 16);
            this.lbl_LogicValue.TabIndex = 17;
            this.lbl_LogicValue.Text = "LogicValue";
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(623, 94);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(61, 34);
            this.btnAdd.TabIndex = 27;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            // 
            // listBox
            // 
            this.listBox.FormattingEnabled = true;
            this.listBox.ItemHeight = 16;
            this.listBox.Location = new System.Drawing.Point(690, 31);
            this.listBox.Name = "listBox";
            this.listBox.Size = new System.Drawing.Size(407, 244);
            this.listBox.TabIndex = 28;
            // 
            // Frm_AutomationTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1140, 671);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.statusStrip1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Frm_AutomationTool";
            this.Text = "Excel Automation Tool";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.grpBoxCoulmn.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel1.PerformLayout();
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.pnl_XLookup.ResumeLayout(false);
            this.pnl_XLookup.PerformLayout();
            this.pnl_If.ResumeLayout(false);
            this.pnl_If.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog oFD_SourceFile;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lbl_SelectWorksheet;
        private System.Windows.Forms.ListBox lstBox_Worksheets;
        private System.Windows.Forms.Button btn_Close;
        private System.Windows.Forms.TextBox txt_ReferenceFilePath_URL;
        private System.Windows.Forms.Label lbl_ReferenceFilePath_URL;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.Button btn_BrowseFile;
        private System.Windows.Forms.TextBox txt_SourceFileName;
        private System.Windows.Forms.Label lbl_SourceFileName;
        private System.Windows.Forms.GroupBox grpBoxCoulmn;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.CheckBox chkBoxAddFormulaInCell;
        private System.Windows.Forms.Panel pnl_XLookup;
        private System.Windows.Forms.Panel pnl_If;
        private System.Windows.Forms.TextBox txtFalseValue;
        private System.Windows.Forms.TextBox txtTrueValue;
        private System.Windows.Forms.TextBox txt_LogicValue;
        private System.Windows.Forms.Label lbl_FalseValue;
        private System.Windows.Forms.Label lbl_TrueValue;
        private System.Windows.Forms.Label lbl_LogicValue;
        private System.Windows.Forms.ComboBox cmb_Search_Mode;
        private System.Windows.Forms.ComboBox cmb_Match_Mode;
        private System.Windows.Forms.TextBox txt_IfNotFound;
        private System.Windows.Forms.TextBox txt_ReturnArray;
        private System.Windows.Forms.TextBox txt_LookupArray;
        private System.Windows.Forms.TextBox txt_LookupValue;
        private System.Windows.Forms.Label lbl_Search_Mode;
        private System.Windows.Forms.Label lbl_Match_Mode;
        private System.Windows.Forms.Label lbl_if_not_found;
        private System.Windows.Forms.Label lbl_ReturnArray;
        private System.Windows.Forms.Label lbl_LookupArray;
        private System.Windows.Forms.Label lbl_LookupValue;
        private System.Windows.Forms.ComboBox cmb_SelectFormula;
        private System.Windows.Forms.Label lbl_SelectFromula;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.CheckBox chkBox_BetweenColumns;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label lblCoulmnName;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label lbl_before;
        private System.Windows.Forms.ListBox listBox;
        private System.Windows.Forms.Button btnAdd;
    }
}

