namespace ArcVera_Tech_Test
{
    partial class frmMain
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            btnImportEra5 = new Button();
            btnExportCsv = new Button();
            btnExportExcel = new Button();
            dgImportedEra5 = new DataGridView();
            plotView1 = new OxyPlot.WindowsForms.PlotView();
            ((System.ComponentModel.ISupportInitialize)dgImportedEra5).BeginInit();
            SuspendLayout();
            // 
            // btnImportEra5
            // 
            btnImportEra5.Location = new Point(12, 12);
            btnImportEra5.Name = "btnImportEra5";
            btnImportEra5.Size = new Size(125, 23);
            btnImportEra5.TabIndex = 0;
            btnImportEra5.Text = "Import ERA5 Data";
            btnImportEra5.UseVisualStyleBackColor = true;
            btnImportEra5.Click += btnImportEra5_Click;
            // 
            // btnExportCsv
            // 
            btnExportCsv.Location = new Point(12, 549);
            btnExportCsv.Name = "btnExportCsv";
            btnExportCsv.Size = new Size(125, 23);
            btnExportCsv.TabIndex = 1;
            btnExportCsv.Text = "Export CSV Dataset";
            btnExportCsv.UseVisualStyleBackColor = true;
            btnExportCsv.Click += btnExportCsv_Click;
            // 
            // btnExportExcel
            // 
            btnExportExcel.Location = new Point(143, 549);
            btnExportExcel.Name = "btnExportExcel";
            btnExportExcel.Size = new Size(125, 23);
            btnExportExcel.TabIndex = 2;
            btnExportExcel.Text = "Export Excel";
            btnExportExcel.UseVisualStyleBackColor = true;
            btnExportExcel.Click += btnExportExcel_Click;
            // 
            // dgImportedEra5
            // 
            dgImportedEra5.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgImportedEra5.Location = new Point(12, 79);
            dgImportedEra5.Name = "dgImportedEra5";
            dgImportedEra5.Size = new Size(576, 464);
            dgImportedEra5.TabIndex = 3;
            // 
            // plotView1
            // 
            plotView1.Location = new Point(605, 79);
            plotView1.Name = "plotView1";
            plotView1.PanCursor = Cursors.Hand;
            plotView1.Size = new Size(558, 464);
            plotView1.TabIndex = 4;
            plotView1.Text = "pltAvgWd";
            plotView1.ZoomHorizontalCursor = Cursors.SizeWE;
            plotView1.ZoomRectangleCursor = Cursors.SizeNWSE;
            plotView1.ZoomVerticalCursor = Cursors.SizeNS;
            // 
            // frmMain
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1175, 584);
            Controls.Add(plotView1);
            Controls.Add(dgImportedEra5);
            Controls.Add(btnExportExcel);
            Controls.Add(btnExportCsv);
            Controls.Add(btnImportEra5);
            Name = "frmMain";
            Text = "WindSpeed Reader";
            ((System.ComponentModel.ISupportInitialize)dgImportedEra5).EndInit();
            ResumeLayout(false);
        }

        #endregion
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private Button btnImportEra5;
        private Button btnExportCsv;
        private Button btnExportExcel;
        private DataGridView dgImportedEra5;
        private OxyPlot.WindowsForms.PlotView plotView1;
    }
}
