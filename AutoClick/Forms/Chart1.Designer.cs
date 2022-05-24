
namespace AutoClick
{
    partial class Chart1
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Title title2 = new System.Windows.Forms.DataVisualization.Charting.Title();
            this.chart2 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.cMS_VINADataSet = new AutoClick.CMS_VINADataSet();
            this.zTBPOTableBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.zTBPOTableTableAdapter = new AutoClick.CMS_VINADataSetTableAdapters.ZTBPOTableTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cMS_VINADataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.zTBPOTableBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // chart2
            // 
            this.chart2.BorderlineColor = System.Drawing.Color.DarkRed;
            chartArea2.Name = "ChartArea1";
            this.chart2.ChartAreas.Add(chartArea2);
            this.chart2.DataSource = this.zTBPOTableBindingSource;
            legend2.Name = "Legend1";
            this.chart2.Legends.Add(legend2);
            this.chart2.Location = new System.Drawing.Point(23, 12);
            this.chart2.Name = "chart2";
            series2.ChartArea = "ChartArea1";
            series2.Legend = "Legend1";
            series2.Name = "Series1";
            this.chart2.Series.Add(series2);
            this.chart2.Size = new System.Drawing.Size(649, 409);
            this.chart2.TabIndex = 0;
            this.chart2.Text = "Weekly PO";
            title2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            title2.Name = "Weekly PO";
            title2.Text = "Weekly PO";
            this.chart2.Titles.Add(title2);
            // 
            // cMS_VINADataSet
            // 
            this.cMS_VINADataSet.DataSetName = "CMS_VINADataSet";
            this.cMS_VINADataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // zTBPOTableBindingSource
            // 
            this.zTBPOTableBindingSource.DataMember = "ZTBPOTable";
            this.zTBPOTableBindingSource.DataSource = this.cMS_VINADataSet;
            // 
            // zTBPOTableTableAdapter
            // 
            this.zTBPOTableTableAdapter.ClearBeforeFill = true;
            // 
            // Chart1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(698, 445);
            this.Controls.Add(this.chart2);
            this.Name = "Chart1";
            this.Text = "Chart1";
            this.Load += new System.EventHandler(this.Chart1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cMS_VINADataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.zTBPOTableBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chart2;
        private CMS_VINADataSet cMS_VINADataSet;
        private System.Windows.Forms.BindingSource zTBPOTableBindingSource;
        private CMS_VINADataSetTableAdapters.ZTBPOTableTableAdapter zTBPOTableTableAdapter;
    }
}