namespace SaovietTax
{
    partial class frmMatdinhNganHang
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMatdinhNganHang));
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            this.labelControl8 = new DevExpress.XtraEditors.LabelControl();
            this.txtDiengiai = new DevExpress.XtraEditors.TextEdit();
            this.btnLuudinhdanh = new DevExpress.XtraEditors.SimpleButton();
            this.gcDinhdanh = new DevExpress.XtraGrid.GridControl();
            this.matdinhnganhangBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colNoidung = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.textEdit1 = new DevExpress.XtraEditors.TextEdit();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtDiengiai.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcDinhdanh)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.matdinhnganhangBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // panelControl1
            // 
            this.panelControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelControl1.Controls.Add(this.labelControl1);
            this.panelControl1.Controls.Add(this.textEdit1);
            this.panelControl1.Controls.Add(this.labelControl8);
            this.panelControl1.Controls.Add(this.txtDiengiai);
            this.panelControl1.Controls.Add(this.btnLuudinhdanh);
            this.panelControl1.Location = new System.Drawing.Point(12, 12);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(932, 72);
            this.panelControl1.TabIndex = 22;
            // 
            // labelControl8
            // 
            this.labelControl8.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelControl8.Location = new System.Drawing.Point(18, 15);
            this.labelControl8.Name = "labelControl8";
            this.labelControl8.Size = new System.Drawing.Size(65, 19);
            this.labelControl8.TabIndex = 22;
            this.labelControl8.Text = "Nội dung";
            // 
            // txtDiengiai
            // 
            this.txtDiengiai.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDiengiai.Location = new System.Drawing.Point(102, 12);
            this.txtDiengiai.Name = "txtDiengiai";
            this.txtDiengiai.Size = new System.Drawing.Size(459, 27);
            this.txtDiengiai.TabIndex = 13;
            // 
            // btnLuudinhdanh
            // 
            this.btnLuudinhdanh.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnLuudinhdanh.ImageOptions.Image")));
            this.btnLuudinhdanh.Location = new System.Drawing.Point(798, 7);
            this.btnLuudinhdanh.Name = "btnLuudinhdanh";
            this.btnLuudinhdanh.Size = new System.Drawing.Size(119, 34);
            this.btnLuudinhdanh.TabIndex = 20;
            this.btnLuudinhdanh.Text = "Lưu";
            this.btnLuudinhdanh.Click += new System.EventHandler(this.btnLuudinhdanh_Click);
            // 
            // gcDinhdanh
            // 
            this.gcDinhdanh.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gcDinhdanh.DataSource = this.matdinhnganhangBindingSource;
            this.gcDinhdanh.Location = new System.Drawing.Point(12, 90);
            this.gcDinhdanh.MainView = this.gridView2;
            this.gcDinhdanh.Name = "gcDinhdanh";
            this.gcDinhdanh.Size = new System.Drawing.Size(932, 478);
            this.gcDinhdanh.TabIndex = 23;
            this.gcDinhdanh.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView2});
            // 
            // matdinhnganhangBindingSource
            // 
            this.matdinhnganhangBindingSource.DataSource = typeof(SaovietTax.DTO.Matdinhnganhang);
            // 
            // gridView2
            // 
            this.gridView2.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colNoidung,
            this.gridColumn2,
            this.gridColumn1});
            this.gridView2.GridControl = this.gcDinhdanh;
            this.gridView2.Name = "gridView2";
            this.gridView2.OptionsView.ShowGroupPanel = false;
            this.gridView2.RowCellClick += new DevExpress.XtraGrid.Views.Grid.RowCellClickEventHandler(this.gridView2_RowCellClick);
            this.gridView2.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.gridView2_CellValueChanged);
            // 
            // colID
            // 
            this.colID.FieldName = "ID";
            this.colID.MinWidth = 30;
            this.colID.Name = "colID";
            this.colID.Width = 112;
            // 
            // colNoidung
            // 
            this.colNoidung.Caption = "Nội dung";
            this.colNoidung.FieldName = "Noidung";
            this.colNoidung.MinWidth = 30;
            this.colNoidung.Name = "colNoidung";
            this.colNoidung.Visible = true;
            this.colNoidung.VisibleIndex = 0;
            this.colNoidung.Width = 112;
            // 
            // gridColumn1
            // 
            this.gridColumn1.Caption = "Xóa";
            this.gridColumn1.FieldName = "gridColumn1";
            this.gridColumn1.MinWidth = 30;
            this.gridColumn1.Name = "gridColumn1";
            this.gridColumn1.OptionsColumn.AllowEdit = false;
            this.gridColumn1.Visible = true;
            this.gridColumn1.VisibleIndex = 2;
            this.gridColumn1.Width = 112;
            // 
            // textEdit1
            // 
            this.textEdit1.Location = new System.Drawing.Point(664, 14);
            this.textEdit1.Name = "textEdit1";
            this.textEdit1.Size = new System.Drawing.Size(119, 27);
            this.textEdit1.TabIndex = 23;
            // 
            // labelControl1
            // 
            this.labelControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelControl1.Location = new System.Drawing.Point(579, 17);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(70, 19);
            this.labelControl1.TabIndex = 24;
            this.labelControl1.Text = "Tài khoản";
            // 
            // gridColumn2
            // 
            this.gridColumn2.Caption = "Tài khoản";
            this.gridColumn2.FieldName = "TK";
            this.gridColumn2.MinWidth = 30;
            this.gridColumn2.Name = "gridColumn2";
            this.gridColumn2.Visible = true;
            this.gridColumn2.VisibleIndex = 1;
            this.gridColumn2.Width = 112;
            // 
            // frmMatdinhNganHang
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(956, 639);
            this.Controls.Add(this.gcDinhdanh);
            this.Controls.Add(this.panelControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.IconOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("frmMatdinhNganHang.IconOptions.LargeImage")));
            this.Name = "frmMatdinhNganHang";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Mật định ngân hàng";
            this.Load += new System.EventHandler(this.frmMatdinhNganHang_Load);
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            this.panelControl1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtDiengiai.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcDinhdanh)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.matdinhnganhangBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.PanelControl panelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl8;
        private DevExpress.XtraEditors.TextEdit txtDiengiai;
        private DevExpress.XtraEditors.SimpleButton btnLuudinhdanh;
        private DevExpress.XtraGrid.GridControl gcDinhdanh;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
        private System.Windows.Forms.BindingSource matdinhnganhangBindingSource;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colNoidung;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
        private DevExpress.XtraEditors.TextEdit textEdit1;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
    }
}