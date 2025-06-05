namespace SaovietTax
{
    partial class frmDinhdanh
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDinhdanh));
            this.gcDinhdanh = new DevExpress.XtraGrid.GridControl();
            this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colKeyValue = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colTKNo = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colTKCo = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colTKThue = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colType = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDelete = new DevExpress.XtraGrid.Columns.GridColumn();
            this.btnLuudinhdanh = new DevExpress.XtraEditors.SimpleButton();
            this.txtTKThue = new DevExpress.XtraEditors.TextEdit();
            this.labelControl11 = new DevExpress.XtraEditors.LabelControl();
            this.txtTKCo = new DevExpress.XtraEditors.TextEdit();
            this.labelControl10 = new DevExpress.XtraEditors.LabelControl();
            this.txtTKNo = new DevExpress.XtraEditors.TextEdit();
            this.labelControl9 = new DevExpress.XtraEditors.LabelControl();
            this.txtDiengiai = new DevExpress.XtraEditors.TextEdit();
            this.txtTukhoa = new DevExpress.XtraEditors.TextEdit();
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            this.checkEdit2 = new DevExpress.XtraEditors.CheckEdit();
            this.checkEdit1 = new DevExpress.XtraEditors.CheckEdit();
            this.labelControl8 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl7 = new DevExpress.XtraEditors.LabelControl();
            this.behaviorManager1 = new DevExpress.Utils.Behaviors.BehaviorManager(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.gcDinhdanh)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTKThue.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTKCo.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTKNo.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtDiengiai.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTukhoa.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.checkEdit2.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.behaviorManager1)).BeginInit();
            this.SuspendLayout();
            // 
            // gcDinhdanh
            // 
            this.gcDinhdanh.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gcDinhdanh.Location = new System.Drawing.Point(3, 178);
            this.gcDinhdanh.MainView = this.gridView2;
            this.gcDinhdanh.Name = "gcDinhdanh";
            this.gcDinhdanh.Size = new System.Drawing.Size(1034, 387);
            this.gcDinhdanh.TabIndex = 11;
            this.gcDinhdanh.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView2});
            // 
            // gridView2
            // 
            this.gridView2.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colKeyValue,
            this.colTKNo,
            this.colTKCo,
            this.colTKThue,
            this.colType,
            this.colDelete});
            this.gridView2.GridControl = this.gcDinhdanh;
            this.gridView2.Name = "gridView2";
            // 
            // colID
            // 
            this.colID.Caption = "ID";
            this.colID.FieldName = "ID";
            this.colID.MinWidth = 30;
            this.colID.Name = "colID";
            this.colID.Visible = true;
            this.colID.VisibleIndex = 0;
            this.colID.Width = 112;
            // 
            // colKeyValue
            // 
            this.colKeyValue.Caption = "Từ khóa";
            this.colKeyValue.FieldName = "KeyValue";
            this.colKeyValue.MinWidth = 30;
            this.colKeyValue.Name = "colKeyValue";
            this.colKeyValue.Visible = true;
            this.colKeyValue.VisibleIndex = 1;
            this.colKeyValue.Width = 112;
            // 
            // colTKNo
            // 
            this.colTKNo.Caption = "TKNo";
            this.colTKNo.FieldName = "TKNo";
            this.colTKNo.MinWidth = 30;
            this.colTKNo.Name = "colTKNo";
            this.colTKNo.Visible = true;
            this.colTKNo.VisibleIndex = 2;
            this.colTKNo.Width = 112;
            // 
            // colTKCo
            // 
            this.colTKCo.Caption = "TKCo";
            this.colTKCo.FieldName = "TKCo";
            this.colTKCo.MinWidth = 30;
            this.colTKCo.Name = "colTKCo";
            this.colTKCo.Visible = true;
            this.colTKCo.VisibleIndex = 3;
            this.colTKCo.Width = 112;
            // 
            // colTKThue
            // 
            this.colTKThue.Caption = "TKThue";
            this.colTKThue.FieldName = "TKThue";
            this.colTKThue.MinWidth = 30;
            this.colTKThue.Name = "colTKThue";
            this.colTKThue.Visible = true;
            this.colTKThue.VisibleIndex = 4;
            this.colTKThue.Width = 112;
            // 
            // colType
            // 
            this.colType.Caption = "Diễn giải";
            this.colType.FieldName = "Type";
            this.colType.MinWidth = 30;
            this.colType.Name = "colType";
            this.colType.Visible = true;
            this.colType.VisibleIndex = 5;
            this.colType.Width = 112;
            // 
            // colDelete
            // 
            this.colDelete.Caption = "Xóa";
            this.colDelete.FieldName = "colDelete";
            this.colDelete.MinWidth = 30;
            this.colDelete.Name = "colDelete";
            this.colDelete.OptionsColumn.AllowEdit = false;
            this.colDelete.Visible = true;
            this.colDelete.VisibleIndex = 6;
            this.colDelete.Width = 112;
            // 
            // btnLuudinhdanh
            // 
            this.btnLuudinhdanh.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnLuudinhdanh.ImageOptions.Image")));
            this.btnLuudinhdanh.Location = new System.Drawing.Point(899, 110);
            this.btnLuudinhdanh.Name = "btnLuudinhdanh";
            this.btnLuudinhdanh.Size = new System.Drawing.Size(112, 34);
            this.btnLuudinhdanh.TabIndex = 20;
            this.btnLuudinhdanh.Text = "Lưu";
            this.btnLuudinhdanh.Click += new System.EventHandler(this.btnLuudinhdanh_Click);
            // 
            // txtTKThue
            // 
            this.txtTKThue.EditValue = "1331";
            this.txtTKThue.Location = new System.Drawing.Point(103, 117);
            this.txtTKThue.Name = "txtTKThue";
            this.txtTKThue.Size = new System.Drawing.Size(231, 27);
            this.txtTKThue.TabIndex = 19;
            // 
            // labelControl11
            // 
            this.labelControl11.Location = new System.Drawing.Point(22, 119);
            this.labelControl11.Name = "labelControl11";
            this.labelControl11.Size = new System.Drawing.Size(60, 19);
            this.labelControl11.TabIndex = 18;
            this.labelControl11.Text = "TK Thuế";
            // 
            // txtTKCo
            // 
            this.txtTKCo.EditValue = "1111";
            this.txtTKCo.Location = new System.Drawing.Point(103, 83);
            this.txtTKCo.Name = "txtTKCo";
            this.txtTKCo.Size = new System.Drawing.Size(231, 27);
            this.txtTKCo.TabIndex = 17;
            // 
            // labelControl10
            // 
            this.labelControl10.Location = new System.Drawing.Point(22, 85);
            this.labelControl10.Name = "labelControl10";
            this.labelControl10.Size = new System.Drawing.Size(43, 19);
            this.labelControl10.TabIndex = 16;
            this.labelControl10.Text = "TK Có";
            // 
            // txtTKNo
            // 
            this.txtTKNo.EditValue = "152";
            this.txtTKNo.Location = new System.Drawing.Point(103, 49);
            this.txtTKNo.Name = "txtTKNo";
            this.txtTKNo.Size = new System.Drawing.Size(231, 27);
            this.txtTKNo.TabIndex = 15;
            this.txtTKNo.EditValueChanging += new DevExpress.XtraEditors.Controls.ChangingEventHandler(this.txtTKNo_EditValueChanging);
            // 
            // labelControl9
            // 
            this.labelControl9.Location = new System.Drawing.Point(21, 52);
            this.labelControl9.Name = "labelControl9";
            this.labelControl9.Size = new System.Drawing.Size(44, 19);
            this.labelControl9.TabIndex = 14;
            this.labelControl9.Text = "TK Nợ";
            // 
            // txtDiengiai
            // 
            this.txtDiengiai.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDiengiai.Location = new System.Drawing.Point(433, 14);
            this.txtDiengiai.Name = "txtDiengiai";
            this.txtDiengiai.Size = new System.Drawing.Size(576, 27);
            this.txtDiengiai.TabIndex = 13;
            // 
            // txtTukhoa
            // 
            this.txtTukhoa.EditValue = "Ten=\"123\"";
            this.txtTukhoa.Location = new System.Drawing.Point(103, 14);
            this.txtTukhoa.Name = "txtTukhoa";
            this.txtTukhoa.Size = new System.Drawing.Size(231, 27);
            this.txtTukhoa.TabIndex = 12;
            // 
            // panelControl1
            // 
            this.panelControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelControl1.Controls.Add(this.checkEdit2);
            this.panelControl1.Controls.Add(this.checkEdit1);
            this.panelControl1.Controls.Add(this.labelControl8);
            this.panelControl1.Controls.Add(this.labelControl7);
            this.panelControl1.Controls.Add(this.txtDiengiai);
            this.panelControl1.Controls.Add(this.txtTukhoa);
            this.panelControl1.Controls.Add(this.btnLuudinhdanh);
            this.panelControl1.Controls.Add(this.labelControl9);
            this.panelControl1.Controls.Add(this.txtTKThue);
            this.panelControl1.Controls.Add(this.txtTKNo);
            this.panelControl1.Controls.Add(this.labelControl11);
            this.panelControl1.Controls.Add(this.labelControl10);
            this.panelControl1.Controls.Add(this.txtTKCo);
            this.panelControl1.Location = new System.Drawing.Point(3, 12);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(1034, 160);
            this.panelControl1.TabIndex = 21;
            // 
            // checkEdit2
            // 
            this.checkEdit2.Location = new System.Drawing.Point(363, 113);
            this.checkEdit2.Name = "checkEdit2";
            this.checkEdit2.Properties.Caption = "Không phân biệt đơn vị tính";
            this.checkEdit2.Properties.CheckBoxOptions.Style = DevExpress.XtraEditors.Controls.CheckBoxStyle.SvgCheckBox1;
            this.checkEdit2.Size = new System.Drawing.Size(246, 31);
            this.checkEdit2.TabIndex = 1;
            this.checkEdit2.CheckedChanged += new System.EventHandler(this.checkEdit2_CheckedChanged);
            // 
            // checkEdit1
            // 
            this.checkEdit1.Location = new System.Drawing.Point(363, 63);
            this.checkEdit1.Name = "checkEdit1";
            this.checkEdit1.Properties.Caption = "Chọn tên hàng hóa gần giống";
            this.checkEdit1.Properties.CheckBoxOptions.Style = DevExpress.XtraEditors.Controls.CheckBoxStyle.SvgCheckBox1;
            this.checkEdit1.Size = new System.Drawing.Size(264, 31);
            this.checkEdit1.TabIndex = 0;
            this.checkEdit1.CheckedChanged += new System.EventHandler(this.checkEdit1_CheckedChanged);
            // 
            // labelControl8
            // 
            this.labelControl8.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelControl8.Location = new System.Drawing.Point(363, 17);
            this.labelControl8.Name = "labelControl8";
            this.labelControl8.Size = new System.Drawing.Size(62, 19);
            this.labelControl8.TabIndex = 22;
            this.labelControl8.Text = "Diễn giải";
            // 
            // labelControl7
            // 
            this.labelControl7.Location = new System.Drawing.Point(23, 17);
            this.labelControl7.Name = "labelControl7";
            this.labelControl7.Size = new System.Drawing.Size(59, 19);
            this.labelControl7.TabIndex = 21;
            this.labelControl7.Text = "Từ khóa";
            // 
            // frmDinhdanh
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1049, 630);
            this.Controls.Add(this.panelControl1);
            this.Controls.Add(this.gcDinhdanh);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "frmDinhdanh";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Mật định tài khoản";
            this.Load += new System.EventHandler(this.frmDinhdanh_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gcDinhdanh)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTKThue.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTKCo.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTKNo.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtDiengiai.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTukhoa.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            this.panelControl1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.checkEdit2.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.behaviorManager1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.GridControl gcDinhdanh;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
        private DevExpress.XtraEditors.SimpleButton btnLuudinhdanh;
        private DevExpress.XtraEditors.TextEdit txtTKThue;
        private DevExpress.XtraEditors.LabelControl labelControl11;
        private DevExpress.XtraEditors.TextEdit txtTKCo;
        private DevExpress.XtraEditors.LabelControl labelControl10;
        private DevExpress.XtraEditors.TextEdit txtTKNo;
        private DevExpress.XtraEditors.LabelControl labelControl9;
        private DevExpress.XtraEditors.TextEdit txtDiengiai;
        private DevExpress.XtraEditors.TextEdit txtTukhoa;
        private DevExpress.XtraEditors.PanelControl panelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl8;
        private DevExpress.XtraEditors.LabelControl labelControl7;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colTKNo;
        private DevExpress.XtraGrid.Columns.GridColumn colKeyValue;
        private DevExpress.XtraGrid.Columns.GridColumn colTKCo;
        private DevExpress.XtraGrid.Columns.GridColumn colTKThue;
        private DevExpress.XtraGrid.Columns.GridColumn colType;
        private DevExpress.Utils.Behaviors.BehaviorManager behaviorManager1;
        private DevExpress.XtraGrid.Columns.GridColumn colDelete;
        private DevExpress.XtraEditors.CheckEdit checkEdit2;
        private DevExpress.XtraEditors.CheckEdit checkEdit1;
    }
}