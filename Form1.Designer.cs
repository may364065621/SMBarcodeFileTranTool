using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinFormsApp1
{
    partial class Form1
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
            labelLocate = new Label();
            cbb_Locate = new ComboBox();
            label_Date = new Label();
            btn_tran = new Button();
            label_savepath = new Label();
            tb_BarcodeContrast = new TextBox();
            label1 = new Label();
            label2 = new Label();
            tb_InTable = new TextBox();
            label5 = new Label();
            tb_OutTable = new TextBox();
            btn_BarcodeContrast = new Button();
            SuspendLayout();
            // 
            // labelLocate
            // 
            labelLocate.AutoSize = true;
            labelLocate.Location = new Point(57, 87);
            labelLocate.Name = "labelLocate";
            labelLocate.Size = new Size(68, 17);
            labelLocate.TabIndex = 1;
            labelLocate.Text = "位置选择：";
            // 
            // cbb_Locate
            // 
            cbb_Locate.FormattingEnabled = true;
            cbb_Locate.Items.AddRange(new object[] { "FI" });
            cbb_Locate.Location = new Point(127, 84);
            cbb_Locate.Name = "cbb_Locate";
            cbb_Locate.Size = new Size(100, 25);
            cbb_Locate.TabIndex = 2;
            // 
            // label_Date
            // 
            label_Date.AutoSize = true;
            label_Date.Location = new Point(85, 56);
            label_Date.Name = "label_Date";
            label_Date.Size = new Size(0, 17);
            label_Date.TabIndex = 1;
            // 
            // btn_tran
            // 
            btn_tran.Location = new Point(380, 185);
            btn_tran.Name = "btn_tran";
            btn_tran.Size = new Size(118, 50);
            btn_tran.TabIndex = 0;
            btn_tran.Text = "表格转换";
            btn_tran.UseVisualStyleBackColor = true;
            btn_tran.Click += btn_tran_Click;
            // 
            // label_savepath
            // 
            label_savepath.AutoSize = true;
            label_savepath.Location = new Point(122, 202);
            label_savepath.Name = "label_savepath";
            label_savepath.Size = new Size(0, 17);
            label_savepath.TabIndex = 1;
            // 
            // tb_BarcodeContrast
            // 
            tb_BarcodeContrast.Location = new Point(127, 134);
            tb_BarcodeContrast.Name = "tb_BarcodeContrast";
            tb_BarcodeContrast.Size = new Size(247, 23);
            tb_BarcodeContrast.TabIndex = 4;
            tb_BarcodeContrast.DoubleClick += tb_BarcodeContrast_DoubleClick;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(45, 137);
            label1.Name = "label1";
            label1.Size = new Size(80, 17);
            label1.TabIndex = 1;
            label1.Text = "条码对照表：";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(44, 40);
            label2.Name = "label2";
            label2.Size = new Size(81, 17);
            label2.TabIndex = 5;
            label2.Text = "Barcode表：";
            // 
            // tb_InTable
            // 
            tb_InTable.Location = new Point(127, 37);
            tb_InTable.Name = "tb_InTable";
            tb_InTable.Size = new Size(247, 23);
            tb_InTable.TabIndex = 6;
            tb_InTable.DoubleClick += tb_InTable_DoubleClick;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(66, 202);
            label5.Name = "label5";
            label5.Size = new Size(56, 17);
            label5.TabIndex = 5;
            label5.Text = "汇总表：";
            // 
            // tb_OutTable
            // 
            tb_OutTable.Location = new Point(122, 199);
            tb_OutTable.Name = "tb_OutTable";
            tb_OutTable.Size = new Size(252, 23);
            tb_OutTable.TabIndex = 6;
            tb_OutTable.DoubleClick += tb_OutTable_DoubleClick;
            // 
            // btn_BarcodeContrast
            // 
            btn_BarcodeContrast.Location = new Point(380, 120);
            btn_BarcodeContrast.Name = "btn_BarcodeContrast";
            btn_BarcodeContrast.Size = new Size(118, 50);
            btn_BarcodeContrast.TabIndex = 0;
            btn_BarcodeContrast.Text = "读取对照表";
            btn_BarcodeContrast.UseVisualStyleBackColor = true;
            btn_BarcodeContrast.Click += btn_BarcodeContrast_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(510, 237);
            Controls.Add(tb_OutTable);
            Controls.Add(tb_InTable);
            Controls.Add(label5);
            Controls.Add(label2);
            Controls.Add(tb_BarcodeContrast);
            Controls.Add(cbb_Locate);
            Controls.Add(label_savepath);
            Controls.Add(label_Date);
            Controls.Add(label1);
            Controls.Add(labelLocate);
            Controls.Add(btn_BarcodeContrast);
            Controls.Add(btn_tran);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Label labelLocate;
        private ComboBox cbb_Locate;
        private Label label_Date;
        private Button btn_tran;
        private Label label_savepath;
        private TextBox tb_BarcodeContrast;
        private Label label1;
        private Label label2;
        private TextBox tb_InTable;
        private Label label5;
        private TextBox tb_OutTable;
        private Button btn_BarcodeContrast;
    }
}