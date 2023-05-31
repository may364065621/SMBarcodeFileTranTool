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
            btn_readcsv = new Button();
            labelLocate = new Label();
            cbb_Locate = new ComboBox();
            label_Date = new Label();
            btn_tran = new Button();
            label_savepath = new Label();
            SuspendLayout();
            // 
            // btn_readcsv
            // 
            btn_readcsv.Location = new Point(53, 38);
            btn_readcsv.Name = "btn_readcsv";
            btn_readcsv.Size = new Size(118, 50);
            btn_readcsv.TabIndex = 0;
            btn_readcsv.Text = "选择表格";
            btn_readcsv.UseVisualStyleBackColor = true;
            btn_readcsv.Click += btn_readcsv_Click;
            // 
            // labelLocate
            // 
            labelLocate.AutoSize = true;
            labelLocate.Location = new Point(53, 110);
            labelLocate.Name = "labelLocate";
            labelLocate.Size = new Size(68, 17);
            labelLocate.TabIndex = 1;
            labelLocate.Text = "位置选择：";
            // 
            // cbb_Locate
            // 
            cbb_Locate.FormattingEnabled = true;
            cbb_Locate.Items.AddRange(new object[] { "FI" });
            cbb_Locate.Location = new Point(127, 107);
            cbb_Locate.Name = "cbb_Locate";
            cbb_Locate.Size = new Size(121, 25);
            cbb_Locate.TabIndex = 2;
            // 
            // label_Date
            // 
            label_Date.AutoSize = true;
            label_Date.Location = new Point(196, 55);
            label_Date.Name = "label_Date";
            label_Date.Size = new Size(0, 17);
            label_Date.TabIndex = 1;
            // 
            // btn_tran
            // 
            btn_tran.Location = new Point(53, 157);
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
            label_savepath.Location = new Point(196, 174);
            label_savepath.Name = "label_savepath";
            label_savepath.Size = new Size(0, 17);
            label_savepath.TabIndex = 1;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(395, 237);
            Controls.Add(cbb_Locate);
            Controls.Add(label_savepath);
            Controls.Add(label_Date);
            Controls.Add(labelLocate);
            Controls.Add(btn_tran);
            Controls.Add(btn_readcsv);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btn_readcsv;
        private Label labelLocate;
        private ComboBox cbb_Locate;
        private Label label_Date;
        private Button btn_tran;
        private Label label_savepath;
    }
}