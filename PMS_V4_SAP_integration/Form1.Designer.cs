namespace PMS_V4_SAP_integration
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
            textBox1 = new System.Windows.Forms.TextBox();
            listView = new System.Windows.Forms.ListView();
            Time = new System.Windows.Forms.ColumnHeader();
            Message = new System.Windows.Forms.ColumnHeader();
            SuspendLayout();
            // 
            // textBox1
            // 
            textBox1.Location = new System.Drawing.Point(12, 12);
            textBox1.Name = "textBox1";
            textBox1.ReadOnly = true;
            textBox1.Size = new System.Drawing.Size(247, 27);
            textBox1.TabIndex = 0;
            // 
            // listView
            // 
            listView.Activation = System.Windows.Forms.ItemActivation.OneClick;
            listView.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            listView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] { Time, Message });
            listView.HideSelection = false;
            listView.HoverSelection = true;
            listView.Location = new System.Drawing.Point(12, 62);
            listView.MinimumSize = new System.Drawing.Size(776, 376);
            listView.Name = "listView";
            listView.Size = new System.Drawing.Size(776, 376);
            listView.TabIndex = 1;
            listView.UseCompatibleStateImageBehavior = false;
            // 
            // Time
            // 
            Time.Text = "Time";
            Time.Width = 200;
            // 
            // Message
            // 
            Message.Text = "Message";
            Message.Width = 600;
            // 
            // Form1
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(800, 450);
            Controls.Add(listView);
            Controls.Add(textBox1);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ListView listView;
        private System.Windows.Forms.ColumnHeader Time;
        private System.Windows.Forms.ColumnHeader Message;
    }
}
