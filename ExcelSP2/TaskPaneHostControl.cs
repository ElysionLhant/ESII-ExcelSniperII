using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ExcelSP2
{
    public partial class TaskPaneHostControl : UserControl
    {
        private ElementHost elementHost;
        private WpfTaskPaneControl wpfControl;

        public TaskPaneHostControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.elementHost = new ElementHost();
            this.wpfControl = new WpfTaskPaneControl();
            this.SuspendLayout();
            // 
            // elementHost
            // 
            this.elementHost.Dock = DockStyle.Fill;
            this.elementHost.Location = new System.Drawing.Point(0, 0);
            this.elementHost.Name = "elementHost";
            this.elementHost.Size = new System.Drawing.Size(300, 600);
            this.elementHost.TabIndex = 0;
            this.elementHost.Text = "elementHost";
            this.elementHost.Child = this.wpfControl;
            // 
            // TaskPaneHostControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.Controls.Add(this.elementHost);
            this.Name = "TaskPaneHostControl";
            this.Size = new System.Drawing.Size(300, 600);
            this.ResumeLayout(false);
        }
    }
}
