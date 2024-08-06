using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CEC_NumRule
{
    /// <summary>
    /// ProgressUI.xaml 的互動邏輯
    /// </summary>
    public partial class ProgressUI : Window,IDisposable
    {
        public bool IsClosed { get; private set; }
        private Task taskDoEvent { get; set; }
        public ProgressUI(double maximum)
        {
            InitializeComponent();
            InitializeSize();
            Thread.Sleep(100);
            this.ProgressBarStatus.Text = "套管資訊更新中....";
            this.UpdateProgressBar.Maximum = maximum;
            //簡單的事件訂閱
            this.Closed += (s, e) =>
            {
                IsClosed = true;
            };
        }
        public void Dispose()
        {
            if (!IsClosed) Close();
        }
        //Function to Update the progressBar
        public bool Update(double value = 1.0)
        {
            UpdateTaskDoEvent();
            if (this.UpdateProgressBar.Value + value >= UpdateProgressBar.Maximum)
            {
                //算到最後一個迴圈時，需要單獨做這樣的處理
                UpdateProgressBar.Maximum += value;
            }
            UpdateProgressBar.Value += value;
            return IsClosed;

        }
        private void UpdateTaskDoEvent()
        {
            if (taskDoEvent == null) taskDoEvent = GetTaskUpdateEvent();
            if (taskDoEvent.IsCompleted)
            {
                Show();
                DoEvents();
                taskDoEvent = null;
            }
        }

        private Task GetTaskUpdateEvent()
        {
            return Task.Run(async () => { await Task.Delay(100); });
        }
        private void DoEvents()
        {
            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
        }
            private void InitializeSize()
        {
            //this.SizeToContent = SizeToContent.WidthAndHeight;
            this.Topmost = true;
            this.ShowInTaskbar = false;
            this.ResizeMode = ResizeMode.NoResize;
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }
    }
}
