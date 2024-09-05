using System;
using System.Drawing;
using System.Timers;
using System.Windows.Forms;
using Microsoft.Toolkit.Uwp.Notifications;
using Windows.Foundation.Collections;
using Windows.UI.Notifications;

namespace MyWorkDiary
{
    public partial class Form1 : Form
    {
        private NotifyIcon _notifyIcon;
        private ContextMenuStrip _contextMenu;
        private System.Timers.Timer _dailyNotificationTimer;
        private System.Timers.Timer _reminderTimer;
        private bool _notificationSent = false;

        public Form1()
        {
            InitializeComponent();
            InitializeTrayIcon();
            InitializeDailyNotificationTimer();
            InitializeReminderTimer();
            // 监听通知激活(点击)
            ToastNotificationManagerCompat.OnActivated += toastArgs =>
            {
                // 通知参数
                ToastArguments args = ToastArguments.Parse(toastArgs.Argument);
                // 获取任何用户输入
                ValueSet userInput = toastArgs.UserInput;

                BeginInvoke(new Action(delegate
                {
                    // 处理通知点击事件
                    Show();
                    WindowState = FormWindowState.Normal;
                    Activate();
                    // TODO: UI线程的操作
                    //MessageBox.Show("Toast被激活（点击），参数是: " + toastArgs.Argument);
                    _notificationSent = true;
                }));
            };
        }

        private void InitializeTrayIcon()
        {
            // 创建托盘图标
            _notifyIcon = new NotifyIcon
            {
                Icon = new Icon("./logo.ico"), // 替换为你的图标文件路径
                Text = "个人日志",
                Visible = true
            };

            // 创建托盘图标的上下文菜单
            _contextMenu = new ContextMenuStrip();
            _notifyIcon.ContextMenuStrip = _contextMenu;

            // 添加“显示”菜单项
            ToolStripMenuItem showMenuItem = new ToolStripMenuItem("显示");
            showMenuItem.Click += ShowMenuItem_Click;
            _contextMenu.Items.Add(showMenuItem);

            // 添加“退出”菜单项
            ToolStripMenuItem exitMenuItem = new ToolStripMenuItem("退出");
            exitMenuItem.Click += ExitMenuItem_Click;
            _contextMenu.Items.Add(exitMenuItem);
        }

        private void ShowMenuItem_Click(object sender, EventArgs e)
        {
            // 将窗体从托盘恢复显示
            Show();
            WindowState = FormWindowState.Normal;
            Activate();
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            // 退出应用程序
            Application.Exit();
        }

        private void InitializeDailyNotificationTimer()
        {
            // 初始化定时器
            _dailyNotificationTimer = new System.Timers.Timer();
            _dailyNotificationTimer.Interval = 1000; // 1秒
            _dailyNotificationTimer.Elapsed += DailyNotificationTimer_Elapsed;
            _dailyNotificationTimer.Start();
        }

        private void InitializeReminderTimer()
        {
            // 初始化提醒定时器
            _reminderTimer = new System.Timers.Timer();
            _reminderTimer.Interval = 5 * 60 * 1000; // 5分钟
            _reminderTimer.Elapsed += ReminderTimer_Elapsed;
        }

        private void DailyNotificationTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            // 检查当前时间是否为每天下午17点15分
            DateTime now = DateTime.Now;
            if (now.Hour == 17 && now.Minute == 15 && now.DayOfWeek != DayOfWeek.Saturday && now.DayOfWeek != DayOfWeek.Sunday)
            {
                // 发送通知
                SendNotification();
                _notificationSent = false;
                _reminderTimer.Start();
            }
        }

        private void ReminderTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            // 检查当前时间是否超过18点
            DateTime now = DateTime.Now;
            if (now.Hour < 18 && !_notificationSent)
            {
                // 发送通知
                SendNotification();
                _notificationSent = false;
            }
            else
            {
                // 停止定时器
                _reminderTimer.Stop();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 当窗体关闭时最小化到托盘
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }

            base.OnFormClosing(e);
        }

        private void SendNotification()
        {
            // 创建通知
            new ToastContentBuilder()
            .AddText("CodeMissing发来一条消息") // 标题文本
            .AddText("请检查消息内容，并及时处理")
            .Show(); // 7.0以上才提供Show方法
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = "./日报.xlsx";
            var title = textBox1.Text;
            var body = textBox2.Text;
            var time = dateTimePicker1.Text;

            string[] rowData = { title, body, time };
            try
            {
                ExcelHelper helper = new ExcelHelper();

                helper.AppendRowToMonthlySheet(filePath, rowData);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            Console.WriteLine("数据已成功追加到Excel文件中。");
            MessageBox.Show("新的一天辛苦了，收拾好心情准备下班了！Goodbye！");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SendNotification();
        }
    }
}
