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
            // ����֪ͨ����(���)
            ToastNotificationManagerCompat.OnActivated += toastArgs =>
            {
                // ֪ͨ����
                ToastArguments args = ToastArguments.Parse(toastArgs.Argument);
                // ��ȡ�κ��û�����
                ValueSet userInput = toastArgs.UserInput;

                BeginInvoke(new Action(delegate
                {
                    // ����֪ͨ����¼�
                    Show();
                    WindowState = FormWindowState.Normal;
                    Activate();
                    // TODO: UI�̵߳Ĳ���
                    //MessageBox.Show("Toast������������������: " + toastArgs.Argument);
                    _notificationSent = true;
                }));
            };
        }

        private void InitializeTrayIcon()
        {
            // ��������ͼ��
            _notifyIcon = new NotifyIcon
            {
                Icon = new Icon("./logo.ico"), // �滻Ϊ���ͼ���ļ�·��
                Text = "������־",
                Visible = true
            };

            // ��������ͼ��������Ĳ˵�
            _contextMenu = new ContextMenuStrip();
            _notifyIcon.ContextMenuStrip = _contextMenu;

            // ��ӡ���ʾ���˵���
            ToolStripMenuItem showMenuItem = new ToolStripMenuItem("��ʾ");
            showMenuItem.Click += ShowMenuItem_Click;
            _contextMenu.Items.Add(showMenuItem);

            // ��ӡ��˳����˵���
            ToolStripMenuItem exitMenuItem = new ToolStripMenuItem("�˳�");
            exitMenuItem.Click += ExitMenuItem_Click;
            _contextMenu.Items.Add(exitMenuItem);
        }

        private void ShowMenuItem_Click(object sender, EventArgs e)
        {
            // ����������ָ̻���ʾ
            Show();
            WindowState = FormWindowState.Normal;
            Activate();
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            // �˳�Ӧ�ó���
            Application.Exit();
        }

        private void InitializeDailyNotificationTimer()
        {
            // ��ʼ����ʱ��
            _dailyNotificationTimer = new System.Timers.Timer();
            _dailyNotificationTimer.Interval = 1000; // 1��
            _dailyNotificationTimer.Elapsed += DailyNotificationTimer_Elapsed;
            _dailyNotificationTimer.Start();
        }

        private void InitializeReminderTimer()
        {
            // ��ʼ�����Ѷ�ʱ��
            _reminderTimer = new System.Timers.Timer();
            _reminderTimer.Interval = 5 * 60 * 1000; // 5����
            _reminderTimer.Elapsed += ReminderTimer_Elapsed;
        }

        private void DailyNotificationTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            // ��鵱ǰʱ���Ƿ�Ϊÿ������17��15��
            DateTime now = DateTime.Now;
            if (now.Hour == 17 && now.Minute == 15 && now.DayOfWeek != DayOfWeek.Saturday && now.DayOfWeek != DayOfWeek.Sunday)
            {
                // ����֪ͨ
                SendNotification();
                _notificationSent = false;
                _reminderTimer.Start();
            }
        }

        private void ReminderTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            // ��鵱ǰʱ���Ƿ񳬹�18��
            DateTime now = DateTime.Now;
            if (now.Hour < 18 && !_notificationSent)
            {
                // ����֪ͨ
                SendNotification();
                _notificationSent = false;
            }
            else
            {
                // ֹͣ��ʱ��
                _reminderTimer.Stop();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // ������ر�ʱ��С��������
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }

            base.OnFormClosing(e);
        }

        private void SendNotification()
        {
            // ����֪ͨ
            new ToastContentBuilder()
            .AddText("CodeMissing����һ����Ϣ") // �����ı�
            .AddText("������Ϣ���ݣ�����ʱ����")
            .Show(); // 7.0���ϲ��ṩShow����
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = "./�ձ�.xlsx";
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
            Console.WriteLine("�����ѳɹ�׷�ӵ�Excel�ļ��С�");
            MessageBox.Show("�µ�һ�������ˣ���ʰ������׼���°��ˣ�Goodbye��");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SendNotification();
        }
    }
}
