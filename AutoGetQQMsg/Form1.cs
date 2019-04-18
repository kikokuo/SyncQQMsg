using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using UIAutomationBlockingCoreLib;


namespace AutoGetQQMsg
{
    public partial class Form1 : Form
    {
        private IUIAutomation _automation; // Main UIA object required by any UIA client app.
        private int _propertyIdName = 30005; // UIA_NamePropertyId
        private int patternId = 10018;       //IUIAutomationLegacyIAccessiblePattern
        private int count = 0;
        private string title;
        const int MOUSEEVENTF_MOVE = 0x0001; //移动鼠标
        const int MOUSEEVENTF_LEFTDOWN = 0x0002; //模拟鼠标左键按下
        const int MOUSEEVENTF_LEFTUP = 0x0004; //模拟鼠标左键抬起
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008; //模拟鼠标右键按下
        const int MOUSEEVENTF_RIGHTUP = 0x0010; //模拟鼠标右键抬起
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020; //模拟鼠标中键按下
        const int MOUSEEVENTF_MIDDLEUP = 0x0040; //模拟鼠标中键抬起
        const int MOUSEEVENTF_ABSOLUTE = 0x8000; //标示是否采用绝对坐标
        private List<string> list = new List<string>()
        {
            "VIVEPORT 官方小助手","VIVEPORT 一体机玩家交流群","Viveport合作玩家群","VIVEPORT PC端玩家交流群"
        };
        public Form1()
        {
            InitializeComponent();
            _automation = new CUIAutomation();
        }
        public void Uninitialize()
        {
            if (_automation != null)
            {
                _automation = null;
            }
        }
        public static void LeftClick()
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
        }

        [DllImport("user32.dll")]
        public static extern void mouse_event(UInt32 dwFlags, UInt32 dx, UInt32 dy, UInt32 dwData, IntPtr dwExtraInfo);
        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_automation != null)
            {
                _automation = null;
            }
            timer1.Stop();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (count < 4)
            {
                title = list[count];
                count++;
            }
            else
            {
                count = 0;
                title = list[count];
                count++;
            }
            //找到QQ
            Process[] process = Process.GetProcessesByName("QQ");

            foreach (Process p in process)
            {
                textBox1.Clear();
                //從QQ產生root automation
                IUIAutomationElement _viveportAutomationElement = _automation.ElementFromHandle(
                                                                p.MainWindowHandle);

                if (_viveportAutomationElement == null)
                {
                    throw new InvalidOperationException("QQ must be running");
                }
                //依據屬性名稱 找出tabbox
                IUIAutomationCondition conditionNote = _automation.CreatePropertyCondition(
                                                            _propertyIdName, title);

                IUIAutomationElement _tabBoxAutomationElement = _viveportAutomationElement.FindFirst(
                                                        TreeScope.TreeScope_Descendants,
                                                        conditionNote);
                if (_tabBoxAutomationElement == null)
                {
                    throw new InvalidOperationException("Could not find "+ title);
                }
                //由tab找出按鍵範圍 移動游標點擊
                tagPOINT point = new tagPOINT();
                _tabBoxAutomationElement.GetClickablePoint(out point);
                Cursor.Position = new Point((int)point.x, (int)point.y);
                LeftClick();
                Thread.Sleep(150);
                //找出QQ聊天室顯示框
                conditionNote = _automation.CreatePropertyCondition(
                                                            _propertyIdName, "消息");

                IUIAutomationElement _resultTextBoxAutomationElement = _viveportAutomationElement.FindFirst(
                                                            TreeScope.TreeScope_Descendants,
                                                            conditionNote);

                if (_resultTextBoxAutomationElement == null)
                {
                    throw new InvalidOperationException("Could not find 消息");
                }
                //取出消息
                UIAutomationClient.IUIAutomationLegacyIAccessiblePattern legacyPattern1 = (UIAutomationClient.IUIAutomationLegacyIAccessiblePattern)_resultTextBoxAutomationElement.GetCurrentPattern(patternId);

                //放到texebox
                List<char> list = legacyPattern1.CurrentValue.ToList();
                foreach (char i in list)
                {
                    if (i != '\r')
                        textBox1.AppendText(i.ToString());
                    else
                        textBox1.AppendText(Environment.NewLine);
                }                   
            }
        }
    }
}
