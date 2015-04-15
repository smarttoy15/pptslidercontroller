using System;
using System.Collections.Generic;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.IO.Ports;


namespace pptslidercontroller
{

    /// <summary>
    /// Window1.xaml 的交互逻辑
    /// </summary>
    partial class MySerialPort
    {
        private ThisAddIn m_addIn;

        public MySerialPort()
        {
        }

        //定义 SerialPort对象
        SerialPort port1;

        //初始化SerialPort对象方法.PortName为COM口名称,例如"COM1","COM2"等,注意是string类型
        public void InitCOM(string PortName, ThisAddIn addIn)
        {
            port1 = new SerialPort(PortName);
            port1.BaudRate = 9600;//波特率
            port1.Parity = Parity.None;//无奇偶校验位
            port1.StopBits = StopBits.One;//两个停止位
            port1.Handshake = Handshake.RequestToSend;//控制协议
            port1.ReceivedBytesThreshold = 1;//设置 DataReceived 事件发生前内部输入缓冲区中的字节数
            port1.DataReceived += new SerialDataReceivedEventHandler(port1_DataReceived);//DataReceived事件委托

            m_addIn = addIn;
        }

        //DataReceived事件委托方法
        private void port1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                char msg = '\0';
                if (port1.BytesToRead > 0)
                {
                    msg = (char)port1.ReadByte();
                }

                switch (msg)
                {
                    case 'a':
                    case 'A':
                        m_addIn.Command_GoPreviousSlider();
                        break;
                    case 'd':
                    case 'D':
                        m_addIn.Command_GoNextSlider();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
            }

            /*           try
                       {
                           StringBuilder currentline = new StringBuilder();
                           //循环接收数据
                           while (port1.BytesToRead > 0)
                           {
                               char ch = (char)port1.ReadByte();
                               currentline.Append(ch);
                           }
                           //在这里对接收到的数据进行处理
                           //
                           currentline = new StringBuilder();
                       }
                       catch (Exception ex)
                       {
                           Console.WriteLine(ex.Message.ToString());
                       }
           */
        }

        //打开串口的方法
        public void OpenPort()
        {
            try
            {
                port1.Open();
            }
            catch { }
            if (port1.IsOpen)
            {
                Console.WriteLine("the port is opened!");
            }
            else
            {
                Console.WriteLine("failure to open the port!");
            }
        }

        //关闭串口的方法
        public void ClosePort()
        {
            port1.Close();
            if (!port1.IsOpen)
            {
                Console.WriteLine("the port is already closed!");
            }
        }

        //向串口发送数据
        public void SendCommand(string CommandString)
        {
            byte[] WriteBuffer = Encoding.ASCII.GetBytes(CommandString);
            port1.Write(WriteBuffer, 0, WriteBuffer.Length);
        }
    }

    public partial class ThisAddIn
    {
        private SlideShowWindow m_slidWnd;
        private MySerialPort m_serialPort;

        public void Command_GoNextSlider()
        {
            if (m_slidWnd != null)
            {
                m_slidWnd.View.Next();
            }
        }


        public void Command_GoPreviousSlider()
        {
            if (m_slidWnd != null)
            {
                m_slidWnd.View.Previous();
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SlideShowBegin += Application_SlideShowBegin;
            this.Application.SlideShowEnd += Application_SlideShowEnd;

            m_serialPort = new MySerialPort();
            m_serialPort.InitCOM("COM1", this);
       }

        private void Application_SlideShowEnd(Presentation Pres)
        {
            m_slidWnd = null;

            m_serialPort.ClosePort();
        }

        private void Application_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            m_slidWnd = Wn;

            m_serialPort.OpenPort();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            m_serialPort = null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
