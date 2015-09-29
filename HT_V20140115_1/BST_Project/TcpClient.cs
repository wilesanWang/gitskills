using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;


namespace BST_Project
{
    class TcpClient
    {
        int rc = CommonVary.RUNNING_OK;
        IntPtr ptr;
        GCHandle gch;
        public Socket newclient;
        IPEndPoint ie;
        Exception TCPException;
        int receivedDataLebgth;
        private string address
        {
            get;
            set;
        }
        private int serverPort
        {
            get;
            set;
        }
        private int localPort
        {
            get;
            set;
        }

        byte[] data = new byte[1024];
        public TcpClient(string address, int serverPort, int localPort)
        {
            this.address = address;
            this.serverPort = serverPort;
            this.localPort = localPort;
            rc = ConfigTcpClient();
        }
        /// <summary>
        /// 配置TCPClient
        /// </summary>
        /// <returns></returns>
        public int ConfigTcpClient()
        {
            
            try
            {
                newclient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                newclient.Bind(new IPEndPoint(IPAddress.Any, localPort));
                ie = new IPEndPoint(IPAddress.Parse(address), serverPort);   //服务器IP和端口
                newclient.Connect(ie);   
            }
            catch (System.Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCP初始化出错!\\n出错信息：" + ex.Message);
                //MessageBox.Show("TCP没有初始化");
                return CommonVary.RUNNNING_WRONG;
            }
            
            return CommonVary.RUNNING_OK;
        }

        public int RecevedData()
        {
            if (ConnectPLC(ie) != CommonVary.RUNNING_OK)
                return CommonVary.RUNNNING_WRONG;
            try
            {
                int testSend = newclient.Send(ConverIntToByteAndBigToSmall());
                receivedDataLebgth = newclient.Receive(data);
                ConvertByteData(data);
            }
            catch (System.Net.Sockets.SocketException ex)
            {
                throw ex;
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCPClient连接Server_PLC异常！\n信息：" + ex.Message);
                ConnectPLC(ie);
                return CommonVary.RUNNNING_WRONG;
            }
            catch (System.ArgumentNullException ex)
            {
                throw ex;
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "发送PLC时数据为空！\n信息：" + ex.Message);
                MessageBox.Show("发送PLC时数据为空，请检查日志！");
                return CommonVary.RUNNNING_WRONG;

            }
            catch (System.InvalidOperationException ex)
            {
                throw ex;
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCPClient连接Server_PLC已断开！\n信息：" + ex.Message);
                MessageBox.Show("连接已断开！");
                return CommonVary.RUNNNING_WRONG;

            }
            catch (System.Exception ex)
            {
                throw ex;
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCP数据传输出错！\n信息：" + ex.Message);
                MessageBox.Show("TCP数据传输出错，请检查日志！");
                return CommonVary.RUNNNING_WRONG;

            }
            
            return CommonVary.RUNNING_OK;
        }

        private int ConnectPLC(IPEndPoint ie)
        {
            try
            {
                if (newclient.Connected != true)
                    newclient.Connect(ie);
            }
            catch (System.Net.Sockets.SocketException ex)
            {
                throw ex;
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCPClient连接Server_PLC出错，检查地址和端口信息！\n信息：" + ex.Message);
                MessageBox.Show("TCPClient连接Server_PLC出错，检查地址和端口信息");
                return CommonVary.RUNNNING_WRONG;
            }
            catch (System.InvalidOperationException ex)
            {
                throw ex;
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCPClient连接Server_PLC已断开！\n信息：" + ex.Message);
                MessageBox.Show("连接已断开！");
                return CommonVary.RUNNNING_WRONG;

            }
            catch (System.Exception ex)
            {
                throw ex;
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCP数据传输出错！\n信息：" + ex.Message);
                MessageBox.Show("TCP数据传输出错，请检查日志！");
                return CommonVary.RUNNNING_WRONG;

            }
            return CommonVary.RUNNING_OK;
        }

        private byte[] ConverIntToByteAndBigToSmall()
        {
            int size = Marshal.SizeOf(MainForm.sendMapData);
            Int16[] tempArray = new Int16[size / 2];
            byte[] bf = new byte[size];
            try
            {
                gch = GCHandle.Alloc(bf, GCHandleType.Pinned);
                ptr = gch.AddrOfPinnedObject();
                //ptr = Marshal.AllocHGlobal(size);
                Marshal.StructureToPtr(MainForm.sendMapData, ptr, false);
                Marshal.Copy(ptr, tempArray, 0, size / 2);
                byte[] tempbf = new byte[2];
                int j = 0;
                for (int i = 0; i < size / 2; i++)
                {
                    tempbf = BitConverter.GetBytes(tempArray[i]);
                    bf[j] = tempbf[1];
                    bf[j + 1] = tempbf[0];
                    //Console.Write(bf[i] + " ");
                    j = j + 2;
                }
                gch.Free();
            }
            catch (System.Exception ex)
            {
                gch.Free();
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "Struct映射到Int[]出错！\n信息：" + ex.Message);
                MessageBox.Show("Struct映射到Int[]出错，请查询日志信息");
                return null;
            }
            return bf;
        }
        /// <summary>
        /// 原数据转换包括大小端转换，小端转换大端并存入变量
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public int ConvertByteData(byte[] data)
        {
            
            byte[] tempData = new byte[receivedDataLebgth];
            try
            {
                for (int i = 0; i < receivedDataLebgth; i = i + 2)
                {
                    tempData[i] = data[i + 1];
                    tempData[i + 1] = data[i];
                    CommonVary.data[i / 2] = BitConverter.ToInt16(tempData, i);
                }
                gch = GCHandle.Alloc(CommonVary.data, GCHandleType.Pinned);
                ptr = gch.AddrOfPinnedObject();
                Marshal.Copy(CommonVary.data, 0, ptr, CommonVary.data.Length);
                MainForm.receiveMapData = (MainForm.Map_data)Marshal.PtrToStructure(ptr, typeof(MainForm.Map_data));
                gch.Free();

                //for (int i = 0; i < CommonVary.data.Length; i++)
                //{
                //    Console.Write(CommonVary.data[i].ToString() + " ");
                //}
                //Console.WriteLine();
            }
            catch (IndexOutOfRangeException indexOutex)
            {
                gch.Free();
                MessageBox.Show("大小端转换出错");
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "大小端转换出错\n信息： " + indexOutex.Message);
                return CommonVary.RUNNNING_WRONG;
            }
            catch (System.Exception ex)
            {
                gch.Free();
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "变量映射到Struct出错\n信息" + ex.Message);
                return CommonVary.RUNNNING_WRONG;
            }
            return CommonVary.RUNNING_OK;
        }
    }
}
