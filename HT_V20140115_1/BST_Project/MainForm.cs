using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using ZedGraph;
using System.Text.RegularExpressions;
using System.IO.Ports;

namespace BST_Project
{
    public partial class MainForm : Form
    {
#region 全局参数
        const int INITIALIZATIONERRRO = -1;
        const int EXECUTERIGHT = 0;

        bool readPortFlag = false;
        bool saveFlag = false;
        bool startMach = false;
        bool startSystem = false;

        public static Map_data receiveMapData;
        public static Map_data sendMapData;

        bool displayOrSetFlag = false;//false表示显示，true表示设置
        bool TCPInitialized = false;

        List<Point> WideLocation= new List<Point> { };
        int h1, w1;
        Point location;
        PointPairList diagramList1 = new PointPairList();
        PointPairList diagramList2 = new PointPairList();
        PointPairList diagramList3 = new PointPairList();
        PointPairList diagramList4 = new PointPairList();
        PointPairList diagramListStandard = new PointPairList();
        PointPairList diagramListUpper = new PointPairList();
        PointPairList diagramListLower = new PointPairList();


        PointPairList rediagramList1 = new PointPairList();
        PointPairList rediagramList2 = new PointPairList();
        PointPairList rediagramList3 = new PointPairList();
        PointPairList rediagramList4 = new PointPairList();
        PointPairList rediagramListStandard = new PointPairList();
        PointPairList rediagramListUpper = new PointPairList();
        PointPairList rediagramListLower = new PointPairList();

        GraphPane myPane1, myPane2, myPane3, myPane4, myPane5;

        TcpClient tcpClient;

        DataTable dtDataGridView, dtDataExcel;
        string saveFileName = "";

        Color PanelBorderColor = Color.Gray;
        Color ColorBlue = Color.Blue;
        Color ColorOrange = Color.Silver;

        string standardWidth, upperErrorWidth, lowerErrorWidth,upperWarWidth,lowerWarWidth;
        string sName, sAngle;
        string standardHisWidth, upperHisWidth, LowerHisWidth;

        Thread ToExcelThread, ToGridViewThread;
        int toGridViewSucc = CommonVary.RUNNNING_WRONG;

        MySqlCommand mySqlCommand;
        int dataGridView2SelecteRow = -1;
        int progressId;

        enum ColumnName { A1 = 1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, U1, V1, W1, X1, Y1, Z1 };

        string[] wStatusDisplay = new string[]{
            "没有标定任务",
            "标定过程中",
            "继续标定、标定完成",
            "标定孔数不足",
            "标定内部错误",
            "相机不存在",
            "CAN总线SDO信息错误",
            "扫描结束但标定错误",
            "扫描失败"
        };
#endregion
        
#region 主界面操作
        public MainForm()
        {

            int rc = CommonVary.RUNNING_OK;
            InitializeComponent();
            GetDefaultLocation();
            rc = InitializeSpecialPreperty();
            RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "开始运行！");
            if (rc != CommonVary.RUNNING_OK)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "初始化特殊属性失败。");
            }
            else
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "初始化特殊属性成功");
            }
            rc = InitalizeControl();
            if (rc != CommonVary.RUNNING_OK)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "初始化控件失败。");
            }
            else
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "初始化控件成功");
            }

            TMTCP.Enabled = false;
            // rc = InitalizeTcp();
            if (rc != CommonVary.RUNNING_OK)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "初始化TCP连接失败。");
            }
            else
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "初始化TCP成功");
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (MyPort.IsOpen)
                MyPort.Close();
            RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "退出成功！\n");
        }
#endregion
        
#region 初始化相关

        


        /// <summary>
        /// 初始化控件属性
        /// </summary>
        /// <returns></returns>
        private int InitializeSpecialPreperty()
        {
            try
            {
                //Solid MainForm
                RecordLog._CreateNewLog();
                //this.FormBorderStyle = FormBorderStyle.FixedSingle;
                //this.MaximizeBox = false;
                if (CommonVary.ReadConfig() == CommonVary.RUNNING_OK)
                {
                    initialProcessbar();
                    InitalHandle();
                    InitialZedgraph();
                    InitalizeTcp();
                    InitialSetInfo();
                    InitialWriteCOM();

                }
       
                
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("工程初始化失败");
                return INITIALIZATIONERRRO;
            }

            return EXECUTERIGHT;

        }

        private void InitialSetInfo()
        {
            textBox30.Text = CommonVary.PLCAddress;
            textBox39.Text = CommonVary.PLCPort.ToString();
            textBox43.Text = CommonVary.PCFromPLCPort.ToString();
            textBox44.Text = (CommonVary.SaveDataInterval / 1000).ToString();
            comboBox7.Items.Clear();
           
            textBox52.Text = CommonVary.COMBaunRate.ToString();
            string[] COMList = SerialPort.GetPortNames();
            foreach (string name in COMList)
            {
                comboBox7.Items.Add(name);
            }
            comboBox7.Text = CommonVary.myCOMPort;
            
        }

        

        private int InitialZedgraph()
        {
            try
            {
                myPane1 = zedGraphControl1.GraphPane;
                myPane1.Title.Text = "实时数据曲线";
                myPane1.XAxis.Title.Text = "日期";
                myPane1.YAxis.Title.Text = "宽度";
                myPane1.XAxis.Type = AxisType.Date;

                myPane2 = zedGraphControl2.GraphPane;
                myPane2.Title.Text = "实时数据曲线";
                myPane2.XAxis.Title.Text = "日期";
                myPane2.YAxis.Title.Text = "宽度";
                myPane2.XAxis.Type = AxisType.Date;

                myPane3 = zedGraphControl3.GraphPane;
                myPane3.Title.Text = "实时数据曲线";
                myPane3.XAxis.Title.Text = "日期";
                myPane3.YAxis.Title.Text = "宽度";
                myPane3.XAxis.Type = AxisType.Date;

                myPane4 = zedGraphControl4.GraphPane;
                myPane4.Title.Text = "实时数据曲线";
                myPane4.XAxis.Title.Text = "日期";
                myPane4.YAxis.Title.Text = "宽度";
                myPane4.XAxis.Type = AxisType.Date;

                myPane5 = zedGraphControl5.GraphPane;
                myPane5.Title.Text = "历史数据曲线";
                myPane5.XAxis.Title.Text = "日期";
                myPane5.YAxis.Title.Text = "宽度";
                myPane5.XAxis.Type = AxisType.Date;
                zedGraphControl1.IsEnableHZoom = false;
                zedGraphControl2.IsEnableHZoom = false;
                zedGraphControl3.IsEnableHZoom = false;
                zedGraphControl4.IsEnableHZoom = false;

                zedGraphControl1.AxisChange();
                zedGraphControl2.AxisChange();
                zedGraphControl3.AxisChange();
                zedGraphControl4.AxisChange();
                
                rediagramList1.Clear();
                rediagramList2.Clear();
                rediagramList3.Clear();
                rediagramList4.Clear();
                rediagramListStandard.Clear();
                rediagramListLower.Clear();
                rediagramListUpper.Clear();
                myPane1.CurveList.Clear();
                myPane2.CurveList.Clear();
                myPane3.CurveList.Clear();
                myPane4.CurveList.Clear();
                CurveItem myCurve1standard = myPane1.AddCurve("标准宽度", rediagramListStandard, Color.Black, SymbolType.None);
                CurveItem myCurve1 = myPane1.AddCurve("宽度1", rediagramList1, Color.Blue, SymbolType.None);                
                CurveItem myCurve1upper = myPane1.AddCurve("宽度上限", rediagramListUpper, Color.Red, SymbolType.None);
                CurveItem myCurve1lower = myPane1.AddCurve("宽度下限", rediagramListLower, Color.Brown, SymbolType.None);

                CurveItem myCurve2standard = myPane2.AddCurve("标准宽度", rediagramListStandard, Color.Black, SymbolType.None);
                CurveItem myCurve2 = myPane2.AddCurve("宽度2", rediagramList2, Color.Blue, SymbolType.None);
                CurveItem myCurve2upper = myPane2.AddCurve("宽度上限", rediagramListUpper, Color.Red, SymbolType.None);
                CurveItem myCurve2lower = myPane2.AddCurve("宽度下限", rediagramListLower, Color.Brown, SymbolType.None);

                CurveItem myCurve3standard = myPane3.AddCurve("标准宽度", rediagramListStandard, Color.Black, SymbolType.None);
                CurveItem myCurve3 = myPane3.AddCurve("宽度3", rediagramList3, Color.Blue, SymbolType.None);
                CurveItem myCurve3upper = myPane3.AddCurve("宽度上限", rediagramListUpper, Color.Red, SymbolType.None);
                CurveItem myCurve3lower = myPane3.AddCurve("宽度下限", rediagramListLower, Color.Brown, SymbolType.None);

                CurveItem myCurve4standard = myPane4.AddCurve("标准宽度", rediagramListStandard, Color.Black, SymbolType.None);
                CurveItem myCurve4 = myPane4.AddCurve("宽度4", rediagramList4, Color.Blue, SymbolType.None);
                CurveItem myCurve4upper = myPane4.AddCurve("宽度上限", rediagramListUpper, Color.Red, SymbolType.None);
                CurveItem myCurve4lower = myPane4.AddCurve("宽度下限", rediagramListLower, Color.Brown, SymbolType.None);
            }
            catch (System.Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "Zedgraph初始化失败！\n信息：" + ex.Message);
                return CommonVary.RUNNNING_WRONG;
            }
            return CommonVary.RUNNING_OK;


        }

        /// <summary>
        /// 事件添加
        /// </summary>
        /// <returns></returns>
        private int InitalHandle()
        {
            this.textBox11.GotFocus += new EventHandler(this.textBox11_GotFocus);
            this.textBox20.GotFocus += new EventHandler(this.textBox20_GotFocus);

            return CommonVary.RUNNING_OK;
        }

        private void InitalizeTcp()
        {
            try
            {
                tcpClient = new TcpClient(CommonVary.PLCAddress,CommonVary.PLCPort,CommonVary.PCFromPLCPort);
                TCPInitialized = true;
            }
            catch (Exception e)
            {
                TCPInitialized = false;
            }
                
        }

        /// <summary>
        /// 初始化窗体控件
        /// </summary>
        /// <returns></returns>
        private int InitalizeControl()
        {

            /******************************WWM Configuration*****************************/
            comboBox4.Items.Add(CommonVary.cbx4Item1);
            comboBox4.Items.Add(CommonVary.cbx4Item2);
            comboBox4.Text = CommonVary.cbx4Item1;
            sendMapData.wWM_config.wStrip_CCD = (Int16)(sendMapData.wWM_config.wStrip_CCD & 0xfff0);

            comboBox5.Items.Add(CommonVary.cbx5Item1);
            comboBox5.Items.Add(CommonVary.cbx5Item2);
            comboBox5.Text = CommonVary.cbx5Item1;
            sendMapData.wWM_config.wStrip_CCD = (Int16)(sendMapData.wWM_config.wStrip_CCD & 0x000D);

            comboBox3.Items.Add(CommonVary.cbx3Item1);
            comboBox3.Items.Add(CommonVary.cbx3Item2);
            comboBox3.Text = CommonVary.cbx3Item1;
            sendMapData.sys_Inform.NumPos_NumWeb = (Int16)(sendMapData.sys_Inform.NumPos_NumWeb & 0xfff0);

            comboBox6.Items.Add(CommonVary.cbx6Item1);
            comboBox6.Items.Add(CommonVary.cbx6Item2);
            comboBox6.Text = CommonVary.cbx6Item1;
            sendMapData.sys_Inform.NumPos_NumWeb = (Int16)(sendMapData.sys_Inform.NumPos_NumWeb & 0x000D);

            label11.Visible = false;

            label35.Visible = false;
            label36.Visible = false;
            label37.Visible = false;
            label38.Visible = false;
            textBox23.Visible = false;
            textBox10.Visible = false;
            textBox26.Visible = false;
            textBox14.Visible = false;
            textBox24.Visible = false;
            textBox9.Visible = false;
            textBox25.Visible = false;
            textBox13.Visible = false;

            panel6.Location = new Point(220, 12);
            panel7.Location = new Point(220, 12);
            panel8.Location = new Point(220, 12);
            /************************************************************************/
            /***********************Combobox*******************************/
            string[] groupInfo = new string[] { "早班", "中班", "晚班" };
            foreach (string str in groupInfo)
            {
                CBGroup.Items.Add(str);
                CBHisGroup.Items.Add(str);
                comboBox2.Items.Add(str);
            }

            try
            {
                CommonVary.OpenDataConnection();
                using (MySqlDataAdapter myDataAdapter = new MySqlDataAdapter("select sName from specificationdata order by sName", CommonVary.conn))
                {
                    using (DataTable dt = new DataTable("specification"))
                    {
                        myDataAdapter.Fill(dt);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            CBSpecification.Items.Add(dt.Rows[i][0].ToString());
                            CBHisSpecification.Items.Add(dt.Rows[i][0].ToString());
                            comboBox1.Items.Add(dt.Rows[i][0].ToString());
                        }
                    }
                    comboBox1.Items.Add("ALL");
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("规格初始化错误");
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "规格初始化错误!\n信息" + ex.Message);
                return CommonVary.RUNNNING_WRONG;
            }
            CBSpecification.Text = CBSpecification.Items[0].ToString();
            CBHisSpecification.Text = CBHisSpecification.Items[0].ToString();
            comboBox1.Text = comboBox1.Items[0].ToString();
            /**************************************************************************************************/

            /***************************************获取默认规格宽度********************************************/
            if (CommonVary.RUNNNING_WRONG == GetStandardWidth(CBSpecification.Text))
            {
                return CommonVary.RUNNNING_WRONG;
            }

            CBGroup.Text = CBGroup.Items[0].ToString();
            CBHisGroup.Text = CBHisGroup.Items[0].ToString();
            comboBox2.Text = comboBox2.Items[0].ToString();
            /*************************************************************************************************/
            /********************************初始化datatimepicker*********************************************/
            dateTimePicker1.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker3.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            dateTimePicker4.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker4.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            /***************************************************************************************************/

            /*********************************Camera设置************************************************/
            radioButton1.Checked = true;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
            panel6.Visible = true;
            panel7.Visible = false;
            panel8.Visible = false;
            button23.Visible = false;
            button24.Visible = false;

            /******************************************************************************************/
            /************************************绘图初始化**********************************************/
            h1 = splitContainer1.Panel2.Height - 2;
            w1 = splitContainer1.Panel2.Width - 2;

            location = new Point(3, 2);
            SetZedGraphControlDefaultStyle();
            //Point p1 = splitContainer1.
            /*****************************************************************************************/
            return CommonVary.RUNNING_OK;
        }
        private void SetZedGraphControlDefaultStyle()
        {
           
            zedGraphControl1.Location = location;
            zedGraphControl1.Height = h1 / 2;
            zedGraphControl1.Width = w1 / 2;
            zedGraphControl2.Location = new Point(location.X + w1 / 2, location.Y);
            zedGraphControl3.Location = new Point(location.X, location.Y + h1 / 2);
            zedGraphControl4.Location = new Point(location.X + w1 / 2, location.Y + h1 / 2);
            zedGraphControl2.Height = h1 / 2;
            zedGraphControl2.Width = w1 / 2;
            zedGraphControl3.Height = h1 / 2;
            zedGraphControl3.Width = w1 / 2;
            zedGraphControl4.Height = h1 / 2;
            zedGraphControl4.Width = w1 / 2;
            zedGraphControl1.Visible = true;
            zedGraphControl2.Visible = true;
            zedGraphControl3.Visible = true;
            zedGraphControl4.Visible = true;

        }

        /// <summary>
        /// 初始化导出进度条
        /// </summary>
        private void initialProcessbar()
        {
            progressBar1.Visible = false;
            label5.Visible = true;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Step = 5;
            TMProcessbar.Interval = 100;
            TMProcessbar.Enabled = false;

            RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "查询数据进度条初始化成功！");
        }
#endregion
        
#region 数据结构定义
        public struct DataExchange_Header
        {
            public Int16 wCounterToBST;
            public Int16 wCounterFromBST;
            public Int16 uiSoftwareNumber;
            public Int16 uiSoftwareVersion;
            public Int16 uiMappingType;
            public Int16 eAppID;
            public Int16 uiFrameSize;
            public Int16 wStatus;
        };
        public struct WWM_config
        {
            public Int16 iWidth_No;
            public Int16 wStrip_CCD;
            public Int16 iCCD1_No;
            public Int16 iCCD2_No;
            public Int16 iCCD1_EdgeNo;
            public Int16 iCCD2_EdgeNo;
            public Int16 iStripNo;
            public Int16 iDulCCD_refwidth;
            public Int16 iSet_bit;
            public Int16 reversed;
        };
        public struct WWM_Parameter1
        {
            public Int16 SetWebwidth;
            public Int16 SetErrTol_upper;
            public Int16 SetErrTol_lower;
            public Int16 SetWarTol_upper;
            public Int16 SetWarTol_lower;
            public Int16 reversed;
        };
        public struct WWM_Parameter2
        {
            public Int16 SetWebwidth;
            public Int16 SetErrTol_upper;
            public Int16 SetErrTol_lower;
            public Int16 SetWarTol_upper;
            public Int16 SetWarTol_lower;
            public Int16 reversed;
        };
        public struct WWM_Parameter3
        {
            public Int16 SetWebwidth;
            public Int16 SetErrTol_upper;
            public Int16 SetErrTol_lower;
            public Int16 SetWarTol_upper;
            public Int16 SetWarTol_lower;
            public Int16 reversed;
        };
        public struct WWM_Parameter4
        {
            public Int16 SetWebwidth;
            public Int16 SetErrTol_upper;
            public Int16 SetErrTol_lower;
            public Int16 SetWarTol_upper;
            public Int16 SetWarTol_lower;
            public Int16 reversed;
        };
        public struct WWM_Parameter5
        {
            public Int16 SetWebwidth;
            public Int16 SetErrTol_upper;
            public Int16 SetErrTol_lower;
            public Int16 SetWarTol_upper;
            public Int16 SetWarTol_lower;
            public Int16 reversed;
        };
        public struct WWM_Parameter6
        {
            public Int16 SetWebwidth;
            public Int16 SetErrTol_upper;
            public Int16 SetErrTol_lower;
            public Int16 SetWarTol_upper;
            public Int16 SetWarTol_lower;
            public Int16 reversed;
        };
        public struct CCD_Info
        {
            public Int16 iCCDno;//(RW)
            public Int16 state;
            public Int16 edge1;
            public Int16 edge2;
            public Int16 edge3;
            public Int16 edge4;
            public Int16 webwidth;
            public Int16 focus;
            public Int16 exposuretime;
            public Int16 resolution;//(pix/mm)
            public Int16 zero_pix;
            public Int16 CAN_Address;//(RW)
        };
        public struct CCD_Cal
        {
            public Int16 controlInt16;
            //Bit0: bStart    （WO)
            //Bit1: Bok       （WO)
            //Bit2: bCancel   （WO)
            //Bit3: bNext     （WO)
            //Bit4: bReScan   （WO)
            //Bit5: Spare     
            //Bit6: Spare     
            //Bit7: bSetType  （WO)
            //Bit8: bSetLevel （WO)
            //Bit9: Spare
            //Bit10: Spare
            //Bit11: Spare
            //Bit12: Spare
            //Bit13: Spare
            //Bit14: Spare
            //Bit15: Spare
            public Int16 iCalType;	//(RW)
            public Int16 rRefWidth; //0.1mm(WR)
            public Int16 rCCDdistance; //0.1mm(WR)
            public Int16 rLowerEdgeFromCalSheet; //0.1mm(WR)
            public Int16 rThicknessFromLowerEdge; //0.1mm(WR)
            public Int16 wStatus;//(RO)    
            public Int16 iCCDpixelCount;//(RO)
            public Int16 iBorderCount;//(RO)
            public Int16 iScanInProgress;//(RO)
            public Int16 iProgress;//(RO)
            public Int16 iText;//(RO)
            public Int16 iState;//(RO)
            public Int16 iCenterPix;//(RO)
            public Int16 iScanPoInt16s;  //(RO)
            public Int16 iEvalPoInt16s;  //(RO)
            public Int16 iPerpPixHole;//(RO)
            public Int16 iPerpPixBridge;//(RO)
        };
        public struct Width1
        {
            public Int16 wStatus;
            public Int16 webwidth;
            public Int16 Splice1;
            public Int16 Splice2;
        };
        public struct Width2
        {
            public Int16 wStatus;
            public Int16 webwidth;
            public Int16 Splice1;
            public Int16 Splice2;
        };
        public struct Width3
        {
            public Int16 wStatus;
            public Int16 webwidth;
            public Int16 Splice1;
            public Int16 Splice2;
        };
        public struct Width4
        {
            public Int16 wStatus;
            public Int16 webwidth;
            public Int16 Splice1;
            public Int16 Splice2;
        };
        public struct Width5
        {
            public Int16 wStatus;
            public Int16 webwidth;
            public Int16 Splice1;
            public Int16 Splice2;
        };
        public struct Width6
        {
            public Int16 wStatus;
            public Int16 webwidth;
            public Int16 Splice1;
            public Int16 Splice2;
        };
        public struct Sys_Inform
        {
            public Int16 NumPos_NumWeb;
            public Int16 width1_lower_limit_analog_output;
            public Int16 width1_upper_limit_analog_output;
            public Int16 width2_lower_limit_analog_output;
            public Int16 width2_upper_limit_analog_output;
            public Int16 Setbit;
        };
        public struct Map_data
        {
            public DataExchange_Header dataExchange_Header;
            public WWM_config wWM_config;
            public WWM_Parameter1 wWM_Parameter1;
            public WWM_Parameter2 wWM_Parameter2;
            public WWM_Parameter3 wWM_Parameter3;
            public WWM_Parameter4 wWM_Parameter4;
            public WWM_Parameter5 wWM_Parameter5;
            public WWM_Parameter6 wWM_Parameter6;
            public CCD_Info cCD_Info;
            public CCD_Cal cCD_Cal;
            public Width1 width1;
            public Width2 width2;
            public Width3 width3;
            public Width4 width4;
            public Width5 width5;
            public Width6 width6;
            public Sys_Inform sys_Inform;

        };

        public struct Send_Map_data
        {
            public DataExchange_Header dataExchange_Header;
            public WWM_config wWM_config;
            public WWM_Parameter1 wWM_Parameter1;
            public WWM_Parameter2 wWM_Parameter2;
            public WWM_Parameter3 wWM_Parameter3;
            public WWM_Parameter4 wWM_Parameter4;
            public WWM_Parameter5 wWM_Parameter5;
            public WWM_Parameter6 wWM_Parameter6;
            public CCD_Info cCD_Info;
            public CCD_Cal cCD_Cal;
            public Width1 width1;
            public Width2 width2;
            public Width3 width3;
            public Width4 width4;
            public Width5 width5;
            public Width6 width6;
            public Sys_Inform sys_Inform;

        };
#endregion

#region panel边框颜色定义
  
        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                                splitContainer1.Panel1.ClientRectangle,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid);
        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                                splitContainer1.Panel2.ClientRectangle,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid);
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                                    panel2.ClientRectangle,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid);
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                                panel3.ClientRectangle,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid);
        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                                    panel4.ClientRectangle,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid);
        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                                    panel5.ClientRectangle,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid,
                                    PanelBorderColor,
                                    3,
                                    ButtonBorderStyle.Solid);
        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                                panel6.ClientRectangle,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid);
        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                                panel7.ClientRectangle,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid);
        }

        private void panel8_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                                panel8.ClientRectangle,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid,
                                PanelBorderColor,
                                3,
                                ButtonBorderStyle.Solid);
        }

        private void panel9_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                               panel9.ClientRectangle,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid);
        }

        private void panel10_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                               panel10.ClientRectangle,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid);
        }

        private void panel11_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                               panel11.ClientRectangle,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid);
        }

        private void panel12_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics,
                               panel12.ClientRectangle,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid,
                               PanelBorderColor,
                               3,
                               ButtonBorderStyle.Solid);
        }

#endregion

#region 历史数据导出
        private DataTable GetHistryData(string iniSql)
        {
            using (DataTable dt = new DataTable())
            {
                try
                {
                    CommonVary.OpenDataConnection();
                    using (MySqlDataAdapter myData = new MySqlDataAdapter(iniSql, CommonVary.conn))
                    {
                        myData.Fill(dt);
                        if (dt.Rows.Count == 0)
                        {
                            return null;
                        }
                        dt.Columns[0].ColumnName = "测试时间";
                        dt.Columns[1].ColumnName = "宽度1";
                        dt.Columns[2].ColumnName = "宽度2";
                        dt.Columns[3].ColumnName = "宽度3";
                        dt.Columns[4].ColumnName = "宽度4";
                        dt.Columns[5].ColumnName = "标准宽度";

                    }
                    return dt;

                }
                catch (System.Exception ex)
                {

                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "查询历史数据出错！\n信息：" + ex.Message);
                    return null;
                }
            }


        }
        
        /// <summary>
        /// 查询历史数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "ALL")
            {
                MessageBox.Show("不可以通过窗口查询全部数据！\n请选用导出Excel表格查询！");
                return;
            }
            try
            {
                string countSql = "select count(*) from widthdata as w where w.timestamps >= '" +
                    dateTimePicker3.Text + "' " + "and w.timestamps <= '" + dateTimePicker4.Text + "' and w.specification ='" + comboBox1.Text.Trim() + "'";
                using (DataTable dt = new DataTable())
                {
                    CommonVary.OpenDataConnection();
                    using (MySqlDataAdapter myData = new MySqlDataAdapter(countSql, CommonVary.conn))
                    {
                        myData.Fill(dt);
                    }
                    if (dt == null)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGWARNING, "历史数据记录为空！\n信息" + dateTimePicker3.Text + "至" + dateTimePicker4.Text + "数据为空");
                        MessageBox.Show("选择区段记录为空");
                    }
                    else if (int.Parse(dt.Rows[0][0].ToString()) > 10000)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGWARNING, "历史数据记录大于10000条！\n信息" + dateTimePicker3.Text + "至" + dateTimePicker4.Text + "数据条数大于10000");
                        MessageBox.Show("历史数据记录大于10000条！\n请选用Excel导出");
                    }
                    else
                    {


                        dataGridView1.DataSource = null;
                        string selectSql = "select w.timestamps as realtime,w.width1 as w1,w.width2 as w2,w.width3 as w3," +
                            "w.width4 as w4,s.sWidth as ws from widthdata as w,specificationdata as s where w.timestamps >= '" +
                            dateTimePicker3.Text + "' " + "and w.timestamps <= '" + dateTimePicker4.Text + "' and s.sName ='" + comboBox1.Text + "' and w.specification = s.sName";

                        dtDataGridView = GetHistryData(selectSql);
                        if (dtDataGridView == null)
                        {
                            RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGWARNING, "历史数据记录为空！\n信息" + dateTimePicker3.Text + "至" + dateTimePicker4.Text + "数据为空");
                            MessageBox.Show("选择区段记录为空");
                        }
                        else
                        {
                            //ToGridViewThread = new Thread(new ThreadStart(dataGridView1.DataSource = dt));
                            //label5.Text = "正在查询到窗口中";
                            //progressBar1.Visible = true;
                            //progressBar1.Step = 10;
                            //TMProcessbar.Enabled = true;


                            label5.Text = "正在查询到窗口中";
                            Application.DoEvents();
                            //progressBar1.Visible = true;
                            //TMProcessbar.Enabled = true;


                            toGridViewSucc = OutputDataGridView();

                            label5.Text = "没有查询任务";
                            //progressBar1.Visible = false;
                            //TMProcessbar.Enabled = false;
                            dataGridView1.ClearSelection();
                            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                            RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "查询历史数据成功！\n信息" + dateTimePicker3.Text + "至" + dateTimePicker4.Text + "数据");
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "历史数据查询出错！");
                MessageBox.Show("历史数据查询出错!");
                return;
            }
        }


        private int OutputDataGridView()
        {
            dataGridView1.DataSource = dtDataGridView;
            return CommonVary.RUNNING_OK;
        }

        //private void ToGriView(DataTable dt)
        //{

        //}

        private void button8_Click(object sender, EventArgs e)
        {
            string selectSql;
            try
            {
                if (comboBox1.Text == "ALL")
                {
                    selectSql = "select w.timestamps as realtime,w.width1 as w1,w.width2 as w2,w.width3 as w3," +
                        "w.width4 as w4,s.sWidth as ws,w.specification as wsp  from widthdata as w,specificationdata as s where timestamps >= '" +
                        dateTimePicker3.Text + "' " + "and timestamps <= '" + dateTimePicker4.Text + "'and s.sname = w.specification";
                }
                else
                {
                    selectSql = "select w.timestamps as realtime,w.width1 as w1,w.width2 as w2,w.width3 as w3," +
                        "w.width4 as w4,s.sWidth as ws,w.specification as wsp from widthdata as w,specificationdata as s where w.timestamps >= '" +
                        dateTimePicker3.Text + "' " + "and w.timestamps <= '" + dateTimePicker4.Text + "' and s.sName ='" + comboBox1.Text + "' and s.sName = w.specification" ;
                }
                    //string iniSql = @"select realtime,w1,w2,w3,w4,ws from widthview where substring(realtime,1,19) >= '" + dateTimePicker3.Text + "' " +
                //                "and substring(realtime,1,19) <= '" + dateTimePicker4.Text + "' and s ='" + comboBox1.Text + "'";
                if (dtDataExcel != null)
                    dtDataExcel.Clear();
                dtDataExcel = GetHistryData(selectSql);
                if (dtDataExcel == null)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGWARNING, "历史数据记录为空！\n信息" + dateTimePicker3.Text + "至" + dateTimePicker4.Text + "数据为空");
                    MessageBox.Show("选择区段记录为空");
                }
                else
                {
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.DefaultExt = "xls|xlsx";
                    saveDialog.Filter = "Excel文件|*.xls|Excel文件|*.xlsx";
                    saveDialog.FileName = "Sheet" + DateTime.Today.ToString("yyyy-MM-dd");
                    saveDialog.ShowDialog();
                    saveFileName = saveDialog.FileName;

                    if (saveFileName.IndexOf(":") < 0)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "没有选择路径!");
                        return;
                    }
                    else
                    {
                        label5.Text = "正在查询到Excel中";

                        progressBar1.Visible = true;
                        TMProcessbar.Enabled = true;
                        ToExcelThread = new Thread(new ThreadStart(OutputToExcel));
                        ToExcelThread.Start();

                    }

                }
            }
            catch (System.Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "历史数据查询Excel出错！");
                MessageBox.Show("历史数据查询到Excel出错!");
                return;
            }
        }

        /// <summary>
        /// 导出数据到Excel文件中
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <returns></returns>
        private void OutputToExcel()
        {

            try
            {
                Excel.Application xlApp = new Excel.Application();
                if (xlApp == null)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "无法创建Excel对象!");
                    MessageBox.Show("无法创建Excel对象");
                    return;
                }
                object misValue = System.Reflection.Missing.Value;
                Excel.Workbook wb = xlApp.Workbooks.Add(misValue);

                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                ws.get_Range("A1", "F1").Merge(ws.get_Range("A1", "F1").MergeCells);

                Excel.Range rngHead = (Excel.Range)ws.Cells[1, 1];
                ws.Cells[1, 1] = "韩泰轮胎测宽数据查询";
                ws.Cells[2, 2] = DateTime.Today.ToString("yyyy-MM-dd");
                rngHead.Font.Size = 20;
                rngHead.Font.Name = "宋体";
                rngHead.RowHeight = 50;
                rngHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rngHead.Font.Bold = true;
                rngHead.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                Excel.Range range = (Excel.Range)ws.Columns["A", Type.Missing];
                range.NumberFormatLocal = "@";
                range.ColumnWidth = 26;

                ws.get_Range("B2", "F2").Merge(ws.get_Range("B2", "F2").MergeCells);

                Excel.Range rngTime = ws.get_Range("B2", Type.Missing);
                rngTime.NumberFormatLocal = @"yyyy-MM-dd";

                ws.Cells[2, 1] = "导出时间";
                ws.Cells[3, 1] = "测试时间";
                ws.Cells[3, 2] = "宽度1";
                ws.Cells[3, 3] = "宽度2";
                ws.Cells[3, 4] = "宽度3";
                ws.Cells[3, 5] = "宽度4";
                ws.Cells[3, 6] = "标准宽度";
                ws.Cells[3, 7] = "规格名称";

                //label5.Visible = false;
                //progressBar1.Visible = true;
                //TMProcessbar.Enabled = true;
                for (int i = 3; i < dtDataExcel.Rows.Count; i++)
                {
                    for (int j = 0; j < dtDataExcel.Columns.Count; j++)
                    {
                        ws.Cells[i + 1, j + 1] = dtDataExcel.Rows[i - 3][j].ToString();
                    }
                }

                wb.SaveAs(saveFileName, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                wb.Close(true, misValue, misValue);
                xlApp.Quit();
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(xlApp);
                //label5.Visible = true;
                //progressBar1.Visible = false;
                //TMProcessbar.Enabled = false;
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "查询历史数据成功！\n信息" + dateTimePicker3.Text + "至" + dateTimePicker4.Text + "数据");
                dtDataExcel.Clear();
                

            }
            catch (System.Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "查询历史到Excel失败！\n信息" + ex.Message);
                MessageBox.Show("查询历史到Excel失败！");
                return;
            }

        }

        /// <summary>
        /// 释放控件
        /// </summary>
        /// <param name="obj"></param>
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred" + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }
#endregion

#region 曲线显示
        /// <summary>
        /// 查询历史数据曲线
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                string selectSql = "select w.width1 as w1,w.width2 as w2,w.width3 as w3," +
                     "w.width4 as w4,s.sWidth,w.timestamps as ws from widthdata as w,specificationdata as s where w.timestamps >= '" +
                     dateTimePicker1.Text + "' " + "and w.timestamps <= '" + dateTimePicker2.Text + "' and s.sName ='" + CBHisSpecification.Text + "' and w.specification = s.sName";
                CommonVary.OpenDataConnection();
                using (DataTable dt = new DataTable())
                {

                    using (MySqlDataAdapter da = new MySqlDataAdapter(selectSql, CommonVary.conn))
                    {
                        da.Fill(dt);
                    }
                    if (dt.Rows.Count == 0)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGWARNING, "历史数据记录为空！\n信息" + dateTimePicker1.Text + "至" + dateTimePicker2.Text + "数据为空");
                        MessageBox.Show("选择区段记录为空");
                        return;
                    }
                    else
                    {
                        diagramList1.Clear();
                        diagramList2.Clear();
                        diagramList3.Clear();
                        diagramList4.Clear();
                        diagramListLower.Clear();
                        diagramListUpper.Clear();
                        diagramListStandard.Clear();
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            diagramList1.Add((double)new XDate(DateTime.Parse(dt.Rows[i][5].ToString())), double.Parse(dt.Rows[i][0].ToString()));
                            diagramList2.Add((double)new XDate(DateTime.Parse(dt.Rows[i][5].ToString())), double.Parse(dt.Rows[i][1].ToString()));
                            diagramList3.Add((double)new XDate(DateTime.Parse(dt.Rows[i][5].ToString())), double.Parse(dt.Rows[i][2].ToString()));
                            diagramList4.Add((double)new XDate(DateTime.Parse(dt.Rows[i][5].ToString())), double.Parse(dt.Rows[i][3].ToString()));
                            diagramListStandard.Add((double)new XDate(DateTime.Parse(dt.Rows[i][5].ToString())), double.Parse(button2.Text.Trim()));
                            diagramListUpper.Add((double)new XDate(DateTime.Parse(dt.Rows[i][5].ToString())), double.Parse(button3.Text.Trim()));
                            diagramListLower.Add((double)new XDate(DateTime.Parse(dt.Rows[i][5].ToString())), double.Parse(button4.Text.Trim()));
                        }
                        myPane5.CurveList.Clear();
                        CurveItem myCurve1 = myPane5.AddCurve("宽度1", diagramList1, Color.Red, SymbolType.None);
                        CurveItem myCurve2 = myPane5.AddCurve("宽度2", diagramList2, Color.Gray, SymbolType.None);
                        CurveItem myCurve3 = myPane5.AddCurve("宽度3", diagramList3, Color.Green, SymbolType.None);
                        CurveItem myCurve4 = myPane5.AddCurve("宽度4", diagramList4, Color.Blue, SymbolType.None);
                        CurveItem myCurve5 = myPane5.AddCurve("标准宽度", diagramListStandard, Color.Black, SymbolType.None);
                        CurveItem myCurve6 = myPane5.AddCurve("宽度上限", diagramListUpper, Color.SkyBlue, SymbolType.None);
                        CurveItem myCurve7 = myPane5.AddCurve("宽度下限", diagramListLower, Color.SkyBlue, SymbolType.None);

                        zedGraphControl5.AxisChange();
                        zedGraphControl5.Refresh();
                    }
                }
            }
            catch (Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "历史曲线查询出错！\n信息：" + ex.Message);
                MessageBox.Show("历史曲线查询出错!" + ex.Message);
                return;
            }
        }

        private void RefreshGraph()
        {

            
          
            rediagramList1.Add((double)new XDate(DateTime.Now), (double)receiveMapData.width1.webwidth / 10);
            rediagramList2.Add((double)new XDate(DateTime.Now), (double)receiveMapData.width2.webwidth / 10);
            rediagramList3.Add((double)new XDate(DateTime.Now), (double)receiveMapData.width3.webwidth / 10);
            rediagramList4.Add((double)new XDate(DateTime.Now), (double)receiveMapData.width4.webwidth / 10);
            rediagramListStandard.Add((double)new XDate(DateTime.Now), double.Parse(standardWidth));
            rediagramListUpper.Add((double)new XDate(DateTime.Now), double.Parse(upperErrorWidth));
            rediagramListLower.Add((double)new XDate(DateTime.Now), double.Parse(lowerErrorWidth));
            if (rediagramList1.Count > 200 || rediagramList2.Count > 200 || rediagramList3.Count > 200 || rediagramList4.Count > 200 || rediagramListStandard.Count > 200)
                InitialZedgraph();
           
            
            
            

            

            
            zedGraphControl1.AxisChange();
            zedGraphControl1.Refresh();
            zedGraphControl2.AxisChange();
            zedGraphControl2.Refresh();
            zedGraphControl3.AxisChange();
            zedGraphControl3.Refresh();
            zedGraphControl4.AxisChange();
            zedGraphControl4.Refresh();
        }
#endregion

#region 设置模式选择


        private int SetModelInitail()
        {
            button23.Visible = false;
            button24.Visible = false;
            button19.Visible = false;
            button22.Visible = false;
            displayOrSetFlag = false;
            return CommonVary.RUNNING_OK;
        }
        /// <summary>
        /// 设置模式选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                panel6.Visible = true;
                panel7.Visible = false;
                panel8.Visible = false;
                button23.Visible = false;
                button24.Visible = false;
                sendMapData.cCD_Cal.iCalType = 0;
                if (SetModelInitail() == CommonVary.RUNNING_OK)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS,"相机模式已经初始化！");
                }
            }
            
            
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                panel6.Visible = false;
                panel7.Visible = true;
                panel8.Visible = false;
                button23.Visible = false;
                button24.Visible = false;
                sendMapData.cCD_Cal.iCalType = 1;
                if (SetModelInitail() == CommonVary.RUNNING_OK)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "相机模式已经初始化！");
                }
            }
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true)
            {
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = true;
                button23.Visible = false;
                button24.Visible = false;
                sendMapData.cCD_Cal.iCalType = 2;
                if (SetModelInitail() == CommonVary.RUNNING_OK)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "相机模式已经初始化！");
                }
            }
        }

       

        #endregion

#region 相机设置
        private void button16_MouseDown(object sender, MouseEventArgs e)
        {
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0001;
            //map_data.cCD_Cal.rRefWidth = Int16.Parse(textBox11.Text);
            button23.Visible = true;
            button24.Visible = true;
        }

        private void button16_MouseUp(object sender, MouseEventArgs e)
        {
            Thread.Sleep(500);
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0000;
        }

        private void button23_MouseDown(object sender, MouseEventArgs e)
        {
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0002;
            
        }

        private void button23_MouseUp(object sender, MouseEventArgs e)
        {
            Thread.Sleep(500);
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0000;
            displayOrSetFlag = false;
            button23.Visible = false;
            button24.Visible = false;
        }

        private void button24_MouseDown(object sender, MouseEventArgs e)
        {
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0004;
            
        }

        private void button24_MouseUp(object sender, MouseEventArgs e)
        {
            Thread.Sleep(500);
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0000;
            
            displayOrSetFlag = false;
            button23.Visible = false;
            button24.Visible = false;
        }

        private void button11_MouseDown(object sender, MouseEventArgs e)
        {
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0001;
            
        }

        private void button11_MouseUp(object sender, MouseEventArgs e)
        {
            Thread.Sleep(500);
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0000;
            button23.Visible = true;
            button24.Visible = true;
        }

        private void button10_MouseDown(object sender, MouseEventArgs e)
        {
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0001;
            button23.Visible = true;
            button24.Visible = true;
        }

        private void button10_MouseUp(object sender, MouseEventArgs e)
        {
            Thread.Sleep(500);
            sendMapData.cCD_Cal.controlInt16 = (Int16)0x0000;           
        }
        
        private void textBox11_GotFocus(object sender, EventArgs e)
        {
            if (textBox11.Focused)
            {
                button19.Visible = true;
                button22.Visible = true;
                displayOrSetFlag = true;
            }
        }
        
        private void textBox20_GotFocus(object sender, EventArgs e)
        {
            if (textBox20.Focused)
            {
                button19.Visible = true;
                button22.Visible = true;
                displayOrSetFlag = true;
            }
        }

        private void button18_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (comboBox4.Text == CommonVary.cbx4Item1)
                {
                    sendMapData.wWM_config.iCCD1_No = Int16.Parse(textBox12.Text);
                    sendMapData.wWM_config.iStripNo = Int16.Parse(textBox15.Text);
                }
                else
                {
                    sendMapData.wWM_config.iCCD1_No = Int16.Parse(textBox10.Text);
                    sendMapData.wWM_config.iCCD2_No = Int16.Parse(textBox9.Text);
                    sendMapData.wWM_config.iCCD1_EdgeNo = Int16.Parse(textBox14.Text);
                    sendMapData.wWM_config.iCCD2_EdgeNo = Int16.Parse(textBox13.Text);
                }
                if (comboBox5.Text == CommonVary.cbx5Item2)
                {
                    sendMapData.wWM_config.iDulCCD_refwidth = Int16.Parse(textBox45.Text);
                }
                sendMapData.wWM_config.iSet_bit = 0x0001;

            }
            catch (System.Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "确认WWM配置输入有误！\n信息:" + ex.Message);
                MessageBox.Show("输入有误!");
            }

        }

        private void button18_MouseUp(object sender, MouseEventArgs e)
        {
            Thread.Sleep(500);
            sendMapData.wWM_config.iSet_bit = 0x0000;
            MessageBox.Show("确认成功!");
        }


        private void button19_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                try
                {
                    sendMapData.cCD_Cal.rRefWidth = (Int16)(float.Parse(textBox11.Text) * 10);
                }
                catch (System.ArgumentNullException ex)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "线型标定参考宽度输入为空！");
                    MessageBox.Show("线型标定参考宽度输入为空!");
                    return;
                }
                catch (System.OverflowException ex)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "线型标定参考宽度超出上限！");
                    MessageBox.Show("线型标定参考宽度超出上限!");
                    return;
                }
                catch (System.FormatException ex)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "线型标定参考宽度输入格式出错！");
                    MessageBox.Show("线型标定参考宽度输入格式出错!");
                    return;
                }
                catch (System.Exception ex)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "线型标定参考宽度错误！");
                    MessageBox.Show("线型标定参考宽度输入错误!");
                    return;
                }
            }
            else
                if (radioButton3.Checked)
                {
                    try
                    {
                        sendMapData.cCD_Cal.rCCDdistance = Int16.Parse(textBox20.Text);
                    }
                    catch (System.ArgumentNullException ex)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "全幅标定CCD距离输入为空！");
                        MessageBox.Show("全幅标定CCD距离输入为空!");
                        return;
                    }
                    catch (System.OverflowException ex)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "全幅标定CCD距离超出上限！");
                        MessageBox.Show("全幅标定CCD距离超出上限!");
                        return;
                    }
                    catch (System.FormatException ex)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "全幅标定CCD距离输入格式出错！");
                        MessageBox.Show("全幅标定CCD距离输入格式出错!");
                        return;
                    }
                    catch (System.Exception ex)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "全幅标定CCD距离错误！");
                        MessageBox.Show("全幅标定CCD距离输入错误!");
                        return;
                    }
                }
            //displayOrSetFlag = false;
            numericUpDown1.Focus();
            MessageBox.Show("已确认输入");
            RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGINFO, "已确认输入！");
            button19.Visible = false;
            button22.Visible = false;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            displayOrSetFlag = false;
            textBox11.Text = "";
            textBox20.Text = "";
            MessageBox.Show("已取消输入");
            RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGINFO, "已取消输入！");
            button19.Visible = false;
            button22.Visible = false;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            sendMapData.cCD_Info.iCCDno = (Int16)numericUpDown1.Value;
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

            sendMapData.wWM_config.iWidth_No = (Int16)numericUpDown2.Value;

        }

        

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.Text == CommonVary.cbx5Item2)
            {
                sendMapData.wWM_config.wStrip_CCD = (Int16)(sendMapData.wWM_config.wStrip_CCD | 0x0002);
                label40.Visible = true;
                textBox27.Visible = true;
                textBox45.Visible = true;
            }
            else
            {
                sendMapData.wWM_config.wStrip_CCD = (Int16)(sendMapData.wWM_config.wStrip_CCD & 0x000D);
                label40.Visible = false;
                textBox27.Visible = false;
                textBox45.Visible = false;
            }
        }


        private void button26_Click(object sender, EventArgs e)
        {
            if (dataGridView2SelecteRow == -1)
            {
                MessageBox.Show("没有选中记录");
            }
            else
            {
                if (textBox29.Text == CBSpecification.Text)
                {
                    MessageBox.Show("所选规格正在测试，不允许删除！");
                    return;
                }
                if (MessageBox.Show("是否删除规格：" + dataGridView2.Rows[dataGridView2SelecteRow].Cells[0].Value.ToString(),
                    "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        string delSql = "delete from specificationdata where sName = '" + dataGridView2.Rows[dataGridView2SelecteRow].Cells[0].Value.ToString() + "'";
                        CommonVary.OpenDataConnection();
                        using (MySqlCommand cm = new MySqlCommand(delSql, CommonVary.conn))
                        {
                            cm.ExecuteNonQuery();

                        }
                        initalDataGridView2();
                        MessageBox.Show("规格删成功");
                        CBSpecification.Items.Remove(textBox29.Text);
                        CBHisSpecification.Items.Remove(textBox29.Text);
                        comboBox1.Items.Remove(textBox29.Text);
                    }
                    catch (Exception ex)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "规格删出错");
                        MessageBox.Show("规格删出错");
                    }

                }
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text == CommonVary.cbx4Item1)
            {
                label35.Visible = false;
                label36.Visible = false;
                label37.Visible = false;
                label38.Visible = false;
                textBox23.Visible = false;
                textBox10.Visible = false;
                textBox26.Visible = false;
                textBox14.Visible = false;
                textBox24.Visible = false;
                textBox9.Visible = false;
                textBox25.Visible = false;
                textBox13.Visible = false;

                label33.Visible = true;
                label34.Visible = true;
                textBox22.Visible = true;
                textBox12.Visible = true;
                textBox21.Visible = true;
                textBox15.Visible = true;
                sendMapData.wWM_config.wStrip_CCD = (Int16)(sendMapData.wWM_config.wStrip_CCD & 0xfff0);
            }
            else
            {
                label33.Visible = false;
                label34.Visible = false;
                textBox22.Visible = false;
                textBox12.Visible = false;
                textBox21.Visible = false;
                textBox15.Visible = false;

                label35.Visible = true;
                label36.Visible = true;
                label37.Visible = true;
                label38.Visible = true;
                textBox23.Visible = true;
                textBox10.Visible = true;
                textBox26.Visible = true;
                textBox14.Visible = true;
                textBox24.Visible = true;
                textBox9.Visible = true;
                textBox25.Visible = true;
                textBox13.Visible = true;
                sendMapData.wWM_config.wStrip_CCD = (Int16)(sendMapData.wWM_config.wStrip_CCD | 0x0001);
            }
        }
        #endregion 

#region Timer控制
        private void TMGraph_Tick(object sender, EventArgs e)
        {
            RefreshGraph();
        }

        private void TMDisplay_Tick(object sender, EventArgs e)
        {
            RefreshDisplay();
        }

        /// <summary>
        /// 进度条变化定时器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TMProcessbar_Tick(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 2)
            {
                if (progressBar1.Value >= 100)
                {
                    progressBar1.Value = 0;
                }
                progressBar1.PerformStep();
                Application.DoEvents();
                try
                {
                    if (ToExcelThread.IsAlive == false)
                    {
                        label5.Text = "没有查询任务";
                        toGridViewSucc = CommonVary.RUNNNING_WRONG;
                        progressBar1.Visible = false;
                        TMProcessbar.Enabled = false;
                        if (comboBox1.Text != "ALL")
                        {
                            MessageBox.Show("查询成功");
                        }
                        else if (MessageBox.Show("是否清空数据库！", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            ClearHistryData();
                            return;
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                catch(Exception ex)
                {

                }
            }
            
        }

        

        /// <summary>
        /// TCP数据采集定时器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TMTCP_Tick(object sender, EventArgs e)
        {
            try
            {
                if (tcpClient.RecevedData() == CommonVary.RUNNNING_WRONG)
                {
                    TMTCP.Enabled = false;
                    TMDisplay.Enabled = false;
                    TMGraph.Enabled = false;
                    TMSaveData.Enabled = false;

                    button21.Text = "开始测试";
                }
                sendMapData.dataExchange_Header.wCounterToBST++;
            }
            catch (System.Net.Sockets.SocketException ex)
            {
                //TMTCP.Enabled = false;
                //TMDisplay.Enabled = false;
                //TMGraph.Enabled = false;
                //TMSaveData.Enabled = false;
                //button21.Text = "开始测试";
                tcpClient.newclient.Close();
                tcpClient.ConfigTcpClient();
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCPClient连接Server_PLC异常！\n信息：" + ex.Message);

            }
            catch (System.ArgumentNullException ex)
            {
                TMTCP.Enabled = false;
                TMDisplay.Enabled = false;
                TMGraph.Enabled = false;
                TMSaveData.Enabled = false;
                button21.Text = "开始测试";
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "发送PLC时数据为空！\n信息：" + ex.Message);
                MessageBox.Show("发送PLC时数据为空，请检查日志！");


            }
            catch (System.InvalidOperationException ex)
            {
                //TMTCP.Enabled = false;
                //TMDisplay.Enabled = false;
                //TMGraph.Enabled = false;
                //TMSaveData.Enabled = false;
                //button21.Text = "开始测试";
                tcpClient.newclient.Close();
                tcpClient.ConfigTcpClient();
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCPClient连接Server_PLC已断开！\n信息：" + ex.Message);
                //MessageBox.Show("连接已断开！");

            }
            catch (System.Exception ex)
            {
                TMTCP.Enabled = false;
                TMDisplay.Enabled = false;
                TMGraph.Enabled = false;
                TMSaveData.Enabled = false;
                button21.Text = "开始测试";
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "TCP数据传输出错！\n信息：" + ex.Message);
                MessageBox.Show("TCP数据传输出错，请检查日志！");

            }
        }

        /// <summary>
        /// 时间调整
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        DateTime dtNow,dtLastMonth;
        private void TMTimeChange_Tick(object sender, EventArgs e)
        {
            dtNow = DateTime.Now;
            LBDate.Text = dtNow.Date.ToLongDateString();
            LBHisDate.Text = dtNow.Date.ToLongDateString();
            LBTime.Text = dtNow.ToLongTimeString();
            LBHisTime.Text = dtNow.ToLongTimeString();
            dtLastMonth = Convert.ToDateTime(dtNow.Date.ToLongDateString() + " 01:00:00");
            //Console.WriteLine("1:" + d.ToString());
            //Console.WriteLine("2:" + dt.ToString());
            if (dtLastMonth.ToString() == dtNow.ToString())
            {
                int startMonth;
                int startYear;
                int startDay;
                startDay = dtNow.Day;
                if(dtNow.Month == 1)
                {
                    startMonth = 11;
                    startYear = dtNow.Year -1;
                }
                else if(dtNow.Month == 2)
                {
                    startMonth = 12;
                    startYear = dtNow.Year - 1;
                }
                else 
                {
                    startMonth = dtNow.Month - 2;
                    startYear = dtNow.Year;
                }
                DeleteDataEveryDay(Convert.ToDateTime(startYear.ToString()+ "-" +startMonth.ToString()+ "-"+ startDay.ToString() + " 00:00:00"));
                
            }
            if (MyPort.IsOpen)
                label68.Text = "COM口状态：" + CommonVary.myCOMPort +"开";
            else
                label68.Text = "COM口状态：" + CommonVary.myCOMPort + "关";
            if (startSystem == true )
            {
                if (saveFlag == true)
                {
                    AutoUpdateSaveData(saveFlag, "正在保存数据！！！\n(频率：" + CommonVary.SaveDataInterval / 1000 + "秒每次)");
                }
                else
                {
                    //AutoUpdateSaveData(saveFlag, "所测数据没有保存！！！\n机器停止运行");
                    AutoUpdateSaveData(true, "正在保存数据！！！\n(频率：" + CommonVary.SaveDataInterval / 1000 + "秒每次)\n机器停止运行");
                }
            }
          
        }

        private void TMSaveData_Tick(object sender, EventArgs e)
        {
            SaveData();
        }
        #endregion

#region 其它用
        /// <summary>
        /// 界面刷新函数
        /// </summary>
        private void RefreshDisplay()
        {
            if (tabControl1.SelectedIndex == 3)
            {

                LBedge1.Text = "边1:" + receiveMapData.cCD_Info.edge1.ToString() + "像素";
                LBedge2.Text = "边2:" + receiveMapData.cCD_Info.edge2.ToString() + "像素";
                LBedge3.Text = "边3:" + receiveMapData.cCD_Info.edge3.ToString() + "像素";
                LBedge4.Text = "边4:" + receiveMapData.cCD_Info.edge4.ToString() + "像素";

                label22.Text = "物料宽度：" + (((float)receiveMapData.cCD_Info.webwidth)/10).ToString();
                label21.Text = "焦距：" + receiveMapData.cCD_Info.focus.ToString();
                label20.Text = "曝光时间" + receiveMapData.cCD_Info.exposuretime.ToString();
                label19.Text = "分辨率：" + (((float)receiveMapData.cCD_Info.resolution)/10).ToString();
                label18.Text = "零点像素：" + receiveMapData.cCD_Info.zero_pix.ToString();
                label43.Text = "CAN地址：" + receiveMapData.cCD_Info.CAN_Address.ToString();

                if (radioButton1.Checked && !textBox11.Focused && !displayOrSetFlag)
                {
                    textBox11.Text = (((float)receiveMapData.cCD_Cal.rRefWidth)/10).ToString();
                }
                if (radioButton2.Checked)
                {
                    richTextBox1.Text = wStatusDisplay[receiveMapData.cCD_Cal.iText];
                    
                    label23.Text = receiveMapData.cCD_Cal.iProgress.ToString();
                }
                if (radioButton3.Checked)
                {
                    if (!textBox20.Focused && !displayOrSetFlag)
                    {
                        textBox20.Text = (((float)receiveMapData.cCD_Cal.rCCDdistance)/10).ToString();
                    }
                    richTextBox2.Text = wStatusDisplay[receiveMapData.cCD_Cal.iText];
                    label26.Text = receiveMapData.cCD_Cal.iProgress.ToString();
                }
                if (comboBox4.Text == CommonVary.cbx4Item1)
                {
                    textBox22.Text = receiveMapData.wWM_config.iCCD1_No.ToString();
                    textBox21.Text = receiveMapData.wWM_config.iStripNo.ToString();
                }
                else
                    if (comboBox4.Text == CommonVary.cbx4Item2)
                    {
                        textBox23.Text = receiveMapData.wWM_config.iCCD1_No.ToString();
                        textBox24.Text = receiveMapData.wWM_config.iCCD2_No.ToString();
                        textBox26.Text = receiveMapData.wWM_config.iCCD1_EdgeNo.ToString();
                        textBox25.Text = receiveMapData.wWM_config.iCCD2_EdgeNo.ToString();
                    }
                textBox27.Text = ((float)receiveMapData.wWM_config.iDulCCD_refwidth / 10).ToString();
                textBox56.Text = ((float)receiveMapData.sys_Inform.width1_lower_limit_analog_output / 10).ToString();
                textBox55.Text = ((float)receiveMapData.sys_Inform.width1_upper_limit_analog_output / 10).ToString();
                textBox54.Text = ((float)receiveMapData.sys_Inform.width2_lower_limit_analog_output / 10).ToString();
                textBox53.Text = ((float)receiveMapData.sys_Inform.width2_upper_limit_analog_output / 10).ToString();
                if (receiveMapData.cCD_Info.state == 0)
                {
                    button9.BackColor = Color.Red;
                }
                else
                    button9.BackColor = Color.Green;
            }
            else
                if (tabControl1.SelectedIndex == 0)
                {

                    textBox2.Text = (((float)receiveMapData.width1.webwidth)/10).ToString();
                    textBox3.Text = (((float)receiveMapData.width2.webwidth)/10).ToString();
                    textBox5.Text = (((float)receiveMapData.width3.webwidth)/10).ToString();
                    textBox7.Text = (((float)receiveMapData.width4.webwidth)/10).ToString();
                    if (float.Parse(textBox2.Text) > float.Parse(upperErrorWidth) || float.Parse(textBox2.Text) < float.Parse(lowerErrorWidth))
                        textBox2.BackColor = Color.Red;
                    else if (float.Parse(textBox2.Text) > float.Parse(upperWarWidth) || float.Parse(textBox2.Text) < float.Parse(lowerWarWidth))
                        textBox2.BackColor = Color.Yellow;
                    else
                        textBox2.BackColor = Color.Silver;
                    if (float.Parse(textBox3.Text) > float.Parse(upperErrorWidth) || float.Parse(textBox3.Text) < float.Parse(lowerErrorWidth))
                        textBox3.BackColor = Color.Red;
                    else if (float.Parse(textBox3.Text) > float.Parse(upperWarWidth) || float.Parse(textBox3.Text) < float.Parse(lowerWarWidth))
                        textBox3.BackColor = Color.Yellow;
                    else
                        textBox3.BackColor = Color.Silver;

                    if (float.Parse(textBox5.Text) > float.Parse(upperErrorWidth) || float.Parse(textBox5.Text) < float.Parse(lowerErrorWidth))
                        textBox5.BackColor = Color.Red;
                    else if (float.Parse(textBox5.Text) > float.Parse(upperWarWidth) || float.Parse(textBox5.Text) < float.Parse(lowerWarWidth))
                        textBox5.BackColor = Color.Yellow;
                    else
                        textBox5.BackColor = Color.Silver;

                    if (float.Parse(textBox7.Text) > float.Parse(upperErrorWidth) || float.Parse(textBox7.Text) < float.Parse(lowerErrorWidth))
                        textBox7.BackColor = Color.Red;
                    else if (float.Parse(textBox7.Text) > float.Parse(upperWarWidth) || float.Parse(textBox7.Text) < float.Parse(lowerWarWidth))
                        textBox7.BackColor = Color.Yellow;
                    else
                        textBox7.BackColor = Color.Silver;
                    textBox31.Text = (((float)receiveMapData.width1.Splice1)/10).ToString();
                    textBox32.Text = (((float)receiveMapData.width1.Splice2)/10).ToString();

                    textBox34.Text = (((float)receiveMapData.width2.Splice1)/10).ToString();
                    textBox33.Text = (((float)receiveMapData.width2.Splice2)/10).ToString();

                    textBox36.Text = (((float)receiveMapData.width3.Splice1)/10).ToString();
                    textBox35.Text = (((float)receiveMapData.width3.Splice2)/10).ToString();

                    textBox38.Text = (((float)receiveMapData.width4.Splice1)/10).ToString();
                    textBox37.Text = (((float)receiveMapData.width4.Splice2)/10).ToString();
                }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 3)
            {
                sendMapData.cCD_Info.iCCDno = (Int16)numericUpDown1.Value;
                sendMapData.wWM_config.iWidth_No = (Int16)numericUpDown2.Value;
                TMDisplay.Enabled = true;
            }
            else
            {
                TMDisplay.Enabled = false;
                if (tabControl1.SelectedIndex == 2)
                {
                    GetStandardHisWidth(CBHisSpecification.Text);
                }
                else 
                {
                    if (tabControl1.SelectedIndex == 0)
                    {
                        TMDisplay.Enabled = true;
                    }
                    else
                    {
                        TMDisplay.Enabled = false;
                        if (tabControl1.SelectedIndex == 4)
                        {
                            if (dataGridView2.DataSource == null)
                                initalDataGridView2();
                            InitialSetInfo();
                            //textBox30.Text = CommonVary.PLCAddress;
                            //textBox39.Text = CommonVary.PLCPort.ToString();
                            //textBox43.Text = CommonVary.PCFromPLCPort.ToString();
                            //textBox44.Text = (CommonVary.SaveDataInterval/1000).ToString();
                        }
                    }
                }
               
                    
            }

        }


        


        private void CBSpecification_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetStandardWidth(CBSpecification.Text);
            SendDataToCOM(sName);
            InitialZedgraph();
        }
        /// <summary>
        /// 获取指定规格的宽度及上限和下限
        /// </summary>
        /// <param name="specification"></param>
        /// <returns></returns>
        private int GetStandardWidth(string specification)
        {
            try
            {
                string sWidthString = "select sName,sWidth,sWidthUpper,sWidthLower,sWidthWarUpper,sWidthWarLower,sAngle from specificationdata where sName = '" + specification + "'";
                using (DataTable dt = new DataTable())
                {
                    using (MySqlDataAdapter myDataAdapter = new MySqlDataAdapter(sWidthString, CommonVary.conn))
                    {
                        myDataAdapter.Fill(dt);
                    }
                    if (dt.Columns.Count > 1)
                    {
                        sName = dt.Rows[0]["sName"].ToString().Trim();
                        standardWidth = dt.Rows[0]["sWidth"].ToString();
                        upperErrorWidth = dt.Rows[0]["sWidthUpper"].ToString();
                        lowerErrorWidth = dt.Rows[0]["sWidthLower"].ToString();
                        upperWarWidth = dt.Rows[0]["sWidthWarUpper"].ToString();
                        lowerWarWidth = dt.Rows[0]["sWidthWarLower"].ToString();
                        sAngle = dt.Rows[0]["sAngle"].ToString();

                        sendMapData.wWM_Parameter1.SetWebwidth = (Int16)((float.Parse(standardWidth)) * 10f);
                        sendMapData.wWM_Parameter2.SetWebwidth = sendMapData.wWM_Parameter1.SetWebwidth;
                        sendMapData.wWM_Parameter3.SetWebwidth = sendMapData.wWM_Parameter1.SetWebwidth;
                        sendMapData.wWM_Parameter4.SetWebwidth = sendMapData.wWM_Parameter1.SetWebwidth;
                        sendMapData.wWM_Parameter5.SetWebwidth = sendMapData.wWM_Parameter1.SetWebwidth;
                        sendMapData.wWM_Parameter6.SetWebwidth = sendMapData.wWM_Parameter1.SetWebwidth;

                        sendMapData.wWM_Parameter1.SetErrTol_upper = (Int16)((float.Parse(upperErrorWidth) - float.Parse(standardWidth)) * 10f);
                        sendMapData.wWM_Parameter2.SetErrTol_upper = sendMapData.wWM_Parameter1.SetErrTol_upper;
                        sendMapData.wWM_Parameter3.SetErrTol_upper = sendMapData.wWM_Parameter1.SetErrTol_upper;
                        sendMapData.wWM_Parameter4.SetErrTol_upper = sendMapData.wWM_Parameter1.SetErrTol_upper;
                        sendMapData.wWM_Parameter5.SetErrTol_upper = sendMapData.wWM_Parameter1.SetErrTol_upper;
                        sendMapData.wWM_Parameter6.SetErrTol_upper = sendMapData.wWM_Parameter1.SetErrTol_upper;

                        sendMapData.wWM_Parameter1.SetErrTol_lower = (Int16)((float.Parse(standardWidth) - float.Parse(lowerErrorWidth)) * 10f);
                        sendMapData.wWM_Parameter2.SetErrTol_lower = sendMapData.wWM_Parameter1.SetErrTol_lower;
                        sendMapData.wWM_Parameter3.SetErrTol_lower = sendMapData.wWM_Parameter1.SetErrTol_lower;
                        sendMapData.wWM_Parameter4.SetErrTol_lower = sendMapData.wWM_Parameter1.SetErrTol_lower;
                        sendMapData.wWM_Parameter5.SetErrTol_lower = sendMapData.wWM_Parameter1.SetErrTol_lower;
                        sendMapData.wWM_Parameter6.SetErrTol_lower = sendMapData.wWM_Parameter1.SetErrTol_lower;

                        sendMapData.wWM_Parameter1.SetWarTol_upper = (Int16)((float.Parse(upperWarWidth) - float.Parse(standardWidth)) * 10f);
                        sendMapData.wWM_Parameter2.SetWarTol_upper = sendMapData.wWM_Parameter1.SetWarTol_upper;
                        sendMapData.wWM_Parameter3.SetWarTol_upper = sendMapData.wWM_Parameter1.SetWarTol_upper;
                        sendMapData.wWM_Parameter4.SetWarTol_upper = sendMapData.wWM_Parameter1.SetWarTol_upper;
                        sendMapData.wWM_Parameter5.SetWarTol_upper = sendMapData.wWM_Parameter1.SetWarTol_upper;
                        sendMapData.wWM_Parameter5.SetWarTol_upper = sendMapData.wWM_Parameter1.SetWarTol_upper;

                        sendMapData.wWM_Parameter1.SetWarTol_lower = (Int16)((float.Parse(standardWidth) - float.Parse(lowerWarWidth)) * 10f);
                        sendMapData.wWM_Parameter2.SetWarTol_lower = sendMapData.wWM_Parameter1.SetWarTol_lower;
                        sendMapData.wWM_Parameter3.SetWarTol_lower = sendMapData.wWM_Parameter1.SetWarTol_lower;
                        sendMapData.wWM_Parameter4.SetWarTol_lower = sendMapData.wWM_Parameter1.SetWarTol_lower;
                        sendMapData.wWM_Parameter5.SetWarTol_lower = sendMapData.wWM_Parameter1.SetWarTol_lower;
                        sendMapData.wWM_Parameter6.SetWarTol_lower = sendMapData.wWM_Parameter1.SetWarTol_lower;
                    }
                }
                textBox1.Text = standardWidth;
                textBox4.Text = standardWidth;
                textBox6.Text = standardWidth;
                textBox8.Text = standardWidth;
                label63.Text = "当前宽度：" + standardWidth;
                label69.Text = "当前角度：" + sAngle;

                RefreshDisplay();
                return CommonVary.RUNNING_OK;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("规格宽度读取错误");
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "规格宽度读取错误!\n信息" + ex.Message);
                return CommonVary.RUNNNING_WRONG;
            }
        }

        private int GetStandardHisWidth(string specification)
        {
            try
            {
                string sWidthString = "select sWidth,sWidthUpper,sWidthLower from specificationdata where sName = '" + specification + "'";
                using (DataTable dt = new DataTable())
                {
                    using (MySqlDataAdapter myDataAdapter = new MySqlDataAdapter(sWidthString, CommonVary.conn))
                    {
                        myDataAdapter.Fill(dt);
                    }
                    if (dt.Columns.Count > 1)
                    {
                        standardHisWidth = dt.Rows[0]["sWidth"].ToString();
                        upperHisWidth = dt.Rows[0]["sWidthUpper"].ToString();
                        LowerHisWidth = dt.Rows[0]["sWidthLower"].ToString();
                    }
                }
                button2.Text = standardHisWidth;
                button3.Text = upperHisWidth;
                button4.Text = LowerHisWidth;

                return CommonVary.RUNNING_OK;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("历史规格宽度读取错误");
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "历史规格宽度读取错误!\n信息" + ex.Message);
                return CommonVary.RUNNNING_WRONG;
            }

        }

        private void CBHisSpecification_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetStandardHisWidth(CBHisSpecification.Text);
        }

        /// <summary>
        /// 开始测试
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button21_Click(object sender, EventArgs e)
        {
            if (button21.Text == "开始测试")
            {
                //InitalizeTcp();
               // if (MessageBox.Show("是否保存数据", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                try
                {
                    if (!MyPort.IsOpen)
                        MyPort.Open();
                }
                catch (Exception ex)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGWARNING, CommonVary.myCOMPort + "被占用！");
                    MessageBox.Show(CommonVary.myCOMPort + "被占用！");
                    return;
                }
                startSystem = true;
                if(saveFlag)
                {
                    AutoUpdateSaveData(true, "正在保存数据！！！\n(频率：" + CommonVary.SaveDataInterval / 1000 + "秒每次)");
                }
                else
                {
                    AutoUpdateSaveData(false, "所测数据没有保存！！！\n机器停止运行！！！");
                }
                TMTCP.Interval = 50;
                TMTCP.Enabled = true;
                TMDisplay.Enabled = true;
                
                rediagramList1.Clear();
                rediagramList2.Clear();
                rediagramList3.Clear();
                rediagramList4.Clear();
                rediagramListLower.Clear();
                rediagramListUpper.Clear();
                rediagramListStandard.Clear();

                TMGraph.Interval = 500;
                TMGraph.Enabled = true;
                button21.Text = "中止测试";


            }
            else
            {
                try
                {
                    if (MyPort.IsOpen)
                        MyPort.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("串口关闭失败");
                }
                startSystem = false;
                startMach = false;
                saveFlag = false;
                TMTCP.Enabled = false;
                TMDisplay.Enabled = false;
                TMGraph.Enabled = false;
                AutoUpdateSaveData(false, "");
                button21.Text = "开始测试";
                label11.Visible = false;
            }
        }

        private int SaveData()
        {
            string saveSql = "insert into widthdata(width1,width2,width3,width4,specification,timestamps) values('" + (((float)receiveMapData.width1.webwidth)/10).ToString() +
                "','" + (((float)receiveMapData.width2.webwidth) / 10).ToString() + "','" + (((float)receiveMapData.width3.webwidth) / 10).ToString() +
                "','" + (((float)receiveMapData.width4.webwidth) / 10).ToString() + "','" + CBSpecification.Text +
                "','" + DateTime.Now + "')";
            try
            {
                CommonVary.OpenDataConnection();
                mySqlCommand = new MySqlCommand(saveSql, CommonVary.conn);
                mySqlCommand.ExecuteNonQuery();
                return CommonVary.RUNNING_OK;
            }
            catch (System.Exception ex)
            {
                TMTCP.Enabled = false;
                TMDisplay.Enabled = false;
                TMGraph.Enabled = false;
                TMSaveData.Enabled = false;
                button21.Text = "开始测试";
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "保存数据出错!\n信息：" + ex.Message);
                MessageBox.Show("保存数据出错");
                return CommonVary.RUNNNING_WRONG;
            }
        }

        

#endregion

#region 测试用
        private void splitContainer1_Panel2_MouseMove(object sender, MouseEventArgs e)
        {
            //Console.WriteLine(e.X + ","+e.Y);
        }
        /// <summary>
        /// 产生模拟数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button20_Click(object sender, EventArgs e)
        {
            ClearHistryData();
            //Random rd = new Random();
            //float width1, width2, width3, width4, specification;
            //string[] s = new string[] { "A1", "A2", "A3" };
            //DateTime timestamps = DateTime.Now;


            //string minsecond;

            //MySqlCommand myCommand;
            //for (int i = 0; i < 500; i++)
            //{
            //    width1 = 20 + (float)rd.NextDouble() * 5;
            //    width2 = 20 + (float)rd.NextDouble() * 5;
            //    width3 = 20 + (float)rd.NextDouble() * 5;
            //    width4 = 20 + (float)rd.NextDouble() * 5;
            //    timestamps = timestamps.AddSeconds(10);
            //    minsecond = DateTime.Now.Millisecond.ToString();
            //    string tSql = "insert into widthdata(specification,width1,width2,width3,width4,timestamps,minsecond)value(" +
            //            "'" + s[rd.Next(0, 2)] + "','" + width1.ToString() + "','" + width2.ToString() + "','" + width3.ToString() +
            //            "','" + width4.ToString() + "','" + timestamps + "','" + minsecond + "')";
            //    myCommand = new MySqlCommand(tSql, CommonVary.conn);
            //    myCommand.ExecuteNonQuery();
            //}
        }
#endregion
#region 规格参数设置
        private int initalDataGridView2()
        {
            dataGridView2.DataSource = null;
            try
            {
                CommonVary.OpenDataConnection();
                string insertString = "select sName,sWidth,sWidthUpper,sWidthWarUpper,sWidthLower,sWidthWarLower,sAngle,createtime from specificationdata order by sName";
                using (DataTable ds = new DataTable())
                {
                    using (MySqlDataAdapter da = new MySqlDataAdapter(insertString, CommonVary.conn))
                    {
                        da.Fill(ds);
                    }
                    ds.Columns[0].ColumnName = "规格";
                    ds.Columns[1].ColumnName = "标准宽度";
                    ds.Columns[2].ColumnName = "出错上限";
                    ds.Columns[3].ColumnName = "报警上限";
                    ds.Columns[4].ColumnName = "出错下限";
                    ds.Columns[5].ColumnName = "报警下限";
                    ds.Columns[6].ColumnName = "裁剪角度";
                    ds.Columns[7].ColumnName = "修改时间";


                    dataGridView2.DataSource = ds;
                    dataGridView2.ClearSelection();


                    dataGridView2.DefaultCellStyle.Font = new Font("微软雅黑", 8);
                    dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("微软雅黑", 8);
                    dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                    for (int i = 0; i < dataGridView2.Columns.Count; i++)
                    {
                        dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                    }

                    dataGridView2SelecteRow = -1;
                }
            }
            catch (Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "规格参数设置读取失败!\n信息：" + ex.Message);
                MessageBox.Show("规格参数设置读取失败");
                return CommonVary.RUNNNING_WRONG;
            }
            return CommonVary.RUNNING_OK;
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dataGridView2SelecteRow = e.RowIndex;
                textBox29.Text = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                textBox16.Text = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox17.Text = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                textBox18.Text = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                textBox19.Text = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
                textBox28.Text = dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString();
                textBox46.Text = dataGridView2.Rows[e.RowIndex].Cells[6].Value.ToString();
                button17.Enabled = true;
                button26.Enabled = true;
                button27.Enabled = true;
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            textBox29.Text = "";
            textBox16.Text = "";
            textBox17.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
            textBox28.Text = "";
            textBox46.Text = "";
            button17.Enabled = false;
            button26.Enabled = false;
            button27.Enabled = false;
            initalDataGridView2();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (CheckTextBox() == CommonVary.RUNNING_OK&&CheckSame()==CommonVary.RUNNING_OK)
            {
                if (textBox29.Text != dataGridView2.Rows[dataGridView2SelecteRow].Cells[0].Value.ToString())
                {
                    MessageBox.Show("规格名称不可以改");
                    return;
                }
                else
                {
                    try
                    {
                        string updateString = "update specificationdata set sWidth ='" + textBox16.Text +
                        "',sWidthUpper ='" + textBox17.Text + "',sWidthWarUpper ='" + textBox18.Text +
                        "',sWidthLower ='" + textBox19.Text + "',sWidthWarLower ='" + textBox28.Text +
                        "',sAngle ='" + textBox46.Text.Trim() + "' where sName = '" + textBox29.Text + "'";
                        CommonVary.OpenDataConnection();
                        using (MySqlCommand cm = new MySqlCommand(updateString, CommonVary.conn))
                        {
                            cm.ExecuteNonQuery();
                        }
                        MessageBox.Show("修改成功！");
                        textBox29.Text = "";
                        textBox16.Text = "";
                        textBox17.Text = "";
                        textBox18.Text = "";
                        textBox19.Text = "";
                        textBox28.Text = "";
                        textBox46.Text = "";
                        button17.Enabled = false;
                        button26.Enabled = false;
                        button27.Enabled = false;
                        initalDataGridView2();
                        GetStandardWidth(CBSpecification.Text.Trim());
                        GetStandardHisWidth(CBHisSpecification.Text.Trim());
                        InitialZedgraph();
                    }
                    catch (System.Exception ex)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "规格修改出错!\n信息" + ex.Message);
                        MessageBox.Show("规格修改插入出错！");
                    }

                }

            }
        }

        private int CheckSame()
        {
            if (textBox16.Text == dataGridView2.Rows[dataGridView2SelecteRow].Cells[1].Value.ToString()&&
               textBox17.Text == dataGridView2.Rows[dataGridView2SelecteRow].Cells[2].Value.ToString() &&
               textBox18.Text == dataGridView2.Rows[dataGridView2SelecteRow].Cells[3].Value.ToString() &&
               textBox19.Text == dataGridView2.Rows[dataGridView2SelecteRow].Cells[4].Value.ToString() &&
               textBox28.Text == dataGridView2.Rows[dataGridView2SelecteRow].Cells[5].Value.ToString() &&
               textBox29.Text == dataGridView2.Rows[dataGridView2SelecteRow].Cells[0].Value.ToString() &&
               textBox46.Text == dataGridView2.Rows[dataGridView2SelecteRow].Cells[6].Value.ToString())
             
            {
                MessageBox.Show("没有改变规格信息");
                return CommonVary.RUNNNING_WRONG;
            }


            return CommonVary.RUNNING_OK;

        }

        private int CheckTextBox()
        {
            if (textBox16.Text == "" ||
                textBox17.Text == "" ||
                textBox18.Text == "" ||
                textBox19.Text == "" ||
                textBox28.Text == "" ||
                textBox46.Text == "" ||
                textBox29.Text == "")
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "输入不可以为空!");
                MessageBox.Show("输入不可以为空！");
                return CommonVary.RUNNNING_WRONG;
            }
            
                try
                {
                    float.Parse(textBox16.Text);
                    float.Parse(textBox17.Text);
                    float.Parse(textBox18.Text);
                    float.Parse(textBox19.Text);
                    float.Parse(textBox46.Text);
                    float.Parse(textBox28.Text);
                }
                catch (System.Exception ex)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "宽度只能为数字!");
                    MessageBox.Show("宽度只能为数字！");
                    return CommonVary.RUNNNING_WRONG;
                }
            

            return CommonVary.RUNNING_OK;
        }


        private int CheckSpecification(string name)
        {
            string checkSql = "select * from specificationdata where sName = '" + textBox29.Text + "'";
            CommonVary.OpenDataConnection();
            using (DataTable dt = new DataTable())
            {
                using (MySqlDataAdapter da = new MySqlDataAdapter(checkSql, CommonVary.conn))
                {
                    da.Fill(dt);
                }
                if (dt.Rows.Count != 0)
                {
                    return CommonVary.RUNNNING_WRONG;
                }
                else
                    return CommonVary.RUNNING_OK;
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            if (CheckTextBox() == CommonVary.RUNNING_OK)
            {
                if (CheckSpecification(textBox29.Text)==CommonVary.RUNNNING_WRONG)
                {
                    MessageBox.Show("规格已经存在");
                    return;
                }
                else
                {
                    try
                    {
                        string updateString = "insert into specificationdata (sWidth,sWidthUpper,sWidthWarUpper,sWidthLower,sWidthWarLower,sName,sAngle) values ('" + textBox16.Text +
                        "','" + textBox17.Text + "','" + textBox18.Text +"','" + textBox19.Text + "','" + textBox28.Text +
                        "','" + textBox29.Text + "','" + textBox46.Text +"')";
                        CommonVary.OpenDataConnection();
                        using (MySqlCommand cm = new MySqlCommand(updateString, CommonVary.conn))
                        {
                            cm.ExecuteNonQuery();
                        }
                        MessageBox.Show("新增规格成功！");
                        CBSpecification.Items.Add(textBox29.Text);
                        CBHisSpecification.Items.Add(textBox29.Text);
                        comboBox1.Items.Add(textBox29.Text);
                        textBox29.Text = "";
                        textBox16.Text = "";
                        textBox17.Text = "";
                        textBox18.Text = "";
                        textBox19.Text = "";
                        textBox28.Text = "";
                        textBox46.Text = "";
                        button17.Enabled = false;
                        button26.Enabled = false;
                        button27.Enabled = false;
                        initalDataGridView2();
                    }
                    catch (System.Exception ex)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "数据插入出错!\n信息" + ex.Message);
                        MessageBox.Show("数据插入出错！");
                    }

                }

            }
        }

#endregion


        #region 系统设置
        private void button12_Click(object sender, EventArgs e)
        {
            if (button21.Text =="中止测试")
                MessageBox.Show("请先中止测试");
            else
            {
                CommonVary.WriteConfig("PLCAddress", CommonVary.PLCAddress, textBox30.Text.Trim());
                InitialSetInfo();
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            //CommonVary.ServerAddress = textBox41.Text;
            //CommonVary.ServerPort = int.Parse(textBox40.Text);
            //MessageBox.Show("修改成功");
            if (button21.Text == "中止测试")
                MessageBox.Show("请先中止测试");
            else
            {
                CommonVary.WriteConfig("PCFromPLCPort", CommonVary.PCFromPLCPort.ToString(), textBox43.Text.Trim());
                InitialSetInfo();
            }

        }
        private void button5_Click(object sender, EventArgs e)
        {
            if (button21.Text == "中止测试")
                MessageBox.Show("请先中止测试");
            else
            {
                CommonVary.WriteConfig("SaveDataInterval", (CommonVary.SaveDataInterval / 1000).ToString(), (int.Parse(textBox44.Text.Trim()).ToString()));
                InitialSetInfo();
            }
        }
        private void button14_Click(object sender, EventArgs e)
        {
            if (button21.Text == "中止测试")
                MessageBox.Show("请先中止测试");
            else
            {
                CommonVary.WriteConfig("PLCPort", CommonVary.PLCPort.ToString(), textBox39.Text.Trim());
                InitialSetInfo();
            }
            //CommonVary.PCFromPLCPort = int.Parse(textBox43.Text);
            //CommonVary.PCFromServerPort = int.Parse(textBox2.Text);
            //MessageBox.Show("修改成功");
        }
        #endregion


        #region 串口通讯

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox7.Text.Trim() == CommonVary.myCOMPort)
                return;
            if (button21.Text == "中止测试")
            {
                MessageBox.Show("请先中止测试");
                comboBox7.Text = CommonVary.myCOMPort;
            }
            else
            {


                if (CommonVary.WriteConfig("WriteCOMPort", CommonVary.myCOMPort.ToString(), comboBox7.Text.Trim()) == CommonVary.RUNNING_OK)
                {
                    InitialSetInfo();
                    InitialWriteCOM();
                }
                
            }
        }

       

        private int InitialWriteCOM()
        {
            try
            {
                if (MyPort.IsOpen)
                {
                    MyPort.Close();
                }
                MyPort.PortName = CommonVary.myCOMPort;
                MyPort.Open();
                return CommonVary.RUNNING_OK;
            }
            catch (Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "初始化写COM口错误被占用" + Environment.NewLine + "信息：" + ex.Message);
                MessageBox.Show("初始化写COM口错误被占用" + ex.Message);
                return CommonVary.RUNNNING_WRONG;

            }
        }
    
        /// <summary>
        /// 发送指定规格数据到COM口
        /// </summary>
        /// <param name="name"></param>
        private void SendDataToCOM(string name)
        {
            if (startSystem == true && MyPort.IsOpen)
            {
                string sendName = name;
                string sendWidth = standardWidth;
                string sendAngle = sAngle;
                //string stSQL = "select sName,sWidth,sAngle from specificationdata where sName = '" + name + "'";
                //try
                //{
                //    CommonVary.OpenDataConnection();
                //    MySqlDataAdapter da = new MySqlDataAdapter(stSQL, CommonVary.conn);
                //    DataTable dt = new DataTable();
                //    da.Fill(dt);
                //    if (dt.Rows.Count != 1)
                //    {
                //        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGDEBUG, "数据库中没有所选规格");
                //        MessageBox.Show("数据库中没有所选规格");
                //        return;
                //    }
                //    else
                //    {
                //        sName = dt.Rows[0]["sName"].ToString();
                //        sWidth = dt.Rows[0]["sWidth"].ToString();
                //        sAngle = dt.Rows[0]["sAngle"].ToString();
                //    }
                    if (ushort.Parse(sendWidth.Trim()) < 100)
                        sendWidth = "0" + (int.Parse(sendWidth) * 10).ToString();
                    else
                        sendWidth = (int.Parse(sendWidth) * 10).ToString();
                    //sendWidth = "0" + sendWidth;

                    string stringInfoSend = "$"+ sendName.Trim() + sendWidth.Trim() + sendAngle.Trim();
                    char[] charInfoSend = stringInfoSend.ToCharArray();

                    ushort[] ushortInfoSend = new ushort[charInfoSend.Length];
                    byte[] sendBufSend = new byte[charInfoSend.Length * 2];
                    byte[] tempbfSend = new byte[2];
                    int j = 0;
                    for (int i = 0; i < charInfoSend.Length; i++)
                    {
                        ushortInfoSend[i] = (ushort)charInfoSend[i];
                        tempbfSend = BitConverter.GetBytes(ushortInfoSend[i]);
                        sendBufSend[j] = tempbfSend[1];
                        sendBufSend[j + 1] = tempbfSend[0];
                        j = j + 2;
                    }
                    try
                    {
                        if (!MyPort.IsOpen)
                        {
                            MyPort.Open();
                        }
                    }
                    catch (Exception ex)
                    {
                        RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGWARNING, CommonVary.myCOMPort + "被占用！");
                        MessageBox.Show(CommonVary.myCOMPort + "被占用！");
                    }
                    //byte[] sendBuf = new byte[]{48,49,50,65,66,67};

                    MyPort.Write(sendBufSend, 0, sendBufSend.Length);

                //}
                //catch (Exception ex)
                //{
                //    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGDEBUG, "发送规格到COM口出错\n信息：" + ex.Message);
                //    MessageBox.Show("发送规格到COM口出错\n信息：" + ex.Message);
                //    return;
                //}
        
            }

        }


        private void MyPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (startSystem == true)
            {
                readPortFlag = true;
                int n = MyPort.BytesToRead;
                byte[] readPort = new byte[n];
                MyPort.Read(readPort, 0, n);
                try
                {
                    if (readPort[readPort.Length - 1] == 49)
                    {
                        saveFlag = true;
                        //AutoUpdateSaveData(saveFlag, "正在保存数据！！！\n(频率：" + CommonVary.SaveDataInterval / 1000 + "秒每次)");
                    }
                    else
                    {
                        saveFlag = false;
                        //AutoUpdateSaveData(saveFlag, "所测数据没有保存！！！\n机器停止运行");
                    }
                }
                catch (Exception ex)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "传入起至位出错！");
                }
            }
        }
 

        private void AutoUpdateSaveData(bool saveFlag,string Info)
        {
            if (saveFlag)
            {
                
                TMSaveData.Interval = CommonVary.SaveDataInterval;
                TMSaveData.Enabled = true;
                label11.Text = Info;
                label11.Visible = true;
                label11.ForeColor = Color.Black;
            }
            else
            {
                TMSaveData.Enabled = false;
                label11.Text = Info;
                label11.Visible = true;
                label11.ForeColor = Color.Red;
            }

        }

        #endregion

        #region 数据删除

        private int DeleteDataEveryDay(DateTime startDateTime)
        {
            try
            {
                string clearSql = "delete from widthdata where timestamps <= '" + startDateTime + "'";
                CommonVary.OpenDataConnection();
                MySqlCommand cm = new MySqlCommand(clearSql, CommonVary.conn);
                cm.ExecuteNonQuery();
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "数据库每日删除完毕！\n 信息：" + startDateTime.ToLongDateString() + "之前数据删除");

            }
            catch (Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "清空数据库出错！\n 信息：" + ex.Message);
                return CommonVary.RUNNNING_WRONG;
            }
            return CommonVary.RUNNING_OK;
        }

        private int ClearHistryData()
        {
            try
            {
                string clearSql = "truncate table widthdata";
                CommonVary.OpenDataConnection();
                MySqlCommand cm = new MySqlCommand(clearSql, CommonVary.conn);
                cm.ExecuteNonQuery();
                MessageBox.Show("清空成功！");
            }
            catch (Exception ex)
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "清空数据库出错！\n 信息：" + ex.Message);
                return CommonVary.RUNNNING_WRONG;
            }
            return CommonVary.RUNNING_OK;
        }

        #endregion
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text == CommonVary.cbx3Item1)
            {
                sendMapData.sys_Inform.NumPos_NumWeb = (Int16)(sendMapData.sys_Inform.NumPos_NumWeb & 0xfffe);
            }
            else
            {
                sendMapData.sys_Inform.NumPos_NumWeb = (Int16)(sendMapData.sys_Inform.NumPos_NumWeb | 0x0001);
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox6.Text == CommonVary.cbx6Item2)
            {
                sendMapData.sys_Inform.NumPos_NumWeb = (Int16)(sendMapData.sys_Inform.NumPos_NumWeb | 0x0002);
            }
            else
            {
                sendMapData.sys_Inform.NumPos_NumWeb = (Int16)(sendMapData.sys_Inform.NumPos_NumWeb & 0x000D);
            }
        }

        private void button15_MouseDown(object sender, MouseEventArgs e)
        {
            sendMapData.sys_Inform.width1_lower_limit_analog_output = Int16.Parse(textBox51.Text);
            sendMapData.sys_Inform.width1_upper_limit_analog_output = Int16.Parse(textBox50.Text);
            sendMapData.sys_Inform.width2_lower_limit_analog_output = Int16.Parse(textBox48.Text);
            sendMapData.sys_Inform.width2_upper_limit_analog_output = Int16.Parse(textBox47.Text);
            sendMapData.sys_Inform.Setbit = 0x0001;

        }

        private void button15_MouseUp(object sender, MouseEventArgs e)
        {
            Thread.Sleep(500);
            sendMapData.sys_Inform.Setbit = 0x0000;
            MessageBox.Show("确认成功");

        }
        #region 修改20140115
        private void zedGraphControl1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ZedGraph.ZedGraphControl z = (ZedGraphControl)(sender);


            if (z.Height == h1)
            {
                SetZedGraphControlDefaultStyle();
                SetDefaultLocation();
                
                
            }
            else
            {
                zedGraphControl1.Visible = false;
                zedGraphControl2.Visible = false;
                zedGraphControl3.Visible = false;
                zedGraphControl4.Visible = false;
                z.Visible = true;
                z.Location = location;
                z.Height = h1;
                z.Width = w1;
                UpdateWidth(z);
            }
            
        }
        private void UpdateWidth(object sender)
        {
            ZedGraph.ZedGraphControl z = (ZedGraphControl)(sender);
            label1.Visible = false;
            textBox2.Visible = false;
            textBox1.Visible = false;
            textBox31.Visible = false;
            textBox32.Visible = false;

            label2.Visible = false;
            textBox4.Visible = false;
            textBox3.Visible = false;
            textBox34.Visible = false;
            textBox33.Visible = false;

            label3.Visible = false;
            textBox5.Visible = false;
            textBox6.Visible = false;
            textBox35.Visible = false;
            textBox36.Visible = false;

            label4.Visible = false;
            textBox7.Visible = false;
            textBox8.Visible = false;
            textBox37.Visible = false;
            textBox38.Visible = false;

            if (z.Equals(zedGraphControl1))
            {
                label1.Visible = true;
                textBox2.Visible = true;
                textBox1.Visible = true;
                textBox31.Visible = true;
                textBox32.Visible = true;
            }
            else if (z.Equals(zedGraphControl2))
            {
                label2.Visible = true;
                textBox4.Visible = true;
                textBox3.Visible = true;
                textBox34.Visible = true;
                textBox33.Visible = true;

                SetWidthLocation(label2, textBox4, textBox3, textBox34, textBox33);
            }
            else if (z.Equals(zedGraphControl3))
            {
                label3.Visible = true;
                textBox5.Visible = true;
                textBox6.Visible = true;
                textBox35.Visible = true;
                textBox36.Visible = true;
                SetWidthLocation(label3, textBox6, textBox5, textBox36, textBox35);
            }
            else if (z.Equals(zedGraphControl4))
            {
                label4.Visible = true;
                textBox7.Visible = true;
                textBox8.Visible = true;
                textBox37.Visible = true;
                textBox38.Visible = true;
                SetWidthLocation(label4, textBox8, textBox7, textBox38, textBox37);
            }
        }
        private void SetWidthLocation(System.Windows.Forms.Label l, TextBox t1, TextBox t2, TextBox t3, TextBox t4)
        {
            l.Location = label1.Location;
            t1.Location =textBox1.Location;
            t2.Location = textBox2.Location;
            t3.Location =textBox31.Location;
            t4.Location =textBox32.Location;
          
        }

        private void GetDefaultLocation()
        {
            WideLocation.Add(label1.Location);
            WideLocation.Add(textBox2.Location);
            WideLocation.Add(textBox1.Location);
            WideLocation.Add(textBox31.Location);
            WideLocation.Add(textBox32.Location);

            WideLocation.Add(label2.Location);
            WideLocation.Add(textBox4.Location);
            WideLocation.Add(textBox3.Location);
            WideLocation.Add(textBox34.Location);
            WideLocation.Add(textBox33.Location);

            WideLocation.Add(label3.Location);
            WideLocation.Add(textBox5.Location);
            WideLocation.Add(textBox6.Location);
            WideLocation.Add(textBox35.Location);
            WideLocation.Add(textBox36.Location);

            WideLocation.Add(label4.Location);
            WideLocation.Add(textBox7.Location);
            WideLocation.Add(textBox8.Location);
            WideLocation.Add(textBox37.Location);
            WideLocation.Add(textBox38.Location);

        }

        private void SetDefaultLocation()
        {
            label1.Visible = true;
            textBox2.Visible = true;
            textBox1.Visible = true;
            textBox31.Visible = true;
            textBox32.Visible = true;

            label2.Visible = true;
            textBox4.Visible = true;
            textBox3.Visible = true;
            textBox34.Visible = true;
            textBox33.Visible = true;

            label3.Visible = true;
            textBox5.Visible = true;
            textBox6.Visible = true;
            textBox35.Visible = true;
            textBox36.Visible = true;

            label4.Visible = true;
            textBox7.Visible = true;
            textBox8.Visible = true;
            textBox37.Visible = true;
            textBox38.Visible = true;

            label1.Location = WideLocation[0];   
            textBox2.Location = WideLocation[1]; 
            textBox1.Location = WideLocation[2]; 
            textBox31.Location = WideLocation[3];
            textBox32.Location = WideLocation[4];
                                    
            label2.Location = WideLocation[5];   
            textBox4.Location = WideLocation[6]; 
            textBox3.Location = WideLocation[7]; 
            textBox34.Location = WideLocation[8];
            textBox33.Location = WideLocation[9];
                                    
            label3.Location = WideLocation[10];   
            textBox5.Location = WideLocation[11]; 
            textBox6.Location = WideLocation[12]; 
            textBox35.Location = WideLocation[13];
            textBox36.Location = WideLocation[14];
                                    
            label4.Location = WideLocation[15];   
            textBox7.Location = WideLocation[16]; 
            textBox8.Location = WideLocation[17]; 
            textBox37.Location = WideLocation[18];
            textBox38.Location = WideLocation[19];


        }
        #endregion














    }
}
