using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
//
using MyCompany.ReportSystem.BLL;
using MyCompany.ReportSystem.COL;
using MyCompany.ReportSystem.SFL;
using System.IO;

using Emgu.CV.UI;
using Emgu.CV;
using Emgu.CV.Structure;
using Emgu.CV.CvEnum;
using System.Drawing.Imaging;

namespace MyCompany.ReportSystem.UIL.Product
{
    public partial class APOE : UserControl
    {
        #region 字段
        public static Video.WebCam wcam = null;
        private static APOE instance;
        public static string judgeId;
        private static string conditions = "[JudgeID] = 0";
        private Test_APOE testSheet;
        public List<string> list = new List<string>();
        public int flag = 0;  // 默认是0，添加；为1 则为修改
        String[] APOEChipInfo={"158T","112T","158C","112C"};
        #region //EMGU
        Emgu.CV.Capture grabber;
        Image<Bgr, Byte> currentFrameOri;
        Image<Bgr, Byte> currentFrame;
        Image<Gray, Byte> RectanglePic_b_gray;
        Image<Gray, Byte> RectanglePic_g_gray;
        Image<Gray, Byte> RectanglePic_r_gray;
        Image<Gray, Byte> RectanglePic_Reimg;
        Bitmap currentFramePic;
        Bitmap RectanglePic;
        Bitmap ResFramePic;

        int PointDet;  
        int PointClear;
        //红色矩形框参数
        static int IntAPOE_RedRectangleX = 300; //矩形框右上角的X值  470
        static int IntAPOE_RedRectangleY = 50; //矩形框右上角的Y值  310
        static int IntAPOE_RedRectangleWidth = 700; //矩形框宽度
        static int IntAPOE_RedRectangleHeigh = 900; //矩形框高度
        Rectangle APOE_RedRectangle = new Rectangle(IntAPOE_RedRectangleX, IntAPOE_RedRectangleY,
                                                    IntAPOE_RedRectangleWidth, IntAPOE_RedRectangleHeigh);//int x, int y, int width, int height

        //质控点检测绿色/蓝色线性参数
        static int IntAPOE_GreenLineRecXstep = 40;
        static int IntAPOE_GreenLineRecYstep = 20;
        static int IntAPOE_GreenLineFirstX = IntAPOE_RedRectangleX + IntAPOE_RedRectangleWidth - IntAPOE_GreenLineRecXstep; //线性第一个点X值
        static int IntAPOE_GreenLineFirstY = IntAPOE_RedRectangleY + IntAPOE_GreenLineRecYstep; //线性第一个点Y值
        static int IntAPOE_GreenLineSecondX = IntAPOE_RedRectangleX + IntAPOE_RedRectangleWidth - IntAPOE_GreenLineRecXstep; //线性第二个点X值
        static int IntAPOE_GreenLineSecondY = IntAPOE_RedRectangleY + IntAPOE_RedRectangleHeigh - IntAPOE_GreenLineRecYstep; //线性第二个点Y值
        LineSegment2DF APOE_GreenLine = new LineSegment2DF(new Point(IntAPOE_GreenLineFirstX, IntAPOE_GreenLineFirstY),
                                                           new Point(IntAPOE_GreenLineSecondX, IntAPOE_GreenLineSecondY));
                                                          //质控点检测绿色线形,矩形框上下间距各20，右边距40
        //质控点蓝色范围
        static int FirstZhiKongDetMax = 100;//第一个质控点检测范围
        LineSegment2DF APOE_BlueLineUp = new LineSegment2DF(new Point(IntAPOE_GreenLineFirstX - 1, IntAPOE_GreenLineFirstY + FirstZhiKongDetMax),
                                                           new Point(IntAPOE_GreenLineFirstX + 1, IntAPOE_GreenLineFirstY + FirstZhiKongDetMax));
                                                          //第一个质控点检测范围蓝色线形,上限
        LineSegment2DF APOE_BlueLineDown = new LineSegment2DF(new Point(IntAPOE_GreenLineSecondX - 1, IntAPOE_GreenLineSecondY - FirstZhiKongDetMax),
                                                          new Point(IntAPOE_GreenLineSecondX + 1, IntAPOE_GreenLineSecondY - FirstZhiKongDetMax));
                                                         //第一个质控点检测范围蓝色线形,下限

        private static PixelFormat[] indexedPixelFormats = { PixelFormat.Undefined, PixelFormat.DontCare,PixelFormat.Format16bppArgb1555, 
                                                             PixelFormat.Format1bppIndexed, PixelFormat.Format4bppIndexed,PixelFormat.Format8bppIndexed};
        //质控点参数
        int IntFirstZhiKongPointX;
        int IntFirstZhiKongPointY;
        static int IntXstepB = -60; //十字形之间X的B步长
        static int IntXstepS = -50; //十字形之间X的S步长
        static int[] XstepB = new int[] { 0, 1, 1, 2, 2 };
        static int[] XstepS = new int[] { 0, 0, 1, 1, 2 };
        static int IntYstep = 60; //十字形之间Y的步长
        static int CrossSize = 15; //十字形大小
        bool[,] APOEDetectResult = new bool[2, 2];
        
        
        #endregion
        #endregion

        /// <summary>
        /// 返回一个该控件的实例。如果之前该控件已经被创建，直接返回已创建的控件。
        /// 此处采用单键模式对控件实例进行缓存，避免因界面切换重复创建和销毁对象。
        /// </summary>
        public static APOE Instance
        {
             get
            {
                if (instance == null)
                {
                    instance = new APOE();
                }
                List<JudgeTemplet> list = new List<JudgeTemplet>();
                list = BLL.JudgeTempletBLL.GetListByName("APOE");
                if (list.Count > 0)
                {
                    judgeId = list[0].JudgeID.ToString();
                    conditions = " [JudgeID]=" + judgeId + " Order by  TestDate DESC, SampleID Asc ";
                }
                BindDataGrid(conditions);
                instance.testSheet = new Test_APOE();
                instance.InitControl();   // 每次返回该控件的实例前，都将对下拉框等界面显示控件的数据源进行初始化。
                instance.DisplaySinoChips();
                instance.BindObjectToForm(); // 每次返回该控件的实例前，都将关联对象的默认值，绑定至界面控件进行显示。
                return instance;
            }
        }

        /// <summary>
        /// 私有的控件实例化方法，创建实例只能通过该控件的Instance属性实现。
        /// </summary>
        private APOE()
        {
            InitializeComponent();
            this.toolStrip.CanOverflow = true;

            this.Dock = DockStyle.Fill;
            list.Clear();
            list.Add(" ");
            BtnChipDetect.Enabled = false;

            string mypath = "D:\\1";

            if (!Directory.Exists(mypath))
            {
                Directory.CreateDirectory(mypath);
            }

        }

        private void APOE_Load(object sender, EventArgs e)
        {
            start();

            this.TsBtnAdd.Click += new System.EventHandler(this.TsBtnAdd_Click);
            this.TsBtnSave.Click += new System.EventHandler(this.TsBtnSave_Click);
            this.TsBtnUpdate.Click += new System.EventHandler(this.TsBtnUpdate_Click);
            this.TsBtnDel.Click += new System.EventHandler(this.TsBtnDel_Click);
            this.TsBtnSearch.Click += new System.EventHandler(this.TsBtnSearch_Click);
            this.TsBtnSet.Click += new System.EventHandler(this.TsBtnSet_Click);
            this.TsBtnVideo.Click += new System.EventHandler(this.TsBtnVideo_Click);
            this.TsBtnVideoMove.Click += new System.EventHandler(this.TsVideo_Click);
            this.tsBtnSheet.Click += new System.EventHandler(this.tsBtnSheet_Click);
            this.TsBtmBackup.Click += new System.EventHandler(this.TsBtmBackup_Click);
            this.TsBtnRestore.Click += new System.EventHandler(this.TsBtnRestore_Click);
            this.TsBtnInfo.Click += new System.EventHandler(this.TsBtnInfo_Click);
            this.TsBtnExit.Click += new System.EventHandler(this.TsBtnExit_Click);

            this.BtnSelectRemark.Click += new System.EventHandler(this.BtnSelectRemark_Click);
            this.btnDelete.Click += new System.EventHandler(TsBtnDel_Click);
            this.btnSelectAll.Click += new System.EventHandler(this.BtnSelectAll_Click);
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            this.DgvGrid.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DgvGrid_CellDoubleClick);
            this.DgvGrid.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(DgvGrid_RowEnter);
            //this.DgvGrid.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(DgvGrid_CellClick);
            this.PageBar.PageChanged += new System.EventHandler(this.PageBar_PageChanged);

            this.BtnPrint.Click += new System.EventHandler(this.BtnPrint_Click);
          //  this.BtnBatchPrint.Click += new System.EventHandler(this.BtnBatchPrint_Click);

            this.TestDate.ValueChanged += new System.EventHandler(this.TestDate_ValueChanged);
            this.RBtnNotCheck.CheckedChanged += new System.EventHandler(this.TestDate_ValueChanged);
        }

        #region 界面控件与关联对象之间的绑定方法

        /// <summary>
        /// 初始化下拉框等控件的数据源。
        /// </summary>
        private void InitControl()
        {
            // 报告单位
            BindDrpList<COL.ReportUnit>(BLL.ReportUnitBLL.GetAllList(), DrpReportUnit, "UnitID", "UnitName");
            // 报告名称
            BindDrpList<COL.ReportName>(BLL.ReportNameBLL.GetAllList(judgeId), DrpReportName, "ReportID", "Name");
            // 标本状态
            BindDrpList<COL.SampleState>(BLL.SampleStateBLL.GetAllList(), DrpSampleState, "StateID", "StateName");
            //指定“住院科室”下拉框的数据源
            BindDrpList<COL.HospitalDept>(BLL.HospitalDeptBLL.GetAllList(), DrpHospitalDept, "DeptID", "DeptName");

            //指定“送检医师”下拉框的数据源
            BindDrpList<COL.Doctor>(BLL.DoctorBLL.GetAllList("送检"), DrpSendDoctor, "DoctorID", "DoctorName");

            //指定“检测医师”下拉框的数据源
            BindDrpList<COL.Doctor>(BLL.DoctorBLL.GetAllList("检测"), DrpTestDoctor, "DoctorID", "DoctorName");

            //指定“审核医师”下拉框的数据源
            BindDrpList<COL.Doctor>(BLL.DoctorBLL.GetAllList("审核"), DrpCheckDoctor, "DoctorID", "DoctorName");

            //指定“样本类型”下拉框的数据源
            BindDrpList<COL.SampleType>(BLL.SampleTypeBLL.GetAllList(judgeId), DrpSampleType, "TypeID", "TypeName");

            ////指定“检测方法”下拉框的数据源
            BindDrpList<COL.TestMethod>(BLL.TestMethodBLL.GetAllList(), DrpMethodName, "MethodID", "MethodName");
           
            ////指定“送检单位”下拉框的数据源
            BindDrpList<COL.SendUnit>(BLL.SendUnitBLL.GetAllList(), DrpSendUnit, "SendID", "SendName");
           
            // 绑定“性别”下拉框的数据源
            DrpGender.DisplayMember = "Name";
            DrpGender.ValueMember = "ID";
            DrpGender.DataSource = Sex.AllList;
            DrpGender.SelectedIndex = 0;
        }

        private void BindDrpList<T>(List<T> list, ComboBox drpList, string valueMember, string displayMember)
        {
            try
            {
                List<ListItem> items = new List<ListItem>();//添加项的集合

                drpList.DisplayMember = "Name";
                drpList.ValueMember = "Id";

                ListItem item;
                item = new ListItem("0", "--请选择--");
                items.Add(item);
                foreach (var v in list)
                {
                    items.Add(new ListItem(v.GetType().GetProperty(valueMember).GetValue(v, null).ToString(), v.GetType().GetProperty(displayMember).GetValue(v, null).ToString()));
                }
                drpList.DataSource = items;
            }
            catch
            {

            }

        }

        /// <summary>
        /// 将界面控件中的值，绑定给关联对象。
        /// </summary>
        private void BindFormlToObject()
        {
            if (testSheet == null)
            {
                testSheet = new Test_APOE();
            }
            #region
            //testSheet.JudgeID = (DataValid.GetNullOrInt(DrpJudgeTemplet.SelectedValue.ToString()) != null) ? DataValid.GetNullOrInt(DrpJudgeTemplet.SelectedValue.ToString()) : null;  // 住院科室
            testSheet.JudgeID = Int32.Parse(judgeId);
            testSheet.SampleID = DataValid.GetNullOrString(TxtSampleID.Text);  // 标本ID
            testSheet.SampleCode = DataValid.GetNullOrString(TxtSampleCode.Text);  // 标本代码
            if (DataValid.IsNullOrInt(TxtAge.Text.Trim()))
            {
                testSheet.Age = DataValid.GetNullOrInt(TxtAge.Text.Trim());
            }
            else
            {
                throw new CustomException("“年龄”须为整数，请您确认输入是否正确。");
            }
            testSheet.Phone = DataValid.GetNullOrString(TxtPhone.Text);  // 联系电话
            testSheet.PatientName = DataValid.GetNullOrString(TxtPatientName.Text);  // 姓名
            testSheet.CardNumber = DataValid.GetNullOrString(TxtCardNumber.Text);  // 病历号码
            testSheet.RoomID = DataValid.GetNullOrString(TxtRoomID.Text);  // 房/床号
            testSheet.AdmissionNum = DataValid.GetNullOrString(TxtAdmissionNum.Text);  // 门诊号
            testSheet.SendDate = DataValid.GetNullOrDateTime(DtpSendDate.Value.ToShortDateString());  // 送检日期
            testSheet.ReportDate = DataValid.GetNullOrDateTime(DtpReport.Value.ToShortDateString());   // 报告日期
            testSheet.Diagnosis = DataValid.GetNullOrString(txtDiagnosis.Text);
            testSheet.TestResult = DataValid.GetNullOrString(txtTestResult.Text);
            testSheet.ResultRemark = DataValid.GetNullOrString(txtResultRemark.Text);
            testSheet.TestDate = DataValid.GetNullOrDateTime(DtpTestDate.Value.ToShortDateString());  // 检测日期

            if (DrpGender.SelectedItem != null && DataValid.IsInt(DrpGender.SelectedValue.ToString()))
            {
                testSheet.Gender = Sex.GetDataById(DataValid.GetNullOrInt(DrpGender.SelectedValue.ToString()).Value);  // 性别
            }

            // 标本类型
            if (DrpSampleType.SelectedItem != null && DataValid.IsInt(DrpSampleType.SelectedValue.ToString()))
                testSheet.SampleType = (DataValid.GetNullOrInt(DrpSampleType.SelectedValue.ToString()) != null && DrpSampleType.SelectedValue.ToString() != "0") ? SampleTypeBLL.GetDataByTypeID(DataValid.GetNullOrInt(DrpSampleType.SelectedValue.ToString()).Value) : null;
            // 标本状态
            if (DrpSampleState.SelectedItem != null && DataValid.IsInt(DrpSampleState.SelectedValue.ToString()))
                testSheet.SampleState = (DataValid.GetNullOrInt(DrpSampleState.SelectedValue.ToString()) != null && DrpSampleState.SelectedValue.ToString() != "0") ? SampleStateBLL.GetDataByStateID(DataValid.GetNullOrInt(DrpSampleState.SelectedValue.ToString()).Value) : null;

            if (DrpReportName.SelectedItem != null && DataValid.IsInt(DrpReportName.SelectedValue.ToString()))
                testSheet.ReportName = (DataValid.GetNullOrInt(DrpReportName.SelectedValue.ToString()) != null && DrpReportName.SelectedValue.ToString() != "0") ? ReportNameBLL.GetDataByReportID(DataValid.GetNullOrInt(DrpReportName.SelectedValue.ToString()).Value) : null;
            // 住院科室
            if (DrpHospitalDept.SelectedItem != null && DataValid.IsInt(DrpHospitalDept.SelectedValue.ToString()))
                testSheet.HospitalDept = (DataValid.GetNullOrInt(DrpHospitalDept.SelectedValue.ToString()) != null && DrpHospitalDept.SelectedValue.ToString() != "0") ? HospitalDeptBLL.GetDataByDeptID(DataValid.GetNullOrInt(DrpHospitalDept.SelectedValue.ToString()).Value) : null;
            // 测试方法
            if (DrpMethodName.SelectedItem != null && DataValid.IsInt(DrpMethodName.SelectedValue.ToString()))
                testSheet.TestMethod = (DataValid.GetNullOrInt(DrpMethodName.SelectedValue.ToString()) != null && DrpMethodName.SelectedValue.ToString() != "0") ? TestMethodBLL.GetDataByMethodID(DataValid.GetNullOrInt(DrpMethodName.SelectedValue.ToString()).Value) : null;
            // 送检医师
            if (DrpSendDoctor.SelectedItem != null && DataValid.IsInt(DrpSendDoctor.SelectedValue.ToString()))
                testSheet.SendDoctor = (DataValid.GetNullOrInt(DrpSendDoctor.SelectedValue.ToString()) != null && DrpSendDoctor.SelectedValue.ToString() != "0") ? DoctorBLL.GetDataByDoctorID(DataValid.GetNullOrInt(DrpSendDoctor.SelectedValue.ToString()).Value) : null;
            // 检测医师
            if (DrpTestDoctor.SelectedItem != null && DataValid.IsInt(DrpTestDoctor.SelectedValue.ToString()))
                testSheet.TestDoctor = (DataValid.GetNullOrInt(DrpTestDoctor.SelectedValue.ToString()) != null && DrpTestDoctor.SelectedValue.ToString() != "0") ? DoctorBLL.GetDataByDoctorID(DataValid.GetNullOrInt(DrpTestDoctor.SelectedValue.ToString()).Value) : null;
            // 审核医师
            if (DrpCheckDoctor.SelectedItem != null && DataValid.IsInt(DrpCheckDoctor.SelectedValue.ToString()))
                testSheet.CheckDoctor = (DataValid.GetNullOrInt(DrpCheckDoctor.SelectedValue.ToString()) != null && DrpCheckDoctor.SelectedValue.ToString() != "0") ? DoctorBLL.GetDataByDoctorID(DataValid.GetNullOrInt(DrpCheckDoctor.SelectedValue.ToString()).Value) : null;
            // 报告单位
            if (DrpReportUnit.SelectedItem != null && DataValid.IsInt(DrpReportUnit.SelectedValue.ToString()))
                testSheet.ReportUnit = (DataValid.GetNullOrInt(DrpReportUnit.SelectedValue.ToString()) != null && DrpReportUnit.SelectedValue.ToString() != "0") ? ReportUnitBLL.GetDataByUnitID(DataValid.GetNullOrInt(DrpReportUnit.SelectedValue.ToString()).Value) : null; ;
            // 送检单位
            if (DrpSendUnit.SelectedItem != null  && DataValid.IsInt(DrpSendUnit.SelectedValue.ToString()))
                testSheet.SendUnit = (DataValid.GetNullOrInt(DrpSendUnit.SelectedValue.ToString()) != null && DrpSendUnit.SelectedValue.ToString() != "0") ? SendUnitBLL.GetDataBySendID(DataValid.GetNullOrInt(DrpSendUnit.SelectedValue.ToString()).Value) : null; ;

            #endregion
            testSheet._112C = list.Contains("112C") ? true : false;
            testSheet._112T = list.Contains("112T") ? true : false;
            testSheet._158C = list.Contains("158C") ? true : false;
            testSheet._158T = list.Contains("158T") ? true : false;

            testSheet.阳性探针 = list.Contains("阳性探针") ? true : false;
            testSheet.阴性探针 = list.Contains("阴性探针") ? true : false;
            testSheet.质控探针 = list.Contains("质控探针") ? true : false;
        }

        /// <summary>
        /// 将关联对象中的值，绑定至界面控件进行显示。
        /// </summary>
        private void BindObjectToForm()
        {
            //if (testSheet.JudgeID != null) DrpJudgeTemplet.SelectedValue = testSheet.JudgeID;

            TxtSampleID.Text = (testSheet.SampleID!=null)?testSheet.SampleID.ToString():string.Empty;  // 标本ID

            TxtSampleCode.Text = (testSheet.SampleCode != null) ? testSheet.SampleCode.ToString() : string.Empty;  // 标本代码

            TxtPatientName.Text = (testSheet.PatientName != null) ? testSheet.PatientName.ToString() : string.Empty;  // 姓名

            TxtAge.Text = (testSheet.Age != null) ? testSheet.Age.ToString() : string.Empty;   // 年龄

            TxtPhone.Text = (testSheet.Phone != null) ? testSheet.Phone : string.Empty; // 联系电话

            TxtCardNumber.Text = (testSheet.CardNumber != null) ? testSheet.CardNumber.ToString() : string.Empty;  // 病历号码

            TxtRoomID.Text = (testSheet.RoomID != null) ? testSheet.RoomID.ToString() : string.Empty;  // 房/床号

            TxtAdmissionNum.Text = (testSheet.AdmissionNum != null) ? testSheet.AdmissionNum.ToString() : string.Empty;  // 门诊号

            DtpTestDate.Value = testSheet.TestDate != null ? Convert.ToDateTime(testSheet.TestDate) : DateTime.Now;       // 检测日期
            DtpReport.Value = testSheet.ReportDate != null ? Convert.ToDateTime(testSheet.ReportDate) : DateTime.Now;     // 报告日期
            txtDiagnosis.Text = (testSheet.Diagnosis != null) ? testSheet.Diagnosis.ToString() : string.Empty;
            txtTestResult.Text = (testSheet.TestResult != null) ? testSheet.TestResult.ToString() : string.Empty;
            txtResultRemark.Text = (testSheet.ResultRemark != null) ? testSheet.ResultRemark.ToString() : string.Empty;
            DtpSendDate.Value = testSheet.SendDate != null ? Convert.ToDateTime(testSheet.SendDate) : DateTime.Now;       // 送检日期
            DtpSendDate.Value = testSheet.SendDate != null ? Convert.ToDateTime(testSheet.SendDate) : DateTime.Now;       // 送检日期

            DrpSampleType.SelectedValue = testSheet.SampleType != null ? testSheet.SampleType.TypeID.ToString() : "0";
            DrpHospitalDept.SelectedValue = testSheet.HospitalDept != null ? testSheet.HospitalDept.DeptID.ToString() : "0";
            DrpSampleState.SelectedValue = testSheet.SampleState != null ? testSheet.SampleState.StateID.ToString() : "0";// 标本状态
            DrpReportName.SelectedValue = testSheet.ReportName != null ? testSheet.ReportName.ReportID.ToString() : "0";
            DrpMethodName.SelectedValue =testSheet.TestMethod != null? testSheet.TestMethod.MethodID.ToString():"0";      // 测试方法
            DrpSendUnit.SelectedValue =testSheet.SendUnit != null? testSheet.SendUnit.SendID.ToString():"0";              // 送检部门

            DrpSendDoctor.SelectedValue = testSheet.SendDoctor != null ? testSheet.SendDoctor.DoctorID.ToString() : "0";  // 送检医生
            DrpTestDoctor.SelectedValue =testSheet.TestDoctor != null? testSheet.TestDoctor.DoctorID.ToString():"0";      // 检测医师
            DrpCheckDoctor.SelectedValue =testSheet.CheckDoctor != null? testSheet.CheckDoctor.DoctorID.ToString():"0";   // 审核医师
            DrpReportUnit.SelectedValue =testSheet.ReportUnit != null? testSheet.ReportUnit.UnitID.ToString():"0";        // 报告单位

            DrpGender.SelectedValue = (testSheet.Gender != null) ? testSheet.Gender.ID : 0; // 性别
            #region
            if (Convert.ToBoolean(testSheet._112C))
            {
                ClickState("112C");
                list.Add("112C");
            }
            if (Convert.ToBoolean(testSheet._112T))
            {
                ClickState("112T");
                list.Add("112T");
            }
            if (Convert.ToBoolean(testSheet._158C))
            {
                ClickState("158C");
                list.Add("158C");
            }
            if (Convert.ToBoolean(testSheet._158T))
            {
                ClickState("158T");
                list.Add("158T");
            }
           
            if (Convert.ToBoolean(testSheet.阳性探针))
            {
                ClickState("阳性探针");
                list.Add("阳性探针");
            }
            if (Convert.ToBoolean(testSheet.阴性探针))
            {
                ClickState("阴性探针");
                list.Add("阴性探针");
            }
            if (Convert.ToBoolean(testSheet.质控探针))
            {
                ClickState("质控探针");
                list.Add("质控探针");
            }
            this.pnlChip.Refresh();
            #endregion
        }

        private void ClickState(string txt)
        {
            if (flag == 1)
            {
                foreach (Control c in pnlChip.Controls)
                {
                    Button btn = c as Button;
                    if (txt.Equals(btn.Text))
                    {
                        if(Regex.IsMatch(txt, @"[\u4e00-\u9fa5]+$"))
                        {
                             btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ProbeA;
                        }
                        else
                        {
                             btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ItemA;
                        }
                    }
                }
            }
        }

        #endregion 界面控件与关联对象之间的绑定方法

        private void TsBtnAdd_Click(object sender, EventArgs e)
        {
            Reset();
            list.Clear();
            list.Add(" ");
            flag = 0;
        }

        /// <summary>
        /// 用户单击“保存”按钮时的事件处理方法。
        /// </summary>
        private void TsBtnSave_Click(object sender, EventArgs e)
        {
            BindFormlToObject(); // 进行数据绑定
            Bitmap bmp = ImageUtil.captureControl(this.pnlChip);

            string path = Application.StartupPath + @"\ChipPic\APOE";

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string picName = @"\ChipPic\APOE\" + DateTime.Now.ToString("yyyyMMdd") + "_" + Convert.ToString(testSheet.SampleID) + ".JPG";
            string picUrl = Application.StartupPath + picName;

            if (File.Exists(picUrl))
            {
                File.Delete(picUrl);
            }

            bmp.Save(picUrl);
            testSheet.ChipPic = picName;
            bmp.Dispose();
            if (flag == 0)
            {
                Test_APOEBLL.Insert(testSheet); // 调用“业务逻辑层”的方法，检查有效性后插入至数据库。
                list.Clear();
                list.Add(" ");
                FormSysMessage.ShowSuccessMsg("“化验单据”添加成功。");
            }
             if (flag == 1)
            {
                flag = 0;
                if (Test_APOEBLL.Update(testSheet) > 0) // 调用“业务逻辑层”的方法，检查有效性后插入至数据库。
                {
                    FormSysMessage.ShowSuccessMsg("“化验单据”修改成功。");
                }
                else
                {
                    FormSysMessage.ShowMessage("“化验单据”修改失败。");
                }
                list.Clear();
                list.Add(" ");
            }
            Reset();
            BindDataGrid(conditions);
        }

        private void Reset()
        {
            //TxtSampleID.Text = "";

            if (TxtSampleID != null)
            {
                int tmpTxtSampleID = Int32.Parse(TxtSampleID.Text);
                tmpTxtSampleID = tmpTxtSampleID + 1;
                TxtSampleID.Text = tmpTxtSampleID.ToString();
            }

            txtDiagnosis.Text = "";
            TxtCardNumber.Text = "";
            TxtPatientName.Text = "";
            TxtPhone.Text = "";
            txtResultRemark.Text = "结果备注：";
            TxtRoomID.Text = "";
            TxtSampleCode.Text = "";
            txtTestResult.Text = "";
            TxtAdmissionNum.Text = "";
            TxtAge.Text = "";
            DisplaySinoChips();  // 重新刷新芯片影响速度
            testSheet = null;
        }

        private string[] GetString(string str, string cutStr)
        {
            char[] cutChar = cutStr.ToCharArray();
            string[] sArray = str.Split(cutChar);
            return sArray;
        }

        public static int FindFirstPositionOfSubString(string sourceString, string childString)
        {
            int offset = 0;
            string substr = null;

            if ((sourceString == null) || (childString == null) || (sourceString.Length < childString.Length))
            {
                return -1;
            }
            else
            {
                while (offset <= (sourceString.Length - childString.Length))
                {
                    substr = sourceString.Substring(offset, childString.Length);
                    if (substr.Equals(childString))
                    {
                        return offset;
                    }
                    offset++;
                }
                return -1;
            }
        }

        private List<string> myJudgeResult(List<string> listTemp, Control btn)
        {
            string btnValue = btn.Text;
            
                //if (Regex.IsMatch(btnValue, @"[\u4e00-\u9fa5]+$"))
                //{
                //    foreach (Control control in this.pnlChip.Controls)
                //    {
                //        btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ProbeA;
                //        if (control is Button && control.Text == btn.Text)
                //        {
                //            control.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ProbeA;
                //        }
                //    }
                //}
                //else
                //{
                    foreach (Control control in this.pnlChip.Controls)
                    {
                        if (control is Button && control.Text == btn.Text)
                        {
                            control.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ItemA;
                        }
                    }
                    //btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ItemA;
                //}
                listTemp.Add(btnValue);
            
            return listTemp;
        }




        /// <summary>
        /// PageBar控件的当前页码发生变更时的事件处理方法。
        /// </summary>

        private List<string> JudgeResult(List<string> listTemp,Control btn)
        {
            string btnValue=btn.Text;
            if (listTemp.Contains(btnValue))
            {
                if (Regex.IsMatch(btnValue, @"[\u4e00-\u9fa5]+$"))
                {
                    foreach (Control control in this.pnlChip.Controls)
                    {
                        btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ProbeB;
                        if (control is Button && control.Text == btn.Text)
                        {
                            control.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ProbeB;
                        }
                    }
                }
                else
                {
                    btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ItemB;
                }
                listTemp.Remove(btnValue);
            }
            else
            {
                if (Regex.IsMatch(btnValue, @"[\u4e00-\u9fa5]+$"))
                {
                    foreach (Control control in this.pnlChip.Controls)
                    {
                        btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ProbeA;
                        if (control is Button && control.Text == btn.Text)
                        {
                            control.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ProbeA;
                        }
                    }
                }
                else
                {
                    btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ItemA;
                }
                listTemp.Add(btnValue);
            }
                return listTemp;
        }

        private string ResultStr(String[] arry)
        {
            string result = "";
            string judge = "";
            string style = "";

            for (int i = 0; i < arry.Length; i++)
            {
                if (!Regex.IsMatch(arry[i], @"[\u4e00-\u9fa5]+$"))
                {
                    List<TempletItem> items = BLL.TempletItemBLL.GetListById(judgeId);
                    for (int j = 0; j < items.Count; j++)
                    {
                        string value = items[j].JudgeContent;
                        if (value.Equals(arry[i]))
                        {
                             string str = items[j].JudgeDescribe;
                             if (Regex.IsMatch(str, @"[\u4e00-\u9fa5]+$"))
                           {
                               style += "," + str;
                           }
                             else
                           {
                               judge += value+",";
                           }
                        }
                    }
                }
            }
            judge =judge.TrimEnd(',');
            string[] judges = judge.Split(',');

            if (judges.Length == 2)
            {
                if (judge.Contains("112T") && judge.Contains("158T"))
                {
                    result = "ApoE基因112T+158T型（载脂蛋白E2/E2型）";
                }
                if (judge.Contains("112C") && judge.Contains("158C"))
                {
                    result = "ApoE基因112C+158C型（载脂蛋白E4/E4型）";
                }
                if (judge.Contains("112T") && judge.Contains("158C"))
                {
                    result = "ApoE基因112T+158C型（载脂蛋白E3/E3型）";
                }
            }
            if (judges.Length == 3)
            {
                if (judge.Contains("112T") && judge.Contains("158C") && judge.Contains("112C"))
                {
                    result = "ApoE基因112T/C+158C型（载脂蛋白E3/E4型）";
                }
                if (judge.Contains("112T") && judge.Contains("158C") && judge.Contains("158T"))
                {
                    result = "ApoE基因112T+158T/C型（载脂蛋白E2/E3型）";
                }
            }
            if (judges.Length == 4)
            {
                if (judge.Contains("112T") && judge.Contains("158C") && judge.Contains("158T") && judge.Contains("112C"))
                {
                    result = "ApoE基因112T/C+158T/C（载脂蛋白E2/E4型）";
                }
            }
            if (judges.Length >= 1 && result.Length <= 1)
            {
                result = "不在检测范围之内";
            }
            return result + style;
        }

        private void txt_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)(sender);
            list = JudgeResult(list, btn);

            List<String> listTemp = list;
       
            string TextResult = "";
            if (listTemp.Count > 0)
            {
                TextResult += listTemp[0];
                for (int i = 1; i < list.Count; i++)
                {
                    TextResult += "," + listTemp[i];
                }
            }
            else
            {
                TextResult = "";
            }

            TextResult = ResultStr(GetString(TextResult, ","));
            //  TextResult = TextResult.TrimEnd(',');
            txtTestResult.Text = TextResult;

            //if (TextResult.Length > 0)
            //{
            //    txtTestResult.Text = TextResult.Substring(0, TextResult.Length - 1);
            //}
            //else
            //{
            //    txtTestResult.Text = "";
            //}
        }

        private string ProduceString(int value, int row, int i, int j)  //row 列
        {
            string Str = "";
            if (value < (int)'A' + row)
            {
                Str = Convert.ToString((char)(value)) + (i + 1).ToString();
            }
            return Str;
        }

        private void DisplaySinoChips()
        {
                JudgeTemplet judgeTemplet = JudgeTempletBLL.GetDataByJudgeID(judgeId);
                string judgeChip = DataValid.GetNullOrString(judgeTemplet.JudgeChip);
                if (!String.IsNullOrEmpty(judgeChip) && Regex.IsMatch(judgeChip, @"[0-9]{1,2}[x][0-9]{1,2}$"))
                {
                    string[] values = GetString(judgeChip, "x");
                    int X = int.Parse(values[0]);
                    int Y = int.Parse(values[1]);
                    List<TempletItem> templetTtems = new List<TempletItem>();
                    
                    int width=((Y + 1) / 2 * 75 + ((Y + 1) / 2 + 1) * 2);
                    int height=(X * 30 + (X + 1) * 2);

                    pnlChip.Size = new Size( width,height); // Location(440, 1500)
                    pnlChip.Controls.Clear();

                    for (int i = 0; i < X; i++)       // 行 column            A1|B1  C1|D1  E1
                    {
                        for (int j = 0; j < Y; j = j + 2)  //  列 row          A2|B2    j=0 a1b1   j=2 c1|d1   j=4 e1f1
                        {
                            Button btn = null;
                            string strPre = ProduceString((int)'A' + j, Y, i, j);
                            string strAfter = ProduceString((int)'A' + j + 1, Y, i, j);
                            string txt = strPre + "|" + strAfter;
                            if (string.IsNullOrEmpty(strAfter))
                            {
                                txt = strPre;
                            }

                            List<TempletItem> listItem = TempletItemBLL.GetDataByValues(judgeId, txt);
                            if (listItem != null && listItem.Count != 0)
                            {
                                btn = new Button();
                                btn.Size = new Size(75, 30);
                                btn.FlatStyle = FlatStyle.Flat;
                                btn.Font = new Font(btn.Font.FontFamily, 9, btn.Font.Style | FontStyle.Bold);
                                btn.FlatAppearance.BorderSize = 0;
                                btn.TextAlign = ContentAlignment.MiddleRight;
                                btn.Click += new System.EventHandler(this.txt_Click);
                                btn.Location = new Point(8 + j / 2 * 75, 8 + i * 30);
                                btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ItemB;

                                if (listItem.Count == 1)
                                {
                                    TempletItem templetItem = listItem[0];
                                    if(!string.IsNullOrEmpty(listItem[0].JudgeContent))
                                    {
                                        btn.Text = listItem[0].JudgeContent;
                                        if(Regex.IsMatch(listItem[0].JudgeContent, @"[\u4e00-\u9fa5]+$"))
                                        {
                                            btn.Text = listItem[0].JudgeContent;
                                            btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ProbeB;
                                        }
                                    }
                                }
                                else
                                {
                                    for (int m = 0; m < listItem.Count; m++)
                                    {
                                        string[] probe = GetString(listItem[m].ProbePosition, "|");
                                        foreach (string s in probe)
                                        {
                                            if (txt.Equals(s))
                                            {
                                                TempletItem templetItem = listItem[m];
                                                btn.Text = listItem[m].JudgeContent;
                                                btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ProbeB;

                                            }
                                        }
                                    }
                                }

                                pnlChip.Controls.Add(btn);
                              }
                        }
                    }
                }
        }

        private void BtnSelectRemark_Click(object sender, EventArgs e)
        {
            FormRemark remark = new FormRemark(judgeId);
            remark.ShowDialog(this);
            if(remark.remarklist!=null)
            {
                txtResultRemark.Text = null;
                for(int i=0;i<remark.remarklist.Count;i++)
                {
                    txtResultRemark.Text += remark.remarklist[i].ToString() + "\r\n";
                }
            }
        }

        private void TsBtnSearch_Click(object sender, EventArgs e)
        {
            FormSearch searcher = new FormSearch(judgeId,"APOE");
            searcher.ShowDialog(this);
            if(!string.IsNullOrEmpty(searcher.updateID))
            {
                Reset();
                testSheet = Test_APOEBLL.GetDataBySheetID(Convert.ToInt32(searcher.updateID));
                BindObjectToForm();
            }

        }

        private void TsBtnUpdate_Click(object sender, EventArgs e)
        {
            int count = 0;
            string updateID = null;
            for (int i = 0; i < this.DgvGrid.Rows.Count; i++)
            {
                if (this.DgvGrid.Rows[i].Cells["ColCheck"].EditedFormattedValue.ToString() == "True")
                {
                    count++;
                    updateID = this.DgvGrid.Rows[i].Cells["ColSheetID"].Value.ToString().Trim();
                    if (count > 1)
                    {
                        updateID = null;
                        throw new CustomException("请确认您只选择了一项进行修改。");
                    }
                }
            }
            if (!string.IsNullOrEmpty(updateID))
            {
                Reset();
                flag = 1;
                //if (MessageBox.Show("修改化验单后，数据将不可恢复！确认修改吗？", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                //{
                instance.testSheet = Test_APOEBLL.GetDataBySheetID(Convert.ToInt32(updateID));
                    BindObjectToForm();
                //}
            }
        }

        private void TsBtnDel_Click(object sender, EventArgs e)
        {
            int count = 0;
            string deleteIDs = null;
            for (int i = 0; i < this.DgvGrid.Rows.Count; i++)
            {
                if (this.DgvGrid.Rows[i].Cells["ColCheck"].EditedFormattedValue.ToString() == "True")
                {
                    count++;
                    deleteIDs += this.DgvGrid.Rows[i].Cells["ColSheetID"].Value.ToString().Trim()+",";
                }
            }
            if (!string.IsNullOrEmpty(deleteIDs))
            {
                deleteIDs = deleteIDs.TrimEnd(',');
                if (MessageBox.Show("删除化验单后，数据不可恢复！确认删除 " + count + " 条记录吗？", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
               {
                     Test_APOEBLL.Delete(deleteIDs); // 调用“业务逻辑层”的方法，删除关联对象并更新至数据库。
                     FormSysMessage.ShowSuccessMsg("“化验单据”删除成功 !");
                     BindDataGrid(conditions);
                }
            }
        }

        private void TsBtnSet_Click(object sender, EventArgs e)
        {
            SystAdmin formMain = SystAdmin.Instance;
            formMain.ShowDialog();
            Reset();
            list.Clear();
            list.Add(" ");
            BindDataGrid(conditions);
            InitControl();
        }

        private void TsBtnVideo_Click(object sender, EventArgs e)
        {
            //wcam.Start();
            BtnChipDetect.Enabled = true;
            ///* comment for debug hcj 20160301
            initialise_capture();
            // * */
        }

        private void initialise_capture()
        {
            grabber = new Emgu.CV.Capture();
            grabber.QueryFrame();
            //Initialize the FrameGraber event
            Application.Idle += new EventHandler(FrameGrabber);
        }

        void FrameGrabber(object sender, EventArgs e)
        {
            //Get the current frame form capture device
            currentFrameOri = grabber.QueryFrame().Resize(1280, 1024, Emgu.CV.CvEnum.INTER.CV_INTER_CUBIC);
            currentFrame = currentFrameOri.Clone();
            //currentFrame.Save("D:\\1\\2015119.bmp");    //保存图片  

            //Convert it to Grayscale
            if (currentFrame != null)
            {
                //gray_frame = currentFrame.Convert<Gray, Byte>();
                Bitmap bmp = bmp = currentFrame.Bitmap;

                currentFrame.Draw(APOE_RedRectangle, new Bgr(Color.Red), 3);
                //currentFrame.Draw(APOE_GreenLine, new Bgr(Color.Green), 3);
                //currentFrame.Draw(APOE_BlueLineUp, new Bgr(Color.Blue), 3);
                //currentFrame.Draw(APOE_BlueLineDown, new Bgr(Color.Blue), 3);

                ReviewPicBox.Image = currentFrame.Bitmap;
            }
        }



        private void TsBtmBackup_Click(object sender, EventArgs e)
        {
            try
            {

                this.saveFileDialog1.Filter = "备份文件.mdb|*.MDB";
                this.saveFileDialog1.FileName = "ReportProject[" + DateTime.Now.ToString("yyyyMMdd") + "备份].mdb";
                if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    string fileName = this.saveFileDialog1.FileName.ToString();
                    if (fileName != null && fileName.Trim() != "")
                    {
                        System.IO.File.Copy(Application.StartupPath + "\\DataBase\\ReportProject.mdb", fileName, false);
                        FormSysMessage.ShowSuccessMsg("备份成功，请注意保存备份文件！");
                    }
                    else
                    {
                        throw new CustomException("没有指定目标文件名！");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new CustomException(ex.Message.ToString());
            }

        }

        private void TsBtnRestore_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("恢复后的数据将覆盖当前数据库，数据会丢失且不可恢复，建议先备份！\n\n确认继续导入吗？", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    this.openFileDialog1.Filter = "备份文件.mdb|*.MDB";
                    if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                    {

                        string fileName = this.openFileDialog1.FileName.ToString();
                        if (fileName != null && fileName.Trim() != "")
                        {
                            System.IO.File.Copy(fileName, Application.StartupPath + "\\DataBase\\ReportProject.mdb", true);
                            FormSysMessage.ShowSuccessMsg("数据库恢复成功！程序需及时关闭！");
                            Application.Exit();
                        }
                        else
                        {
                            throw new CustomException("没有选定待恢复的文件！");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new CustomException(ex.Message.ToString());
            }
        }

        private void TsBtnInfo_Click(object sender, EventArgs e)
        {
            FormShowInfo showInfo = new FormShowInfo();
            showInfo.Show(this);
        }

        private void TsBtnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void tsBtnSheet_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", Application.StartupPath + "\\Document");
        }

        private void TsVideo_Click(object sender, EventArgs e)
        {
            // 这个是判断，关闭
            // 获得任务管理器中的所有进程
            string fileName = "DLCWAPP.exe";
            string arg = Application.StartupPath + @"\";
            OpenPress(fileName, arg);
        }

        private static bool OpenPress(string FileName, string Arguments)
        {
            Process pro = new Process();
            if (System.IO.File.Exists(FileName))
            {
                pro.StartInfo.FileName = FileName;
                pro.StartInfo.Arguments = Arguments;
                pro.Start();
                return true;
            }
            return false;
        }

        public void start()
        {
            //以panel1为容器显示视频内容   

            wcam = new Video.WebCam(panel1.Handle, 0, 0, this.panel1.Width, this.panel1.Height);
        }

        public void KillProcess(string FileName)
        {
            Process[] p = Process.GetProcesses();
            foreach (Process p1 in p)
            {
                try
                {
                    string processName = p1.ProcessName.ToLower().Trim();
                    //判断是否包含阻碍更新的进程
                    if (processName.Equals(FileName, StringComparison.OrdinalIgnoreCase))
                    {
                        p1.Kill();
                    }
                }
                catch { }
            }
        }

        private void PrintSheet(Test_APOE testSheet)
        {
            Report report = new Report();

            string fileUrl = Application.StartupPath + @"\Model\APOE.dot";
            report.CreateNewDocument(fileUrl);

            #region 公共
            if (testSheet.ReportUnit != null)
                report.InsertValue("Hospital", testSheet.ReportUnit.UnitName);
            if (testSheet.ReportName != null)
                report.InsertValue("ReportName", testSheet.ReportName.Name);
            if (testSheet.PatientName != null)
                report.InsertValue("PatientName", testSheet.PatientName);
            if (testSheet.Age != null)
                report.InsertValue("Age", testSheet.Age.ToString());
            if (testSheet.Gender != null)
                report.InsertValue("Gender", testSheet.Gender.Name);
            if (testSheet.HospitalDept != null)
                report.InsertValue("HospitalDept", testSheet.HospitalDept.DeptName);
            if (testSheet.SendDoctor != null)
                report.InsertValue("SendDoctor", testSheet.SendDoctor.DoctorName);
            if (testSheet.SampleType != null)
                report.InsertValue("SampleType", testSheet.SampleType.TypeName);
            if (testSheet.SendDate != null)
                report.InsertValue("SendDate", Convert.ToDateTime(testSheet.SendDate).ToString("yyyy-MM-dd"));
            if (testSheet.SampleID != null)
                report.InsertValue("SampleID", testSheet.SampleID.ToString());
            if (!string.IsNullOrEmpty(testSheet.ChipPic))
            {
                string picPath = Application.StartupPath + testSheet.ChipPic;
                report.InsertPicture("ChipPic", picPath, 160, 70);
            }
            if (testSheet.TestDoctor != null)
                report.InsertValue("TestDoctor", testSheet.TestDoctor.DoctorName);
            if (testSheet.CheckDoctor != null)
                report.InsertValue("CheckDoctor", testSheet.CheckDoctor.DoctorName);
            if (testSheet.CardNumber != null)
                report.InsertValue("CardNumber", testSheet.CardNumber);
            if (testSheet.Diagnosis != null)
                report.InsertValue("Diagnosis", testSheet.Diagnosis);
            if (testSheet.SampleCode != null)
                report.InsertValue("SampleCode", testSheet.SampleCode);
            if (testSheet.ReportDate != null)
                report.InsertValue("ReportDate", Convert.ToDateTime(testSheet.ReportDate).ToString("yyyy-MM-dd"));
            if (testSheet.RoomID != null)
                report.InsertValue("RoomID", testSheet.RoomID);
            if (testSheet.SampleState != null)
                report.InsertValue("SampleState", testSheet.SampleState.StateName);
            if (testSheet.TestResult != null)
                report.InsertValue("TestResult", testSheet.TestResult);
            if (testSheet.TestMethod != null)
                report.InsertValue("TestMethod", testSheet.TestMethod.MethodName);
            if (testSheet.TestDate != null)
                report.InsertValue("TestDate", Convert.ToDateTime(testSheet.TestDate).ToString("yyyy-MM-dd"));
            if (testSheet.Phone != null)
                report.InsertValue("Phone", testSheet.Phone);
            if (testSheet.SendUnit != null)
                report.InsertValue("SendUnit", testSheet.SendUnit.SendName);
            if (testSheet.ResultRemark != null)
                report.InsertValue("ResultRemark", testSheet.ResultRemark);
            if (testSheet.AdmissionNum != null)
                report.InsertValue("AdmissionNum", testSheet.AdmissionNum);
            if (testSheet.ReportName != null && testSheet.ReportUnit != null)
            {   //.ToString("yyyy-MM-dd")
                report.InsertValue("Title", testSheet.ReportUnit.UnitName + testSheet.ReportName.Name);
            }
            #endregion

            string DateStr = DateTime.Now.ToString("yyyyMMdd");
            string path = Application.StartupPath + @"\Document\APOE\" + DateStr;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string documentName = path + @"\" + DateStr + "APOE_" + Convert.ToString(testSheet.SampleID) + ".doc";
            if (File.Exists(documentName))
            {
                File.Delete(documentName);
            }
            report.SaveDocument(documentName);
            report.killWinWordProcess();
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < this.DgvGrid.Rows.Count; i++)
            {
                if (this.DgvGrid.Rows[i].Cells[0].EditedFormattedValue.ToString() == "True")
                {
                    string myid = this.DgvGrid.Rows[i].Cells["ColSheetID"].Value.ToString().Trim();
                    Test_APOE testSheet = Test_APOEBLL.GetDataBySheetID(myid);
                    PrintSheet(testSheet);
                }
            }
        }

        private void TestDate_ValueChanged(object sender, EventArgs e)
        {
            string datetime = TestDate.Value.ToShortDateString();
            string condition;
            // 1.当天所有；2.当天没有检测清单
            if (RBtnNotCheck.Checked)
            {
                condition = " [JudgeID]=" + judgeId + " AND [TestResult] is NULL AND [TestDate]=#" + datetime + "#" + " Order by TestDate DESC, SampleID Asc "; 
            }
            else
            {
                condition = " [JudgeID]=" + judgeId + " AND [TestDate]=#" + datetime + "#" + " Order by TestDate DESC, SampleID Asc "; ;
            }
            BindDataGrid(condition);
        }



        private static void BindDataGrid(string conditions)
        {
            instance.PageBar.DataControl = instance.DgvGrid;
            instance.PageBar.DataSource = Test_APOEBLL.GetPageList(instance.PageBar.PageSize, instance.PageBar.CurPage, conditions);
            instance.PageBar.DataBind();
        }

        private void PageBar_PageChanged(object sender, EventArgs e)
        {
            BindDataGrid(conditions); //重新对DataGridView控件的数据源进行绑定。select * from smssend where ss_time between #2011-4-12 0:00:00# and #2011-4-13 0:00:00#
        }

        private void DgvGrid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Reset();
            string objID = DgvGrid["ColSheetID", e.RowIndex].Value.ToString();
            if (!string.IsNullOrEmpty(objID))
            {
                //if (MessageBox.Show("进入化验单修改模式？", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                //{
                list.Clear();
                list.Add(" ");
                instance.testSheet = Test_APOEBLL.GetDataBySheetID(Convert.ToInt32(objID));
                flag = 1;

                BindObjectToForm();
                //}
            }
        }

        private void DgvGrid_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (DgvGrid.Rows[e.RowIndex].Selected)
            {
                DgvGrid.Rows[e.RowIndex].Cells["ColCheck"].Value = true;
            }
            DgvGrid.Refresh();
        }

        private void DgvGrid_CellClick(object sender,DataGridViewCellEventArgs e)
        {
            if (Convert.ToBoolean(DgvGrid.Rows[e.RowIndex].Cells["ColCheck"].Value) == true)
            {
                DgvGrid.Rows[e.RowIndex].Cells["ColCheck"].Value = false;
            }
            else
            {
                DgvGrid.Rows[e.RowIndex].Cells["ColCheck"].Value = true;
            }   

            DgvGrid.Refresh();
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < this.DgvGrid.Rows.Count; i++)
            {
                DgvGrid.Rows[i].Cells["ColCheck"].Value = true;
            }
            DgvGrid.Refresh();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < this.DgvGrid.Rows.Count; i++)
            {
                DgvGrid.Rows[i].Cells["ColCheck"].Value = false;
            }
            DgvGrid.Refresh();
        }

        // 在一定区域内查找斑点
        private void findBlob(int bloblen, int rbh, int rbw, ref Image<Gray, byte> rb_r_dest,ref int minx,ref int miny,ref double minSum)
        {

            double th = 100.0;
            int queryRoiSum;
            int yPtr, xPtr;
            for (yPtr = 0; yPtr < rbh - bloblen - 1; yPtr++)
            {
                for (xPtr = 0; xPtr < rbw - bloblen - 1; xPtr++)
                {
                    queryRoiSum = 0;
                    for (int k = 0; k < bloblen; k++)
                    {
                        for (int l = 0; l < bloblen; l++)
                        {
                            queryRoiSum += rb_r_dest.Data[k + yPtr, l + xPtr, 0];
                        }
                    }

                    if (queryRoiSum < minSum)
                    {

                        minx = xPtr;
                        miny = yPtr;
                        minSum = queryRoiSum;
                    }

                }
            }
        }
        private int findBlobs(double leftStartX,double topStartY ,int rbw, int rbh, int offset, 
           int minx, int miny, int bloblen, double detectTh, Image<Gray, byte> roi_dst,String saveFileName)
        {
            leftStartX -= 0.5 * bloblen;
            topStartY -= 0.5 * bloblen;
            if (leftStartX < 0)
                leftStartX = 0;
            if (topStartY < 0)
                topStartY = 0;
            int roilen = (int)(bloblen * 1.8);
            Bitmap roi = roi_dst.Bitmap.Clone(new Rectangle((int)leftStartX,
                (int)topStartY, roilen, roilen), System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
//            roi.Save(saveFileName);
            Image<Gray, byte> roiPtr = new Image<Gray, byte>(roi);
            int blobminx=0;
            int blobminy=0;
            double minSum=bloblen*bloblen*255;
            findBlob(bloblen, roilen, roilen, ref roiPtr, ref blobminx, ref blobminy, ref minSum);

            //                    roiSum = CvInvoke.cvSum(new Image<Bgr, float>(roi));
            if (minSum < detectTh)
            {

                if (blobminx + bloblen > roilen)
                {
                    blobminx = roilen - bloblen;
                }
                if (blobminy + bloblen > roilen)
                {
                    blobminy = roilen - bloblen;
                }

                roi = roi.Clone(new Rectangle(blobminx,
                (int)blobminy, bloblen, bloblen), System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
                roi.Save(saveFileName);
                
                return 1;
            }
            roi = roi.Clone(new Rectangle(blobminx,
(int)blobminy, bloblen, bloblen), System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            roi.Save(saveFileName);
            return 0;
        }
        struct contourInfo
        {
            public int x;
            public int y;
            public int width;
            public int height;
        }
        private void BtnChipDetect_Click(object sender, EventArgs e)
        {
            BtnChipDetect.Enabled = false;

            //PointDet = 1;
            //PointClear = 1;

            PressBack();
            pictureBox2.Image = null;

            //if (PointDet1.Checked == true) PointDet = 1;
            //if (PointDet2.Checked == true) PointDet = 2;
            //if (PointClear1.Checked == true) PointClear = 1;
            //if (PointClear2.Checked == true) PointClear = 2;
            //if (PointClear3.Checked == true) PointClear = 3; 

            ///* comment for debug hcj 2016-03-01
            currentFramePic = currentFrameOri.ToBitmap();

            //按照矩形面积截取图像
            RectanglePic = currentFramePic.Clone(APOE_RedRectangle, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            

            // read image from disk 从磁盘读取图像
            //String temp = textBox1.Text;
            Image<Bgr, Byte> img = new Image<Bgr, Byte>(RectanglePic);

            // 将该图像转换为HSV
            Image<Hsv, float> imgHSV = new Image<Hsv, float>(img.Width,img.Height);
            Image<Bgr, float> img_float = new Image<Bgr, float>(img.Width, img.Height);
            CvInvoke.cvConvertScale(img, img_float, 1.0, 0);                          //将原图转化为float类型的数据
            CvInvoke.cvCvtColor(img_float, imgHSV, Emgu.CV.CvEnum.COLOR_CONVERSION.CV_BGR2HSV);      //根据图像的类型选择转换方式BGR2HSV，还有RGB2HSV
            imgHSV.Save("D:\\1\\1.jpg");

            // get the image from roi of img(HSV)

            Hsv botLimit = new Hsv(0, 0.15, 0);  //0,0.27,0
            Hsv uprLimit = new Hsv(50, 1, 255);   //50,1,255 textBox1
            //Hsv botLimit = new Hsv(float.Parse(textBox1.Text), float.Parse(textBox2.Text), float.Parse(textBox3.Text));  //0,0.27,0
            //Hsv uprLimit = new Hsv(float.Parse(textBox4.Text), float.Parse(textBox5.Text), float.Parse(textBox6.Text));   //50,1,255 textBox1
            Image<Gray, byte> imageHSVDest = imgHSV.InRange(botLimit, uprLimit);
            imageHSVDest.Save("D:\\1\\2.jpg");
            

            StructuringElementEx element = new StructuringElementEx(3, 3, 1, 1, Emgu.CV.CvEnum.CV_ELEMENT_SHAPE.CV_SHAPE_CROSS);
            CvInvoke.cvDilate(imageHSVDest, imageHSVDest, element, 10);
            imageHSVDest.Save("D:\\1\\3.jpg");
            CvInvoke.cvErode(imageHSVDest, imageHSVDest,element,10);
            imageHSVDest.Save("D:\\1\\33.jpg");
            CvInvoke.cvDilate(imageHSVDest, imageHSVDest, element, 10);
            imageHSVDest.Save("D:\\1\\4.jpg");
            //CvInvoke.cvCanny(imageHSVDest, imageHSVDest,120,300,3);
            //imageHSVDest.Save("D:\\1\\5.jpg");

            IntPtr Dyncontour = new IntPtr();//存放检测到的图像块的首地址  
            IntPtr Dynstorage = CvInvoke.cvCreateMemStorage(0);//开辟内存区域  
            int m = 88; 
            int n = CvInvoke.cvFindContours(imageHSVDest, Dynstorage, ref Dyncontour, m,
                Emgu.CV.CvEnum.RETR_TYPE.CV_RETR_EXTERNAL, Emgu.CV.CvEnum.CHAIN_APPROX_METHOD.CV_CHAIN_APPROX_SIMPLE, new Point(1, 1));

            Seq<Point> header = new Seq<Point>(Dyncontour,null);
            Seq<Point> contourTemp = header;
            IntPtr Dynstorage2 = CvInvoke.cvCreateMemStorage(0);//开辟内存区域  
            int cnt = 0;
            int ptr = 0;
            MemStorage stor = new MemStorage();
            contourInfo[] mContourInfo=new contourInfo[2];//用于存放找到的芯片，最多两个
            while (contourTemp != null)// 滤除一些比较小的矩形区域
            {
                if (contourTemp.BoundingRectangle.Height > 300 && contourTemp.BoundingRectangle.Width > 300)    //xx    bilichi
                {
                    mContourInfo[cnt].x = contourTemp.BoundingRectangle.X;
                    mContourInfo[cnt].y = contourTemp.BoundingRectangle.Y;
                    mContourInfo[cnt].width = contourTemp.BoundingRectangle.Width;
                    mContourInfo[cnt].height = contourTemp.BoundingRectangle.Height;
                    cnt++;
                    if (cnt == 2) // 根据x轴坐标对其进行排序，x较小的（即在图像中靠左的）放在mContourInfo[0]
                    {
                        if (mContourInfo[0].x > mContourInfo[1].x)
                        {
                            int swaptemp;
                            swaptemp = mContourInfo[0].x;
                            mContourInfo[0].x = mContourInfo[1].x;
                            mContourInfo[1].x = swaptemp;
                            swaptemp = mContourInfo[0].y;
                            mContourInfo[0].y = mContourInfo[1].y;
                            mContourInfo[1].y = swaptemp;
                            swaptemp = mContourInfo[0].width;
                            mContourInfo[0].width = mContourInfo[1].width;
                            mContourInfo[1].width = swaptemp;
                            swaptemp = mContourInfo[0].height;
                            mContourInfo[0].height = mContourInfo[1].height;
                            mContourInfo[1].height = swaptemp;
                        }
                    }
                    
                }
                contourTemp = contourTemp.HNext;
            }
            int index = 0;
            while(cnt--!=0)
            {

                    
                    //imageHSVDest.Draw(contourTemp.BoundingRectangle, new Gray(255), 2);
                CvInvoke.cvRectangle(imageHSVDest, new Point(mContourInfo[index].x, mContourInfo[index].y),
                    new Point(mContourInfo[index].x + mContourInfo[index].width, mContourInfo[index].y + mContourInfo[index].height),
                    new MCvScalar(255, 0, 0), 5, Emgu.CV.CvEnum.LINE_TYPE.EIGHT_CONNECTED, 0); 

                    imageHSVDest.Save("D:\\1\\5.jpg");
                    Rectangle rect = new Rectangle(mContourInfo[index].x, mContourInfo[index].y, mContourInfo[index].width, mContourInfo[index].height);
                    // 获取原始图像的矩形区域
                    Bitmap oRectanglePic = img.ToBitmap().Clone(rect, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
                    oRectanglePic.Save("D:\\1\\6.jpg");

                    // 截掉边界
                    int len = oRectanglePic.Width > oRectanglePic.Height ? oRectanglePic.Height : oRectanglePic.Width;
                    len = len - 60; // 60可调整至合适大小

                    int h = 3 * len / 4;
                    int resetTimes = 0; //用于动态微调区域
                resetPos:
                    RectanglePic = oRectanglePic.Clone(new Rectangle(oRectanglePic.Width / 2 - len / 2, oRectanglePic.Height / 2 - h / 2 - resetTimes * 10, len, h),
                        System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
                    RectanglePic.Save("D:\\1\\7.jpg");

                    int bloblen = RectanglePic.Width / 10;
                    int rbh = RectanglePic.Height / 2 - 3*bloblen / 5;// right bottom height
                    int rbw = RectanglePic.Width / 4;// right bottom width;
                    Bitmap rb = RectanglePic.Clone(new Rectangle(RectanglePic.Width - rbw, RectanglePic.Height - rbh, rbw, rbh),
                        System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
                    rb.Save("D:\\1\\8.jpg");

                    //RGB模型转换，灰度化取B值    
                    Image<Bgr, Byte> rb_bgr = new Image<Bgr, Byte>(rb);
                    //               Image<Gray, Byte> rb_r = rb_bgr.Split()[2];
                    //               Image<Gray, Byte> rb_g = rb_bgr.Split()[1];
                    //               Image<Gray, Byte> rb_b = rb_bgr.Split()[0];

                    // 将右下角截取到的图像进行hsv转换
                    Image<Hsv, float> rb_r_hsv = new Image<Hsv, float>(rb_bgr.Width, rb_bgr.Height);
                    Image<Bgr, float> rb_r_float = new Image<Bgr, float>(rb_bgr.Width, rb_bgr.Height);
                    CvInvoke.cvConvertScale(rb_bgr, rb_r_float, 1.0, 0);                          //将原图转化为float类型的数据
                    CvInvoke.cvCvtColor(rb_r_float, rb_r_hsv, Emgu.CV.CvEnum.COLOR_CONVERSION.CV_BGR2HSV);
                    Image<Gray, byte> rb_r_dest = rb_r_hsv.InRange(botLimit, uprLimit);
                    rb_r_dest.Save("D:\\1\\9.jpg");
                    //                CvInvoke.cvThreshold(rb_r_dest, rb_r_dest, 0, 255, 
                    //                    Emgu.CV.CvEnum.THRESH.CV_THRESH_OTSU | Emgu.CV.CvEnum.THRESH.CV_THRESH_BINARY_INV);
                    Byte[] rb_data = new Byte[rb_r_dest.Width * rb_r_dest.Height];

                    // 查找bloblen*bloblen最小的区域
                    Bitmap queryRoi;
                    double minSum = bloblen * bloblen * 255;
                    int minx = 0;
                    int miny = 0;

                    findBlob(bloblen, rbh, rbw, ref rb_r_dest, ref minx, ref miny, ref minSum);
                    // 这里得到了最右下角的点的坐标。
                    // 检验一下是否符合标准，当前是1/3的部分为背景色，此处可以调整。
                    if (minSum > 3*(bloblen * bloblen * 255) / 4)
                    {
                        // 提示没有找到相应的点
                        Console.WriteLine("Not Find First Point!");
                        if (resetTimes < 5)
                        {// 如果没有找到合适的点，进行微调
                            resetTimes++;
                            goto resetPos;
                        }
                        BtnChipDetect.Enabled = true;
                        return;
                    }
                    // 得到了最右下角的点的坐标，即该区域的左上角坐标为 (RectanglePic.Width - rbw + minx, RectanglePic.Height - rbh + miny）
                    // 该区域的中心点的坐标为（RectanglePic.Width - rbw + minx + bloblen/2, RectanglePic.Height - rbh + miny + bloblen/2）
                    rb = RectanglePic.Clone(new Rectangle(RectanglePic.Width - rbw + minx, RectanglePic.Height - rbh + miny, bloblen, bloblen),
                       System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
                    rb.Save("D:\\1\\10.jpg");

                    //获取整个检测区域的背景，然后根据对应检测点相应位置是否存在较大面积的黑色（即为0），如果是，则判断该位置是1
                    Image<Bgr, Byte> roi_bgr = new Image<Bgr, Byte>(RectanglePic);
                    Image<Hsv, float> roi_hsv = new Image<Hsv, float>(roi_bgr.Width, roi_bgr.Height);
                    Image<Bgr, float> roi_float = new Image<Bgr, float>(roi_bgr.Width, roi_bgr.Height);
                    CvInvoke.cvConvertScale(roi_bgr, roi_float, 1.0, 0);                          //将原图转化为float类型的数据
                    CvInvoke.cvCvtColor(roi_float, roi_hsv, Emgu.CV.CvEnum.COLOR_CONVERSION.CV_BGR2HSV);

                    Image<Gray, byte> roi_dst = roi_hsv.InRange(botLimit, uprLimit);
                    roi_dst.Save("D:\\1\\11.jpg");

                    

                    //pictureBox2.Image = roi_dst.Bitmap;
                    //roi_dst = roi_dst.SmoothMedian(19);
                    CvInvoke.cvDilate(roi_dst, roi_dst, element, 5);
                    roi_dst.Save("D:\\1\\12.jpg");
                    CvInvoke.cvErode(roi_dst, roi_dst, element, 5);
                    roi_dst.Save("D:\\1\\13.jpg");
                    //pictureBox2.Image = roi_dst.Bitmap;

                    //获取各个子区域  
                    // 获取子区域的时候，可以做一个优化，就是在通过固定offset计算出来的位置之后，再重新定位（在附近寻找最符合特征的位置）,后续优化。
                    int offset = roi_dst.Width / 5;//子区域之间间隔
                    
                    if (offset > 55)
                        offset = 55;
                    double detectTh = 3*(bloblen * bloblen * 255) / 4; //检测阈值，如果roi的灰度总和小于这个值，则判断该roi的值为1

                    double leftStartX = RectanglePic.Width - rbw + minx - offset * 4;//最左边的x坐标，有时可能等于0
                    double topStartY = RectanglePic.Height - rbh + miny - offset * 2; // 最上面的y坐标，有时有可能等于0
                    list.Clear();
                    int re = findBlobs(leftStartX, topStartY, rbw, rbh, offset, minx, miny, bloblen, detectTh, roi_dst, "D:\\1\\point11.jpg");
                    if (1==re && !list.Contains("112T")) 
                        list.Add("112T");
                    /*
                    leftStartX -= 0.5*bloblen;
                    topStartY -= 0.5*bloblen;
                    if (leftStartX < 0)
                        leftStartX = 0;
                    if (topStartY < 0)
                        topStartY = 0;
                    int roilen = (int)(bloblen * 1.5);
                    Bitmap roi = roi_dst.Bitmap.Clone(new Rectangle((int)leftStartX,
                        (int)topStartY, roilen, roilen), System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
//                    roi.Save("D:\\1\\point11.jpg");
                    Image<Gray, byte> roiPtr = new Image<Gray, byte>(roi);
//                    roiSum = CvInvoke.cvSum(new Image<Bgr, float>(roi));
                    if (minSum < detectTh)
                    {
                        if (minx + bloblen > roilen)
                        {
                            minx = roilen - bloblen;
                        }
                        if (miny + bloblen > roilen)
                        {
                            miny = roilen - bloblen;
                        }
                            
                        roi = roi.Clone(new Rectangle(minx,
                        (int)miny, bloblen, bloblen), System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
                        roi.Save("D:\\1\\point11.jpg");
                        if (!list.Contains("112T")) list.Add("112T");
                    }
                    */
                    leftStartX = RectanglePic.Width - rbw + minx - offset * 3;
                    re = findBlobs(leftStartX, topStartY, rbw, rbh, offset, minx, miny, bloblen, detectTh, roi_dst, "D:\\1\\point12.jpg");
                    if (1 == re && !list.Contains("112T"))
                        list.Add("112T");

                    leftStartX = RectanglePic.Width - rbw + minx - offset * 2;
                    re = findBlobs(leftStartX, topStartY, rbw, rbh, offset, minx, miny, bloblen, detectTh, roi_dst, "D:\\1\\point13.jpg");
                    if (1 == re && !list.Contains("158T"))
                        list.Add("158T");

                    leftStartX = RectanglePic.Width - rbw + minx - offset;
                    re = findBlobs(leftStartX, topStartY, rbw, rbh, offset, minx, miny, bloblen, detectTh, roi_dst, "D:\\1\\point14.jpg");
                    if (1 == re && !list.Contains("158T"))
                        list.Add("158T");

                    leftStartX = RectanglePic.Width - rbw + minx - offset * 4;
                    topStartY = RectanglePic.Height - rbh + miny - offset;
                    re = findBlobs(leftStartX, topStartY, rbw, rbh, offset, minx, miny, bloblen, detectTh, roi_dst, "D:\\1\\point21.jpg");
                    if (1 == re && !list.Contains("112C"))
                        list.Add("112C");

                    leftStartX = RectanglePic.Width - rbw + minx - offset * 3;
                    re = findBlobs(leftStartX, topStartY, rbw, rbh, offset, minx, miny, bloblen, detectTh, roi_dst, "D:\\1\\point22.jpg");
                    if (1 == re && !list.Contains("112C"))
                        list.Add("112C");

                    leftStartX = RectanglePic.Width - rbw + minx - offset * 2;
                    re = findBlobs(leftStartX, topStartY, rbw, rbh, offset, minx, miny, bloblen, detectTh, roi_dst, "D:\\1\\point23.jpg");
                    if (1 == re && !list.Contains("158C"))
                        list.Add("158C");

                    leftStartX = RectanglePic.Width - rbw + minx - offset ;
                    re = findBlobs(leftStartX, topStartY, rbw, rbh, offset, minx, miny, bloblen, detectTh, roi_dst, "D:\\1\\point24.jpg");
                    if (1 == re && !list.Contains("158C"))
                        list.Add("158C");

                    DetList(list);

                    index++; // 指向下一个区域


                    /*
                                    //二值化   =======自动取值替代140   OTSU
                                    Image<Gray, Byte> RectanglePic_b_threshimg = RectanglePic_b_gray.ThresholdBinary(new Gray(140), new Gray(256));
                                    Image<Gray, Byte> RectanglePic_b_threshimg_not = RectanglePic_b_threshimg.Not();

                                    //中值滤波
                                    RectanglePic_Reimg = new Image<Gray, byte>(RectanglePic_b_threshimg_not.MIplImage.width, RectanglePic_b_threshimg_not.MIplImage.height);
                                    CvInvoke.cvCopy(RectanglePic_b_threshimg_not.Ptr, RectanglePic_Reimg.Ptr, System.IntPtr.Zero);
                                    //MedianFilter(RectanglePic_b_threshimg_not, 11);  //7为等级
                                    //RectanglePic_b_threshimg_not.SmoothMedian(101);
                                    if (PointClear == 1)  RectanglePic_b_threshimg_not = RectanglePic_b_threshimg_not.SmoothMedian(15);
                                    if (PointClear == 2) RectanglePic_b_threshimg_not = RectanglePic_b_threshimg_not.SmoothMedian(19);
                                    if (PointClear == 3) RectanglePic_b_threshimg_not = RectanglePic_b_threshimg_not.SmoothMedian(23);

                                    ResFramePic = RectanglePic_b_threshimg_not.Bitmap;
                                    using (ResFramePic)
                                    {
                                        //如果原图片是索引像素格式之列的，则需要转换
                                        if (IsPixelFormatIndexed(ResFramePic.PixelFormat))
                                        {
                                            Bitmap bmp = new Bitmap(ResFramePic.Width, ResFramePic.Height, PixelFormat.Format32bppArgb);
                                            using (Graphics g = Graphics.FromImage(bmp))
                                            {
                                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                                                g.DrawImage(ResFramePic, 0, 0);
                                            }

                                            //选择第一个质控点的位置
                                            int FirstDetectY = IntAPOE_GreenLineRecYstep;
                                            int FirstDetectMaxY = FirstZhiKongDetMax;
                                            int y = 0;
                                            IntFirstZhiKongPointX = IntAPOE_RedRectangleWidth - IntAPOE_GreenLineRecXstep;

                                            for (y = 0; y <= FirstDetectMaxY; y++)   //下行探测第一个质控点
                                            {
                                                //g2.DrawLine(p2, new Point(FirstDetectX-5, FirstDetectY + y),
                                                //                 new Point(FirstDetectX+5, FirstDetectY + y));
                                                if (IsGrayInt0_HengDet(RectanglePic_b_threshimg_not, IntFirstZhiKongPointX, FirstDetectY + y, CrossSize))
                                                {
                                                    //g2.DrawLine(p2, new Point(FirstDetectX - 10, FirstDetectY + y),
                                                    //                  new Point(FirstDetectX + 10, FirstDetectY + y));
                                                    if (IsGrayInt0_HengDet(RectanglePic_b_threshimg_not, IntFirstZhiKongPointX, FirstDetectY + y + (CrossSize / 2), CrossSize))
                                                    {
                                                        Console.WriteLine("Up Find First Point!");
                                                        IntFirstZhiKongPointY = FirstDetectY + y + (CrossSize / 2); //
                                                        break;
                                                    }
                                                }
                                            }

                                            if (y >= FirstDetectMaxY)
                                            {
                                                for (y = 0; y <= FirstDetectMaxY; y++)   //上行探测第一个质控点
                                                {
                                                    //g2.DrawLine(p2, new Point(FirstDetectX-5, FirstDetectY + y),
                                                    //                 new Point(FirstDetectX+5, FirstDetectY + y));
                                                    if (IsGrayInt0_HengDet(RectanglePic_b_threshimg_not, IntFirstZhiKongPointX, IntAPOE_RedRectangleHeigh - IntAPOE_GreenLineRecYstep - y, CrossSize))
                                                    {
                                                        //g2.DrawLine(p2, new Point(FirstDetectX - 10, FirstDetectY + y),
                                                        //                  new Point(FirstDetectX + 10, FirstDetectY + y));
                                                        if (IsGrayInt0_HengDet(RectanglePic_b_threshimg_not, IntFirstZhiKongPointX, IntAPOE_RedRectangleHeigh - IntAPOE_GreenLineRecYstep - y - (CrossSize / 2), CrossSize))
                                                        {
                                                            Console.WriteLine("Down Find First Point!");
                                                            IntFirstZhiKongPointY = IntAPOE_RedRectangleHeigh - IntAPOE_GreenLineRecYstep - y - (CrossSize / 2) - IntYstep * 2;
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (y >= FirstDetectMaxY)
                                                {
                                                    IntFirstZhiKongPointY = 40;
                                                    Console.WriteLine("Not Find First Point!");
                                                }
                                            }


                                            //明确第一个质控点之后，开始按矩阵判断信号点

                                            //for (int a = 0; a < 2; a++) for (int b = 0; b < 2; b++) APOEDetectResult[a, b] = false;
                    
                                            PressBack(); //清空list，和重值信号点

                                            for (int i = 0; i <= 1; i++)   //APOE,i为3行，j为5列
                                            {
                                                //Gray gray = RectanglePic_b_threshimg_not[IntFirstWriteCrossY + IntYstep * i,IntFirstWriteCrossX];   //读取像素值,要注意行列值
                                                //if (gray.Intensity == 0)
                                                for (int j = 0; j <= 4; j++)
                                                {
                                                    if (PointDet == 1)
                                                    {
                                                        if (IsGrayInt0_CroDet(RectanglePic_b_threshimg_not, IntFirstZhiKongPointX + XstepB[j] * IntXstepB + XstepS[j] * IntXstepS,
                                                                                                          IntFirstZhiKongPointY + IntYstep * i, CrossSize))   //CrossSize为红十字形判断范围值
                                                        {
                                                            Point CorssMiddlePoint = new Point(IntFirstZhiKongPointX + XstepB[j] * IntXstepB + XstepS[j] * IntXstepS, IntFirstZhiKongPointY + IntYstep * i);
                                                            DrawCross(bmp, CorssMiddlePoint, CrossSize); //画size为CrossSize的十字形
                                                            if (j > 0)
                                                            {
                                                                //textBox1.Text += i.ToString() + "行" + j.ToString() + "列" + "  ";
                                                                if (j == 1 || j == 2)
                                                                {
                                                                    //int DetectResultJ = 0;
                                                                    //APOEDetectResult[i, DetectResultJ] = true;
                                                                    if (i == 0 && !list.Contains("158T")) list.Add("158T");
                                                                    if (i == 1 && !list.Contains("158C")) list.Add("158C");
                                                                }
                                                                if (j == 3 || j == 4)
                                                                {
                                                                    //int DetectResultJ = 1;
                                                                    //APOEDetectResult[i, DetectResultJ] = true;
                                                                    if (i == 0 && !list.Contains("112T")) list.Add("112T");
                                                                    if (i == 1 && !list.Contains("112C")) list.Add("112C");
                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (PointDet == 2)
                                                    {
                                                        if (IsGrayInt0_RecDet(RectanglePic_b_threshimg_not, IntFirstZhiKongPointX + XstepB[j] * IntXstepB + XstepS[j] * IntXstepS,
                                                                                                          IntFirstZhiKongPointY + IntYstep * i, CrossSize))   //CrossSize为红十字形判断范围值
                                                        {
                                                            Point CorssMiddlePoint = new Point(IntFirstZhiKongPointX + XstepB[j] * IntXstepB + XstepS[j] * IntXstepS, IntFirstZhiKongPointY + IntYstep * i);
                                                            DrawCross(bmp, CorssMiddlePoint, CrossSize); //画size为CrossSize的十字形
                                                            if (j > 0)
                                                            {
                                                                //textBox1.Text += i.ToString() + "行" + j.ToString() + "列" + "  ";
                                                                if (j == 1 || j == 2)
                                                                {
                                                                    //int DetectResultJ = 0;
                                                                    //APOEDetectResult[i, DetectResultJ] = true;
                                                                    if (i == 0 && !list.Contains("158T")) list.Add("158T");
                                                                    if (i == 1 && !list.Contains("158C")) list.Add("158C");
                                                                }
                                                                if (j == 3 || j == 4)
                                                                {
                                                                    //int DetectResultJ = 1;
                                                                    //APOEDetectResult[i, DetectResultJ] = true;
                                                                    if (i == 0 && !list.Contains("112T")) list.Add("112T");
                                                                    if (i == 1 && !list.Contains("112C")) list.Add("112C");
                                                                }
                                                            }
                                                        }
                                                    }
                                                 }
                                            }

                    
                                            //DetResPressBtn(APOEDetectResult);
                                            DetList(list);
                                            pictureBox2.Image = bmp;
                                        }
                                        else //否则直接操作
                                        {
                                            //直接对img进行水印操作
                                            Graphics g2 = Graphics.FromImage(ResFramePic);
                                            Brush b = new SolidBrush(Color.Red);
                                            Pen p = new Pen(b);
                                            g2.DrawRectangle(p, IntFirstZhiKongPointX, IntFirstZhiKongPointY, 20, 20);

                                            pictureBox2.Image = ResFramePic;
                                        }
                                    }*/

                    BtnChipDetect.Enabled = true;
            }
            BtnChipDetect.Enabled = true;
        }

        private void PressBack()
        {
            list.Clear();
            list.Add(" ");
            foreach (Button tmpbt in pnlChip.Controls)
            {
                foreach (string tmps in APOEChipInfo)
                {
                    if (tmpbt.Text == tmps)
                    {
                        tmpbt.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ItemB;    //ItemB is null
                    }
                }
            }
        }

               
        private static bool IsPixelFormatIndexed(PixelFormat imgPixelFormat)
        {
            foreach (PixelFormat pf in indexedPixelFormats)
            {
                if (pf.Equals(imgPixelFormat)) return true;
            }

            return false;
        }

        private Boolean IsGrayInt0_HengDet(Image<Gray, Byte> ImageGray, int xx, int yy, int step)
        {
            Gray g1 = ImageGray[yy, xx];   //读取像素值，注意这里是yy行xx列
            if (g1.Intensity == 0) return true;
            for (int gg = 1; gg < step; gg++)
            {
                Gray g2 = ImageGray[yy, xx - gg];   //读取像素值
                Gray g3 = ImageGray[yy, xx + gg];   //读取像素值
                //Gray g4 = ImageGray[xx, yy - step];   //读取像素值
                //Gray g5 = ImageGray[xx, yy + step];   //读取像素值
                if ((g2.Intensity == 0) || (g3.Intensity == 0))
                {
                    return true;
                }
            }
            return false;
        }

        private Boolean IsGrayInt0_CroDet(Image<Gray, Byte> ImageGray, int xx, int yy, int step)
        {
            Gray g1 = ImageGray[yy, xx];   //读取像素值，注意这里是yy行xx列
            if (g1.Intensity == 0) return true;
            for (int gg = 1; gg < step; gg++)
            {
                Gray g2 = ImageGray[yy, xx - gg];   //读取像素值
                Gray g3 = ImageGray[yy, xx + gg];   //读取像素值
                Gray g4 = ImageGray[yy - step, xx];   //读取像素值
                Gray g5 = ImageGray[yy + step, xx];   //读取像素值
                if ((g2.Intensity == 0) || (g3.Intensity == 0) || (g4.Intensity == 0) || (g5.Intensity == 0))
                {
                    return true;
                }
            }
            return false;
        }

        private Boolean IsGrayInt0_RecDet(Image<Gray, Byte> ImageGray, int xx, int yy, int step)
        {
            Gray g1 = ImageGray[yy, xx];   //读取像素值，注意这里是yy行xx列
            if (g1.Intensity == 0) return true;
            for (int gx = 1; gx <= (step * 2); gx++)
            {
                for (int gy = 1; gy <= (step * 2); gy++)
                {
                    Gray g2 = ImageGray[yy - step + gy, xx - step + gx];   //读取像素值
                    if (g2.Intensity == 0) return true;
                }
            }
            return false;
        }

        private void DrawCross(Bitmap bitmap, Point MiddlePoint, int size) //在bitmap上画十字形,红色，粗为3，大小为size
        {
            if ((MiddlePoint.X > size) && (MiddlePoint.Y > size))
            {
                Graphics g1 = Graphics.FromImage(bitmap);
                Brush b = new SolidBrush(Color.Red);
                Pen p = new Pen(b, 3);
                MiddlePoint.Offset(-size, 0);
                Point LeftPoint = MiddlePoint;
                MiddlePoint.Offset(size, 0);
                MiddlePoint.Offset(size, 0);
                Point RightPoint = MiddlePoint;
                MiddlePoint.Offset(-size, 0);
                MiddlePoint.Offset(0, -size);
                Point UpPoint = MiddlePoint;
                MiddlePoint.Offset(0, size);
                MiddlePoint.Offset(0, size);
                Point DownPoint = MiddlePoint;

                g1.DrawLine(p, LeftPoint, RightPoint);
                g1.DrawLine(p, UpPoint, DownPoint);
            }
            else
            {
                Console.WriteLine("逻辑错误，参数越界");
            }
        }


        private void DetList(List<string> list)
        {
            System.Windows.Forms.Button btn = new System.Windows.Forms.Button();

            List<String> listTemp = list;
            int Clist = listTemp.Count;// 一维长度（行数）
            for (int tmpi = 0; tmpi < Clist; tmpi++)
            {
                foreach (Control control in this.pnlChip.Controls)
                {
                    if (control is Button && control.Text == list[tmpi])
                    {
                        control.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.ItemA;
                    }
                }
            }
            string TextResult = "";
            if (Clist > 0)
            {
                TextResult += listTemp[0];
                for (int i = 1; i < Clist; i++)
                {
                    TextResult += "," + listTemp[i];
                }
            }
            else
            {
                TextResult = "";
            }

            TextResult = ResultStr(GetString(TextResult, ","));
            //  TextResult = TextResult.TrimEnd(',');
            txtTestResult.Text = TextResult;                       
                         
        }


        private void DetResPressBtn(bool[,] DetectResult)
        {
            System.Windows.Forms.Button btn = new System.Windows.Forms.Button();

            int rowsC = DetectResult.GetLength(0);// 一维长度（行数）
            int columnsC = DetectResult.GetLength(1); // 二维长度(列数)
            for (int tmpi = 0; tmpi < rowsC; tmpi++)
            {
                for (int tmpj = 0; tmpj < columnsC; tmpj++)
                {
                    if (DetectResult[tmpi, tmpj] == true)
                    {
                        
                        //textBox1.Text += "\r\n TRUE:" + tmpi + "行" + tmpj + "列";   //得到X行Y列结果
                        if (tmpi == 0 && tmpj == 0) btn.Text = "158T";
                        if (tmpi == 0 && tmpj == 1) btn.Text = "112T";
                        if (tmpi == 1 && tmpj == 0) btn.Text = "158C";
                        if (tmpi == 1 && tmpj == 1) btn.Text = "112C";

                        list = myJudgeResult(list, btn);

                        List<String> listTemp = list;

                        string TextResult = "";
                        if (listTemp.Count > 1)
                        {
                            TextResult += listTemp[1];
                            for (int i = 2; i < list.Count; i++)
                            {
                                TextResult += "," + listTemp[i];
                            }
                        }
                        else
                        {
                            TextResult = "";
                        }

                        TextResult = ResultStr(GetString(TextResult, ","));
                        //  TextResult = TextResult.TrimEnd(',');
                        txtTestResult.Text = TextResult;
                       
                    }
                 }
            }

            
            
        }

        //private void BtnBatchPrint_Click(object sender, EventArgs e)
        //{
        //    string ids = string.Empty;
        //    for (int i = 0; i < this.DgvGrid.Rows.Count; i++)
        //    {
        //        if (this.DgvGrid.Rows[i].Cells[0].EditedFormattedValue.ToString() == "True")
        //        {
        //            string myid = this.DgvGrid.Rows[i].Cells["ColSheetID"].Value.ToString().Trim();
        //            ids += myid + ",";
        //        }
        //    }

        //    if (ids.Length == 0)
        //    {
        //        FormSysMessage.ShowMessage("请确认您已选择报表。");
        //        return;
        //    }
        //    ids = ids.TrimEnd(',');
        //    FormPrint crReport = new FormPrint("Test_HBV_GT",judgeId, ids);
        //    crReport.ShowDialog();
        //}
    }
}
