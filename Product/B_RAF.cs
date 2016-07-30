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
using MyCompany.ReportSystem.BLL;
using MyCompany.ReportSystem.COL;
using MyCompany.ReportSystem.SFL;
using System.IO;

namespace MyCompany.ReportSystem.UIL.Product
{
    public partial class B_RAF : UserControl
    {
        #region 字段
        public static Video.WebCam wcam = null;
        private static B_RAF instance;
        public static string judgeId;
        private static string conditions = "[JudgeID] = 0";
        private Test_B_raf testSheet;
        public List<string> list=new List<string>();
        public int flag = 0;  // 默认是0，添加；为1 则为修改
        #endregion

        /// <summary>
        /// 返回一个该控件的实例。如果之前该控件已经被创建，直接返回已创建的控件。
        /// 此处采用单键模式对控件实例进行缓存，避免因界面切换重复创建和销毁对象。
        /// </summary>
        public static B_RAF Instance
        {
             get
            {
                if (instance == null)
                {
                    instance = new B_RAF();
                }
                List<JudgeTemplet> list = new List<JudgeTemplet>();
                list = BLL.JudgeTempletBLL.GetListByName("B_raf");
                if (list.Count > 0)
                {
                    judgeId = list[0].JudgeID.ToString();
                    conditions = " [JudgeID]=" + judgeId + " Order by  TestDate DESC, SampleID Asc ";
                }
                BindDataGrid(conditions);
                instance.testSheet = new Test_B_raf();
                instance.InitControl();   // 每次返回该控件的实例前，都将对下拉框等界面显示控件的数据源进行初始化。
                instance.DisplaySinoChips();
                instance.BindObjectToForm(); // 每次返回该控件的实例前，都将关联对象的默认值，绑定至界面控件进行显示。
                return instance;
            }
        }

        /// <summary>
        /// 私有的控件实例化方法，创建实例只能通过该控件的Instance属性实现。
        /// </summary>
        private B_RAF()
        {
            InitializeComponent();
            this.toolStrip.CanOverflow = true;

            this.Dock = DockStyle.Fill;
            list.Add(" ");
        }

        private void B_RAF_Load(object sender, EventArgs e)
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
                testSheet = new Test_B_raf();
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
            if (DrpSendUnit.SelectedItem != null && DataValid.IsInt(DrpSendUnit.SelectedValue.ToString()))
                testSheet.SendUnit = (DataValid.GetNullOrInt(DrpSendUnit.SelectedValue.ToString()) != null && DrpSendUnit.SelectedValue.ToString() != "0") ? SendUnitBLL.GetDataBySendID(DataValid.GetNullOrInt(DrpSendUnit.SelectedValue.ToString()).Value) : null; ;

            #endregion

            testSheet.V600 = list.Contains("V600") ? true : false;
            testSheet.V600E = list.Contains("V600E") ? true : false;
           
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

            TxtSampleID.Text = (testSheet.SampleID != null) ? testSheet.SampleID.ToString() : string.Empty;  // 标本ID

            TxtSampleCode.Text = (testSheet.SampleCode != null) ? testSheet.SampleCode.ToString() : string.Empty;  // 标本代码

            TxtPatientName.Text = (testSheet.PatientName != null) ? testSheet.PatientName.ToString() : string.Empty;  // 姓名

            DrpGender.SelectedValue = (testSheet.Gender != null) ? testSheet.Gender.ID : 0;  // 性别

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
            DrpMethodName.SelectedValue = testSheet.TestMethod != null ? testSheet.TestMethod.MethodID.ToString() : "0";      // 测试方法
            DrpSendUnit.SelectedValue = testSheet.SendUnit != null ? testSheet.SendUnit.SendID.ToString() : "0";              // 送检部门

            DrpSendDoctor.SelectedValue = testSheet.SendDoctor != null ? testSheet.SendDoctor.DoctorID.ToString() : "0";      // 送检医生
            DrpTestDoctor.SelectedValue = testSheet.TestDoctor != null ? testSheet.TestDoctor.DoctorID.ToString() : "0";      // 检测医师
            DrpCheckDoctor.SelectedValue = testSheet.CheckDoctor != null ? testSheet.CheckDoctor.DoctorID.ToString() : "0";   // 审核医师
            DrpReportUnit.SelectedValue = testSheet.ReportUnit != null ? testSheet.ReportUnit.UnitID.ToString() : "0";        // 报告单位

            #region
            if (Convert.ToBoolean(testSheet.V600))
            {
                ClickState("V600");
                list.Add("V600");
            }
            if (Convert.ToBoolean(testSheet.V600E))
            {
                ClickState("V600E");
                list.Add("V600E");
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
                             btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Probe2A;
                        }
                        else
                        {
                             btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Item2A;
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

            string path = Application.StartupPath + @"\ChipPic\B_RAF";

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string picName = @"\ChipPic\B_RAF\" + DateTime.Now.ToString("yyyyMMdd") + "_" + Convert.ToString(testSheet.SampleID) + ".JPG";
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
                Test_B_rafBLL.Insert(testSheet); // 调用“业务逻辑层”的方法，检查有效性后插入至数据库。
                list.Clear();
                list.Add(" ");
                FormSysMessage.ShowSuccessMsg("“化验单据”添加成功。");
            }
             if (flag == 1)
            {
                flag = 0;
                if (Test_B_rafBLL.Update(testSheet) > 0) // 调用“业务逻辑层”的方法，检查有效性后插入至数据库。
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
            TxtSampleID.Text = "";
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
                        btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Probe2B;
                        if (control is Button && control.Text == btn.Text)
                        {
                            control.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Probe2B;
                        }
                    }
                }
                else
                {
                    btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Item2B;
                }
                listTemp.Remove(btnValue);
            }
            else
            {
                if (Regex.IsMatch(btnValue, @"[\u4e00-\u9fa5]+$"))
                {
                    foreach (Control control in this.pnlChip.Controls)
                    {
                        btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Probe2A;
                        if (control is Button && control.Text == btn.Text)
                        {
                            control.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Probe2A;
                        }
                    }
                }
                else
                {
                    btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Item2A;
                }
                listTemp.Add(btnValue);
            }
                return listTemp;
        }

        private string ResultStr(String[] arry)
        {
            string result = "";

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
                            result += items[j].JudgeDescribe + ",";
                        }
                    }
                }
            }
            return result;
        }

        private void txt_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)(sender);
            list = JudgeResult(list, btn);

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

            if (TextResult.Length > 0)
            {
                txtTestResult.Text = TextResult.Substring(0, TextResult.Length - 1);
            }
            else
            {
                txtTestResult.Text = "";
            }
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
                    
                    int width=((Y + 1) / 2 * 85 + ((Y + 1) / 2 + 1) * 2);
                    int height=(X * 33 + (X + 1) * 2);

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
                                btn.Size = new Size(85, 33);
                                btn.FlatStyle = FlatStyle.Flat;
                                btn.Font = new Font(btn.Font.FontFamily, 9, btn.Font.Style | FontStyle.Bold);
                                btn.FlatAppearance.BorderSize = 0;
                                btn.TextAlign = ContentAlignment.MiddleRight;
                                btn.Click += new System.EventHandler(this.txt_Click);
                                btn.Location = new Point(8 + j / 2 * 85, 8 + i * 33);
                                btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Item2B;

                                if (listItem.Count == 1)
                                {
                                    TempletItem templetItem = listItem[0];
                                    if(!string.IsNullOrEmpty(listItem[0].JudgeContent))
                                    {
                                        btn.Text = listItem[0].JudgeContent;
                                        if(Regex.IsMatch(listItem[0].JudgeContent, @"[\u4e00-\u9fa5]+$"))
                                        {
                                            btn.Text = listItem[0].JudgeContent;
                                            btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Probe2B;
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
                                                btn.BackgroundImage = MyCompany.ReportSystem.UIL.Properties.Resources.Probe2B;

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
            FormSearch searcher = new FormSearch(judgeId,"B_RAF");
            searcher.ShowDialog(this);
            if(!string.IsNullOrEmpty(searcher.updateID))
            {
                Reset();
                testSheet = Test_B_rafBLL.GetDataBySheetID(Convert.ToInt32(searcher.updateID));
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
                    instance.testSheet = Test_B_rafBLL.GetDataBySheetID(Convert.ToInt32(updateID));
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
                     Test_B_rafBLL.Delete(deleteIDs); // 调用“业务逻辑层”的方法，删除关联对象并更新至数据库。
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
            wcam.Start();
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

        private void PrintSheet(Test_B_raf testSheet)
        {
            Report report = new Report();

            string fileUrl = Application.StartupPath + @"\Model\B_RAF.dot";
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
                report.InsertPicture("ChipPic", picPath, 170, 45);
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
            string path = Application.StartupPath + @"\Document\B_RAF\" + DateStr;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string documentName = path + @"\" + DateStr + "B_RAF_" + Convert.ToString(testSheet.SampleID) + ".doc";
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
                    Test_B_raf testSheet = Test_B_rafBLL.GetDataBySheetID(myid);
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
                condition = " [JudgeID]=" + judgeId + " AND [TestDate]=#" + datetime + "#" + " Order by TestDate DESC, SampleID Asc "; 
            }
            BindDataGrid(condition);
        }

        private static void BindDataGrid(string conditions)
        {
            instance.PageBar.DataControl = instance.DgvGrid;
            instance.PageBar.DataSource = Test_B_rafBLL.GetPageList(instance.PageBar.PageSize, instance.PageBar.CurPage, conditions);
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
                instance.testSheet = Test_B_rafBLL.GetDataBySheetID(Convert.ToInt32(objID));
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
        //    FormPrint crReport = new FormPrint("Test_B_RAF",judgeId, ids);
        //    crReport.ShowDialog();
        //}
    }
}
