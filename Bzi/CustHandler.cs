using System;
using System.IO;
using System.Web;
using System.Web.Script.Serialization;
using System.Net;
using System.Text;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Aspose.Cells;

namespace DevFiveApron.Bzi
{
    using QF.BPMN.Business;
    using QF.BPMN.Web.Common;
    using QF.BPMN.Web.EntityDataModel;

    /// <summary>
    /// 自定义对象
    /// </summary>
    public class CustHandler : OABusinessRule
    {
        #region 常量、构造函数
        private const string _baseURL = "https://api.zkcserv.com/";
        private const string _client_id = "F1A785B1-86C3-4B19-8D1C-C4395AABB9C2";
        private const string _client_secret = "580C040E-7949-4145-9DF7-42CD17563E83";
        private const string _cmp_id = "17125";
        private static bool _Permission = true;//许可
        private static int _InterVal = 7;//使用月份，2020-01-01开始计算

        public CustHandler() { }
        #endregion

        #region 清楚缓存
        /// <summary>
        /// 清楚缓存
        /// </summary>
        /// <param name="objContext"></param>
        /// <param name="pParams"></param>
        /// <returns></returns>
        public static string KQ_Clear(Entities objContext, string pParams)
        {
            string strPer = Validation_Permission(objContext);
            if (strPer != string.Empty)
                return strPer;

            objContext.ExecuteStoreCommand("TRUNCATE TABLE OA_KQ_Trans");
            return "清除缓存成功。";
        }
        #endregion

        #region 状态设置
        public static int GetKQStatus(Entities objContext, string pParms)
        {
            string strSQL = "SELECT LValue FROM DM_Status WHERE PID = 1";
            return objContext.ExecuteStoreCommand(strSQL);
        }

        public static void UpdateKQStatus(Entities objContext, string pValue)
        {
            string strSQL = string.Format("UPDATE DM_Status SET LValue = {0}, ModiDate = GETDATE() WHERE PID = 1", pValue);
            objContext.ExecuteStoreCommand(strSQL);
        }
        #endregion

        #region 同步考勤记录
        /// <summary>
        /// 上传到OA_KQ_Trans
        /// </summary>
        /// <param name="objContext"></param>
        /// <param name="pParams"></param>
        /// <returns></returns>
        public static string KQ_Syn(Entities objContext, string pParams)
        {
            string strPer = Validation_Permission(objContext);
            if (strPer != string.Empty)
                return strPer;

            string token = string.Empty, strJson = string.Empty, strSQL;
            DateTime dYearMonth, dVar;
            DataTable dt = new DataTable();

            if (pParams == string.Empty)
                return "请设置同步日期";

            #region 获取token
            try
            {
                dYearMonth = DateTime.Parse(pParams);
                token = PostFunction(_baseURL + "get_token/?client_id=" + _client_id + "&client_secret=" + _client_secret);
                if (token.IndexOf("token") > 0)
                    token = token.Substring(token.IndexOf("token") + 17, token.IndexOf("expire_in") - (token.IndexOf("token") + 17) - 4);
                else
                    return "获取token失败。";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            #endregion

            #region 获取云数据、并同步
            try
            {
                int iDays = DateTime.DaysInMonth(dYearMonth.Year, dYearMonth.Month);
                //获取云数据
                for (int i = 0; i < iDays; i++)
                {
                    dVar = dYearMonth.AddDays(i);
                    if (dVar.CompareTo(DateTime.Now) >= 0)
                        break;

                    strJson = PostFunction(_baseURL + "get_card_record/?token=" + token + "&cmp_id=" + _cmp_id + "&start_date=" + dVar.ToString("yyyy-MM-dd") + "&end_date=" + dVar.ToString("yyyy-MM-dd"));
                    dt = JsonToTable(strJson);
                    if (dt == null || dt.Rows.Count == 0)
                        continue;

                    strSQL = "INSERT INTO OA_KQ_Trans(Emp,fstatus,staff_no,staff_name,dept_name,position_name,[date],[time],[type],[sn],[location]) VALUES";
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        if (j > 0)
                            strSQL += ",";
                        strSQL += "(0,0,'" + dt.Rows[j]["staff_no"].ToString() + "','" + dt.Rows[j]["staff_name"].ToString() + "','" + dt.Rows[j]["dept_name"].ToString() + "','" + dt.Rows[j]["position_name"].ToString() + "','" + dt.Rows[j]["date"].ToString() + "','" + dt.Rows[j]["time"].ToString() + "','" + dt.Rows[j]["type"].ToString() + "','" + dt.Rows[j]["sn"].ToString() + "','" + dt.Rows[j]["location"].ToString() + "')";
                    }
                    strSQL += ";";

                    objContext.ExecuteStoreCommand(strSQL);
                }

                //同步
                objContext.ExecuteStoreCommand("DM_P_SynKQ");
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            #endregion

            return "上传成功";
        }
        #endregion

        #region 导出考勤报表
        /// <summary>
        /// Report
        /// </summary>
        /// <param name="objContext"></param>
        /// <param name="pParams"></param>
        /// <returns></returns>
        public static string KQ_Export(Entities objContext, string pParams)
        {
            string strPer = Validation_Permission(objContext);
            if (strPer != string.Empty)
                return strPer;

            try
            {
                DataTable dt, dt2;

                string strProcedure, strProcedure2;
                DateTime dtParm = Convert.ToDateTime(pParams);
                DateTime dtNow = DateTime.Now;
                int Month = (dtNow.Year - dtParm.Year) * 12 + (dtNow.Month - dtParm.Month);
                if (Month <= 0)
                {
                    strProcedure = "DM_P_KQReport 0";
                    strProcedure2 = "DM_P_KQReportS2 0";
                }
                else
                {
                    strProcedure = "DM_P_KQReport -" + Month;
                    strProcedure2 = "DM_P_KQReportS2 -" + Month;
                }

                dt = DBFactory.GetDataTable(objContext, strProcedure, null, null);
                dt2 = DBFactory.GetDataTable(objContext, strProcedure2, null, null);

                return DataTableToExcel(dt, dt2, pParams.Substring(0, 4) + "年" + pParams.Substring(5, 2) + "月");
            }
            catch (Exception ex)
            {
                return "Error:" + ex.Message;
            }
        }
        #endregion

        #region 获取云数据
        /// <summary>
        /// 根据API获取数据
        /// </summary>
        /// <param name="pJson"></param>
        /// <returns></returns>
        public static string PostFunction(string pJson)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(pJson);
            request.Headers.Add("Access-Token", "Sinopec-Station");
            request.Method = "POST";
            request.ContentType = "application/json";
            //request.CookieContainer = cookie; //cookie信息由CookieContainer自行维护
            using (StreamWriter dataStream = new StreamWriter(request.GetRequestStream()))
            {
                dataStream.Write("json串");
                dataStream.Close();
            }
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            string encoding = response.ContentEncoding;
            if (encoding == null || encoding.Length < 1)
            {
                encoding = "UTF-8";//默认编码  
            }
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding(encoding));
            return reader.ReadToEnd();

        }
        #endregion

        #region JsonToDataTable
        /// <summary>
        /// 根据Json转换成DataTable并去重
        /// </summary>
        /// <param name="strJson"></param>
        /// <returns></returns>
        private static DataTable JsonToTable(string strJson)
        {
            DataTable dt = new DataTable();
            JObject jo = (JObject)JsonConvert.DeserializeObject(strJson);
            JArray ja = (JArray)jo["data_json"];
            string json = ja.ToString();

            JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
            javaScriptSerializer.MaxJsonLength = int.MaxValue; //取得最大数值
            ArrayList arrayList = javaScriptSerializer.Deserialize<ArrayList>(json);
            if (arrayList.Count > 0)
            {
                foreach (Dictionary<string, object> dictionary in arrayList)
                {
                    if (dictionary.Keys.Count == 0)
                    {
                        return dt;
                    }
                    if (dt.Columns.Count == 0)
                    {
                        foreach (string current in dictionary.Keys)
                        {
                            dt.Columns.Add(current, dictionary[current].GetType());
                        }
                    }
                    DataRow dataRow = dt.NewRow();
                    foreach (string current in dictionary.Keys)
                    {
                        dataRow[current] = dictionary[current];
                    }
                    dt.Rows.Add(dataRow);
                }
            }

            if (dt != null)
            {
                DataView dv = new DataView(dt);
                dt = dv.ToTable("oa_kq", true, new string[] { "staff_no", "staff_name", "dept_name", "position_name", "date", "time", "type", "sn", "location" });
            }

            return dt;
        }
        #endregion

        #region 导出到Excel
        /// <summary>
        /// 导出Excel文件
        /// </summary>
        /// <param name="pKQ">数据集</param>
        /// <returns></returns>
        internal static string DataTableToExcel(DataTable pKQ, DataTable pDK, string pParams)
        {
            string pYearMonth = pParams.Substring(0, 4) + "年" + pParams.Substring(5, 2) + "月";
            if (pKQ == null || pKQ.Rows.Count <= 0)
                return "Error:没有考勤数据";
            if (pDK == null || pDK.Rows.Count == 0)
                return "Error:没有打卡数据";

            DateTime now = DateTime.Now;
            Workbook workbook;
            try
            {
                //打开模板
                string filePath = "doc\\";
                string phyPath = HttpContext.Current.Server.MapPath(filePath);

                DirectoryInfo di = new DirectoryInfo(phyPath);
                string tempPath = di.Parent.Parent.GetDirectories("WebUI")[0].FullName + "\\doc\\Template\\";
                string fileName = "考勤记录模板.xlsx";

                if (!File.Exists(tempPath + fileName))
                    return "Error:找不到导出模板";

                workbook = new Workbook(tempPath + fileName);

                //考勤表
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;
                Worksheet worksheet2 = workbook.Worksheets[1];
                worksheet2.AutoFitRows(4, pKQ.Rows.Count + 4);
                Cells cells2 = worksheet2.Cells;

                string strTitle = "广州五号停机坪商业经营管理有限公司" + pYearMonth + "考勤表";
                Cell cell = cells[0, 0];
                cell.PutValue(strTitle);
                string strTitle2 = "员工打卡记录" + pYearMonth + "报表";
                Cell cell2 = cells2[0, 0];
                cell2.PutValue(strTitle2);

                int rowNumber = pKQ.Rows.Count;
                int columnNumber = pKQ.Columns.Count;

                int rowNumber2 = pDK.Rows.Count;
                int columnNumber2 = pDK.Columns.Count;

                //遍历DataTable行
                for (int r = 0; r < rowNumber; r++)
                    for (int c = 0; c < columnNumber; c++)
                        cells[r + 4, c].PutValue(pKQ.Rows[r][c]);

                //处理星期
                string strZQ = "考勤周期:" + pYearMonth;
                cells2[1, 0].PutValue(strZQ);
                string[] strWeek = new string[] { "日", "一", "二", "三", "四", "五", "六" };
                DateTime dateRP = DateTime.Parse(pParams);
                int FirstDW = (int)dateRP.DayOfWeek;

                string[] wks = new string[DateTime.DaysInMonth(dateRP.Year, dateRP.Month)];
                for (int i = 0; i < wks.Length; i++)
                {
                    wks[i] = strWeek[(i + FirstDW) % 7];
                }

                for (int i = 0; i < wks.Length; i++)
                {
                    cells2[2, i + 4].PutValue(wks[i]);
                }

                for (int r = 0; r < rowNumber2; r++)
                    for (int c = 0; c < columnNumber2; c++)
                        cells2[r + 4, c].PutValue(pDK.Rows[r][c]);

                //保存Excel文件
                string savePath = di.Parent.Parent.GetDirectories("WebUI")[0].FullName + "\\doc\\";
                fileName = now.ToString("yyyyMMddHHmmssfff") + ".xlsx";

                if (!Directory.Exists(savePath))
                    Directory.CreateDirectory(savePath);

                workbook.Save(savePath + fileName, SaveFormat.Xlsx);

                return filePath + fileName;
            }
            catch (Exception e)
            {
                return "Error:" + e.Message;
            }
            finally
            {
                workbook = null;
            }
        }
        #endregion

        #region Set_Permission
        /// <summary>
        /// Set Permission
        /// </summary>
        /// <param name="objContext"></param>
        /// <returns></returns>
        internal static string Validation_Permission(Entities objContext)
        {
            string strReturn = "程序试用期已过，请联系开发人员。";
            if (_Permission)
                return string.Empty;

            string strProcedure = "SP_HR_YearMonth -" + _InterVal.ToString();

            try
            {
                DataTable dt = DBFactory.GetDataTable(objContext, strProcedure, null, null);
                if (dt == null || dt.Rows.Count == 0)
                    return strReturn;

                if (dt.Rows[0]["CHECKED"].ToString() != "UPSET")
                    return string.Empty;

                return strReturn;
            }
            catch
            {
                return strReturn;
            }
        }
        #endregion
    }
}