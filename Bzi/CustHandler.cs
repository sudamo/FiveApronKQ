using System;
using System.IO;
using System.Net;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace DevFiveApron.Bzi
{
    using QF.BPMN.Business;
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

        public CustHandler() { }
        #endregion

        public static string KQ_Test(Entities objContext, string pParams)
        {
            SqlParameter[] parms = new SqlParameter[]
            {
                new SqlParameter("@ID", DbType.Int32)
            };
            parms[0].Value = 1;

            try
            {
                objContext.ExecuteStoreCommand("UPDATE Sys_User SET PwdChangeDate = GETDATE() WHERE ID = @ID", parms);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return "TestDll 测试成功[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "]";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="objContext"></param>
        /// <param name="pParams"></param>
        /// <returns></returns>
        public static string KQ_Clear(Entities objContext, string pParams)
        {
            objContext.ExecuteStoreCommand("TRUNCATE TABLE OA_KQ_Trans");
            return "清除缓存成功。";
        }

        /// <summary>
        /// 上传到OA_KQ_Trans
        /// </summary>
        /// <param name="objContext"></param>
        /// <param name="pParams"></param>
        /// <returns></returns>
        public static string KQ_Syn(Entities objContext, string pParams)
        {
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
                //获取云数据
                for (int i = 0; i < DateTime.DaysInMonth(dYearMonth.Year, dYearMonth.Month); i++)
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

            #region 同步数据
            //try
            //{
            //    strSQL = @"BEGIN TRANSACTION
            // --根据员工姓名更新Emp,time,BanChi
            //    UPDATE A SET Emp = ISNULL(B.EmpID,0),A.[time] = DATEADD(DD,DATEDIFF(DAY,A.[time],A.[date]),A.[time])
            //        ,A.BanChi = (SELECT MAX(ISNULL(C.RoleID,376)) FROM Sys_UserRole C WHERE B.ID = C.UserId AND C.RoleID IN(SELECT D.ID FROM Sys_Role D WHERE C.RoleID = D.ID AND D.GroupID = 7 AND D.Forbidden = 0))	--可以提下到OA_AttendanceRecord_TR设置班次
            //    FROM OA_KQ_Trans A
            //    INNER JOIN Sys_User B ON A.staff_name = B.Name AND B.IsLock = 0 AND B.Forbidden = 0
            //    WHERE Emp = 0;
            // --根据OA_KQ_Trans（Emp,date）更新考勤云记录
            //    MERGE INTO OA_AttendanceRecord_TR T
            //    USING
            //    (
            //    SELECT Emp EID,BanChi,[date] OnDate,MIN([time]) OnTime,MAX([time]) OffTime,COUNT(*) RCount
            //    FROM OA_KQ_Trans
            //    WHERE fstatus = 0 AND Emp <> 0
            //    GROUP BY emp,BanChi,[date]
            //    ) AS O ON O.EID = T.EID AND O.OnDate = T.OnDate
            //    WHEN MATCHED
            //     THEN UPDATE SET OnTime = O.OnTime,OffTime = O.OffTime,RCount = O.RCount
            //    WHEN NOT MATCHED
            //        THEN INSERT(EID,OnDate,OnTime,OffTime,RCount,BanChi,FStatus)
            //     VALUES(O.EID,O.OnDate,O.OnTime,O.OffTime,O.RCount,O.BanChi,0);
            // --设置OA_KQ_Trans状态fstatus = 1标识新数据库已经同步到考勤云记录OA_AttendanceRecord_TR
            //    UPDATE OA_KQ_Trans SET fstatus = 1 WHERE fstatus = 0;
            // --根据班次设置考勤云记录班次字段值(OnTimeBC,OffTimeBC)
            // UPDATE A SET OnTimeBC = CASE B.[Name] WHEN NULL THEN NULL ELSE CONVERT(DATETIME,CONVERT(VARCHAR(4),YEAR(A.OnDate)) + ' ' + REPLACE(SUBSTRING(B.[Name],3,CHARINDEX('-',B.[Name]) - 3),'：',':'),109) END
            //    ,OffTimeBC = CASE B.[Name] WHEN NULL THEN NULL ELSE CONVERT(DATETIME,CONVERT(VARCHAR(4),YEAR(A.OnDate)) + ' ' + REPLACE(SUBSTRING(B.[Name],CHARINDEX('-',B.[Name]) + 1,LEN(B.[Name]) - CHARINDEX('-',B.[Name])),'：',':'),109) END
            // FROM OA_AttendanceRecord_TR A
            // LEFT JOIN Sys_Role B ON A.BanChi = B.ID
            // WHERE A.BanChi IS NOT NULL AND A.FStatus = 0;
            // --更新OnTime，OffTime
            // --UPDATE OA_AttendanceRecord_TR SET FDescription = CASE WHEN OnTime > OnTimeBC THEN CASE WHEN OffTime < OffTimeBC THEN '早退且迟到' ELSE '早退' END WHEN OffTime < OffTimeBC THEN '迟到' ELSE '' END
            // --WHERE FStatus = 0;
            // UPDATE OA_AttendanceRecord_TR SET OnTime = CASE WHEN OnTime = OffTime THEN CASE WHEN OnTime > OffTimeBC THEN NULL ELSE OnTime END ELSE OnTime END,
            // OffTime = CASE WHEN OnTime = OffTime THEN CASE WHEN OffTime < OnTimeBC THEN NULL ELSE OffTime END ELSE OffTime END
            // WHERE FStatus = 0;
            // --根据考勤云记录更新到考勤记录
            // MERGE INTO OA_AttendanceRecord T
            //    USING
            //    (
            //    SELECT EID,BanChi,OnDate,OnTime,OffTime,FDescription
            //    FROM OA_AttendanceRecord_TR
            //    WHERE FStatus = 0
            //    ) AS O ON O.EID = T.EID AND O.OnDate = T.OnDate
            //    WHEN MATCHED
            //     THEN UPDATE SET OnTime = O.OnTime,OffTime = O.OffTime,Note = FDescription
            //    WHEN NOT MATCHED
            //        THEN INSERT(EID,BanChi,OnDate,OnTime,OffTime,Note,OnState,OffState,OnDataSource,OffDataSource,OnPlace,OffPlace)
            //     VALUES(O.EID,O.BanChi,O.OnDate,O.OnTime,O.OffTime,O.FDescription,1,1,1,1,'','');
            // --设置OA_AttendanceRecord_TR状态FStatus = 1标识新数据库已经同步到考勤记录OA_AttendanceRecord
            //    UPDATE OA_AttendanceRecord_TR SET FStatus = 1 WHERE FStatus = 0;
            //COMMIT;";

            //    objContext.ExecuteStoreCommand(strSQL);
            //}
            //catch (Exception ex)
            //{
            //    return ex.Message;
            //}
            #endregion

            return "上传成功";
        }

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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="objContext"></param>
        /// <param name="pParams"></param>
        /// <returns></returns>
        public static string KQ_Export(Entities objContext, string pParams)
        {
            return "还未部署上系统......";
        }
    }
}