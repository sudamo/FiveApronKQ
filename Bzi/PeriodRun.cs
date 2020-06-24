using System;

namespace DevFiveApron.Bzi
{
    using QF.BPMN.Web.Contract;
    using QF.BPMN.Web.EntityDataModel;

    public class PeriodRun : IPeriodRun
    {
        /// <summary>
        /// 执行状态
        /// </summary>
        private static bool _Execute;
        /// <summary>
        /// 构造函数
        /// </summary>
        public PeriodRun()
        {
            _Execute = false;
        }

        /// <summary>
        /// 实现接口周期执行方法
        /// </summary>
        /// <param name="objContext"></param>
        /// <param name="interval">常量值：5</param>
        /// <param name="log">日志</param>
        public void Run(Entities objContext, uint interval, log4net.ILog log)
        {
            DateTime now = DateTime.Now;

            string strYMD = now.Year.ToString() + "-" + (now.Month < 10 ? "0" + now.Month.ToString() : now.Month.ToString()) + "-01";

            //int iStatus = CustHandler.GetKQStatus(objContext, "KQ_Syn");
            //if (iStatus == 1)
            //    return;

            if ((now.Hour == 10 || now.Hour == 1) && now.Minute <= interval)
            {
                ////
                //CustHandler.UpdateKQStatus(objContext, "1");
                //
                CustHandler.KQ_Syn(objContext, strYMD);
                log.Info("Execute KQ_Syn:" + now.ToString("yyyy-MM-dd HH:mm:ss"));
                ////
                //CustHandler.UpdateKQStatus(objContext, "0");
            }
        }
    }
}