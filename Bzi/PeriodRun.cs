using System;

namespace DevFiveApron.Bzi
{
    using QF.BPMN.Web.Contract;
    using QF.BPMN.Web.EntityDataModel;

    public class PeriodRun : IPeriodRun
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public PeriodRun() { }

        /// <summary>
        /// 实现接口周期执行方法
        /// </summary>
        /// <param name="objContext"></param>
        /// <param name="interval">常量值：5</param>
        /// <param name="log">日志</param>
        public void Run(Entities objContext, uint interval, log4net.ILog log)
        {
            DateTime now = DateTime.Now;
            if ((now.Hour == 10 || now.Hour == 1) && now.Minute <= interval)
            {
                CustHandler.KQ_Syn(objContext, now.ToString("yyyy-MM-dd"));
                log.Info("Execute KQ_Syn:" + now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
        }
    }
}