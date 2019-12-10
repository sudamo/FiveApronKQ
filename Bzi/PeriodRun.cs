using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DevFiveApron.Bzi
{
    using QF.BPMN.Web.Contract;
    using QF.BPMN.Web.EntityDataModel;
    using log4net;

    public class PeriodRun: IPeriodRun
    {

        public PeriodRun()
        {

        }

        public void Run(Entities objContext, uint interval, ILog log)
        {

        }
    }
}