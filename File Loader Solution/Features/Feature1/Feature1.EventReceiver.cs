using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Administration;
using System.Linq;

namespace FileLoaderTimerJob.Features.Feature1
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("f9ef8ed0-ac64-4782-a005-14ce5c1704a3")]
    public class Feature1EventReceiver : SPFeatureReceiver
    {

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            var webApp = properties.Feature.Parent as SPWebApplication;
            if (webApp == null) throw new Exception("webApp");

            // undeploy the job if already registered
            var ej = from SPJobDefinition job in webApp.JobDefinitions
                     where job.Name == FileLoader.JOB_DEFINITION_NAME
                     select job;

            if (ej.Count() > 0)
                ej.First().Delete();

            // create and configure timerjob
            var schedule = new SPMinuteSchedule
                {
                    BeginSecond = 0,
                    EndSecond = 59,
                    Interval = 55,
                };
            var myJob = new FileLoader(webApp)
                {
                    Schedule = schedule,
                    IsDisabled = false
                };

            // save the timer job deployment
            myJob.Update();
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            var webApp = properties.Feature.Parent as SPWebApplication;
            if (webApp == null) throw new Exception("webApp");

            // undeploy the timerjob
            var ej = from SPJobDefinition job in webApp.JobDefinitions
                     where job.Name == FileLoader.JOB_DEFINITION_NAME
                     select job;
            if (ej.Count() > 0)
                ej.First().Delete();
        }


    }
}
