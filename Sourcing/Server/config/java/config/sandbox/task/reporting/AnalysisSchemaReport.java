package config.sandbox.task.reporting;

import java.util.Map;

import ariba.util.log.Log;
import ariba.util.scheduler.ScheduledTaskException;
import ariba.util.scheduler.Scheduler;

import config.sandbox.excel.analysis.AnalysisSchemaReference;
import config.sandbox.task.TaskUtil;

/**
 * Run a report to document the Analysis schema.
 * 
 * @author fhenri
 */
public class AnalysisSchemaReport extends AribaReport {

    public static final String ClassName = AnalysisSchemaReport.class.getName();
    
    private static final String FilePattern = "%s/%s_AnalysisSchemaReport.xls";
    
    public static final String ReferenceUserParam = "ReferenceUser";
    private String m_referenceUser;
    
    /**
     * Initializes the task
     * 
     * @param scheduler
     * @param scheduledTaskName
     * @param parameters
     * @see ariba.util.scheduler.ScheduledTask#init(ariba.util.scheduler.Scheduler, String, java.util.Map)
     */
    @SuppressWarnings("unchecked")
    public void init (Scheduler scheduler,
            String scheduledTaskName,
            Map parameters)
    {
        super.init(scheduler, scheduledTaskName, parameters);
        m_referenceUser = 
            TaskUtil.getStringParameter(scheduledTaskName, ReferenceUserParam, parameters, false);
        m_filePattern = FilePattern;
        m_emailTitle = "Ariba Analysis Schema";

    }
    
    /**
     * Run the task
     * 
     * @see ariba.util.scheduler.ScheduledTask#run()
     */
    public void run () throws ScheduledTaskException {
        Log.customer.debug("Generate Analysis Schema Report");
        
        AnalysisSchemaReference ref = new AnalysisSchemaReference(getFileName(), m_referenceUser);
        ref.execute();
        
        sendMail(getFileName());
    }

}
