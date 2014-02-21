package sandbox.task.reporting;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import ariba.base.core.Base;
import ariba.base.core.Partition;
import ariba.base.server.BaseServer;
import ariba.user.core.User;
import ariba.util.core.Fmt;
import ariba.util.core.MIME;
import ariba.util.core.StringUtil;
import ariba.util.core.SystemUtil;
import ariba.util.log.Log;
import ariba.util.net.BasicEmailClient;
import ariba.util.scheduler.ScheduledTask;
import ariba.util.scheduler.Scheduler;

import sandbox.task.TaskUtil;

/**
 *
 * This is the super class of all report tasks that exists in the system where we need common
 * actions such as : 
 * save the report at a specific place
 * send the result report by email
 * 
 * @author fhenri
 */
public abstract class AribaReport extends ScheduledTask {

    public static final String ClassName = AribaReport.class.getName();
    
    public static final String UserIDParam     = "UserId";
    public static final String DirectoryParam  = "Directory";

    public static final String ReportDir       = 
        "config/variants/Plain/partitions/None/data/UserReport";

    protected List<String> m_userId;
    protected String m_directory;
    protected String m_filePattern;
    protected boolean m_zipFile;
    protected String m_emailTitle;

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

        m_userId = 
            TaskUtil.getListParameter(UserIDParam, parameters, false);
        
        // where to save report
        m_directory =
            TaskUtil.getStringParameter(ClassName, DirectoryParam, parameters, false);
        if (StringUtil.nullOrEmptyOrBlankString(m_directory)) {
            m_directory = SystemUtil.getSharedTempDirectory().getAbsolutePath();
        }

        m_zipFile = true;
        m_emailTitle = "Ariba Report";
    }
    
    /**
     * Send an email with a file name
     * 
     * @param fileName
     */
    protected void sendMail (String fileName)
    {
        // send email
        Log.customer.debug("send email with file %s to %s", fileName, m_userId);
        try {
            for (int i=0; i<m_userId.size(); i++) {
                String userName = (String) m_userId.get(i);
                User user = User.getUser(userName, "PasswordAdapter1");
                
                Log.customer.debug("Get user %s for username %s", user, userName);
                // skip inactive or locked users
                if (user == null  || !user.getActive() || user.getIsLocked().booleanValue()) {
                    continue;
                }
                String to = user.getEmailAddress();
                Log.customer.debug("send email to %s", to);
                if (!StringUtil.nullOrEmptyOrBlankString(to)) {
                    sendXl(fileName, to);
                }
            }
        } catch (MessagingException me) {
            Log.customer.error(10002, me.toString());
            me.printStackTrace();
        }
    }

    /**
     * @return the name of the file to use with the right date -
     */
    protected String getFileName () {

        SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
        Date today = Calendar.getInstance().getTime();

        return Fmt.S(m_filePattern, m_directory, formatter.format(today));
    }

    /**
     * Send the xl file by mail
     * 
     * @param reportName
     * @throws MessagingException 
     */
    protected void sendXl (String reportName, String to) throws MessagingException {
        
        BasicEmailClient emailClient = BaseServer.baseServer().server.emailClient();

        MimeMessage message = new MimeMessage(emailClient.getDefaultSession());

        String fromAddress = Base.getService().getParameter(
                Partition.None,
                User.ParameterAribaSystemUserEmailAddress);
        if (StringUtil.nullOrEmptyOrBlankString(fromAddress)) {
            User from = User.getAribaSystemUser(Partition.None);
            if (from == null) {
                fromAddress = "powersource@alstom.com";
            } else {
                fromAddress = from.getEmailAddress();
            }
        }

        message.setFrom(new InternetAddress(fromAddress));
        message.setHeader(MIME.HeaderContentEncoding, MIME.CharSetUTF);
        message.setHeader(
            MIME.HeaderContentTransferEncoding,
            MIME.EncodingQuotedPrintable);

        message.setRecipient(
                Message.RecipientType.TO,
                new InternetAddress(to));

        message.setHeader(
                MIME.HeaderContentType,
                MIME.ContentTypeMultipartMixed);

        message.setSubject(m_emailTitle);

        // create and fill the first message part
        MimeBodyPart mbpPart1 = new MimeBodyPart();
        mbpPart1.setText("Attached requested file: " + reportName);

        // create the Multipart and add its parts to it
        Multipart mp = new MimeMultipart();
        mp.addBodyPart(mbpPart1);

        if (!StringUtil.nullOrEmptyOrBlankString(reportName)) {
            // create the second message part with attachment
            MimeBodyPart mbpPart2 = new MimeBodyPart();
            // attach the file to the message
            File attachmentFile = new File(reportName);
            if (m_zipFile) {
                attachmentFile = zippingFile(new File(reportName));
            }
            FileDataSource fds = new FileDataSource(attachmentFile);
            mbpPart2.setDataHandler(new DataHandler(fds));
            mbpPart2.setFileName(fds.getName());
            mp.addBodyPart(mbpPart2);
        }

        // add the Multipart to the message and send it
        message.setContent(mp);
        emailClient.send(message);
    }
    
    /**
     * Zip a file
     * @param inputFile file to zip
     * @return the file zipped
     */
    public static File zippingFile(File inputFile) {
        ZipOutputStream out = null;
        InputStream in      = null;
        try {
            File outputFile = new File(inputFile.getName()+ ".zip");
            
            OutputStream rawOut = new BufferedOutputStream(new FileOutputStream(outputFile));
            out = new ZipOutputStream(rawOut);
            // optional - manages amount of compression
            // out.setLevel(java.util.zip.Deflator.BEST_COMPRESSION);
            
            InputStream rawIn = new FileInputStream(inputFile);
            in = new BufferedInputStream(rawIn);
            
             // entry for our file
            ZipEntry entry = new ZipEntry(inputFile.getName());
            // notify output stream of entry.
            out.putNextEntry(entry);
            
            // pump data from file into zip file
            byte[] buf = new byte[1024];
            int len;
            while ((len = in.read(buf)) > 0) {
                out.write(buf, 0, len);
            }
            return outputFile;
        }
        catch(IOException ioe) {
            Log.customer.error(10002, ioe.toString());
            ioe.printStackTrace();
        }
        finally {
            try {
                if(in != null) {
                    in.close();
                }
                if(out != null) {
                    out.close();
                }
            }
            catch(IOException ioe) { 
                Log.customer.error(10002, ioe.toString());
            }        
        }
        return null;
    }

}
