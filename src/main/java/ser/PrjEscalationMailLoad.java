package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import com.ser.blueline.bpm.IWorkbasket;
import de.ser.doxis4.agentserver.UnifiedAgent;
import de.ser.sst.shared.lang.ArrayUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.concurrent.TimeUnit;


public class PrjEscalationMailLoad extends UnifiedAgent {

    ISession ses;
    IDocumentServer srv;
    IBpmService bpm;

    JSONObject mailTemplates = new JSONObject();
    JSONObject excelConfigs = new JSONObject();
    JSONObject mailConfig = new JSONObject();
    JSONObject projects = new JSONObject();
    ProcessHelper helper;

    String execId;

    @Override
    protected Object execute() {
        if (getBpm() == null)
            return resultError("Null BPM object");

        (new File(Conf.PrjEscalationMailPaths.MainPath)).mkdirs();

        bpm = getBpm();
        ses = getSes();
        srv = ses.getDocumentServer();

        try {
            execId = UUID.randomUUID().toString();
            JSONObject scfg = Utils.getSystemConfig(ses);
            if(scfg.has("LICS.SPIRE_XLS")){
                com.spire.license.LicenseProvider.setLicenseKey(scfg.getString("LICS.SPIRE_XLS"));
            }
            helper = new ProcessHelper(ses);
            mailConfig = Utils.getMailConfig(ses, srv, "");

            IInformationObject[] prjs = getProjects(helper);
            for(IInformationObject cprj : prjs){
                executeForProject(cprj);
            }

            System.out.println("Tested.");
        } catch (Exception e) {
            //throw new RuntimeException(e);
            System.out.println("Exception       : " + e.getMessage());
            System.out.println("    Class       : " + e.getClass());
            System.out.println("    Stack-Trace : " + e.getStackTrace() );
            return resultRestart("Exception : " + e.getMessage(),10);
        }

        System.out.println("Finished");
        return resultSuccess("Ended successfully");
    }

    static IInformationObject[] getReadyTasks(ProcessHelper helper, String type, String prjn)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(type).append("'");
        builder.append(" AND ");
        builder.append("WFL_TASK_STATUS  = 2");
        builder.append(" AND ");
        builder.append(Conf.DescriptorLiterals.PrjCardCode + " = '").append(prjn).append("'");

        String whereClause = builder.toString();

        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] rtrn = helper.createQuery(new String[]{Conf.Databases.BPM} , whereClause , "", 0);
        if(rtrn == null){return ArrayUtils.toArray();}
        return rtrn;
    }
    static IInformationObject[] getProjects(ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.Project).append("'");
        String whereClause = builder.toString();
        
        IInformationObject[] rtrn = helper.createQuery(new String[]{Conf.Databases.ProjectFolder} , whereClause , "",0);
        if(rtrn == null){return ArrayUtils.toArray();}
        return rtrn;
    }
    private void executeForProject(IInformationObject project) throws Exception{
        if(!Utils.hasDescriptor(project, Conf.Descriptors.ProjectNo)){return;}

        String prjn = project.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
        if(prjn == null || prjn.isEmpty()){return;}

        System.out.println("PRJN : " + prjn);
        IDocument ptpl = getMailTplDocument(prjn);
        if(ptpl == null){return;}

        System.out.println("  ---> " + ptpl.getDisplayName());
        JSONObject ecfg = getExcelConfig(ptpl, prjn);
        if(ecfg == null){return;}

        JSONObject tsks = new JSONObject();
        Integer escReviewer = project.getDescriptorValue(Conf.Descriptors.DrtnReviewer, Integer.class);
        if(escReviewer != null){
            tsks.put("Reviewer", new JSONObject());
        }
        Integer escConsalidator = project.getDescriptorValue(Conf.Descriptors.DrtnConsalidator, Integer.class);
        if(escConsalidator != null){
            tsks.put("Consalidator", new JSONObject());
        }
        Integer escDCC = project.getDescriptorValue(Conf.Descriptors.DrtnDCC, Integer.class);
        if(escDCC != null){
            tsks.put("DCC", new JSONObject());
        }

        IInformationObject[] mrvs = getReadyTasks(helper, Conf.ClassIDs.ReviewMain, prjn);
        for(IInformationObject mrvw : mrvs){
            System.out.println("# Review-Main ------> " + mrvw.getDisplayName());
            ITask mtsk = (ITask) mrvw;
            if(mtsk.getCode() == "Step03" && tsks.has("Consalidator")){
                ((JSONObject) tsks.get("Consalidator")).put(mtsk.getID(), mtsk);
            }
            if(mtsk.getCode() == "Step04" && tsks.has("DCC")){
                ((JSONObject) tsks.get("DCC")).put(mtsk.getID(), mtsk);
            }
        }

        IInformationObject[] srvs = getReadyTasks(helper, Conf.ClassIDs.SubReview, prjn);
        for(IInformationObject srvw : srvs){
            System.out.println("# Sub-Review ------> " + srvw.getDisplayName());
            ITask stsk = (ITask) srvw;
            ((JSONObject) tsks.get("Reviewer")).put(stsk.getID(), stsk);
        }

        IInformationObject[] tmls = getReadyTasks(helper, Conf.ClassIDs.Transmittal, prjn);
        for(IInformationObject tmtl : tmls){
            System.out.println("# Transmittal ------> " + tmtl.getDisplayName());
            ITask ttsk = (ITask) tmtl;
            ((JSONObject) tsks.get("DCC")).put(ttsk.getID(), ttsk);
        }

        System.out.println(" $$$ Send-Mail " + prjn);

        if(escReviewer != null && tsks.has("Reviewer")){
            executeForEscalation(project, prjn, ecfg, ptpl,"Reviewer", (JSONObject) tsks.get("Reviewer"), escReviewer);
        }
        if(escConsalidator != null && tsks.has("Consalidator")){
            executeForEscalation(project, prjn, ecfg, ptpl,"Consalidator", (JSONObject) tsks.get("Consalidator"), escReviewer);
        }
        if(escDCC != null && tsks.has("DCC")){
            executeForEscalation(project, prjn, ecfg, ptpl,"DCC", (JSONObject) tsks.get("DCC"), escReviewer);
        }

    }
    private String[] getPrjMails(IInformationObject project, JSONObject ecfg, String znam){
        List<String> rtrn = new ArrayList<>();
        List<String> list = new ArrayList<>();
        List<Object> cvls = (ecfg.has(znam) ? (JSONArray) ecfg.get(znam) : new JSONArray()).toList();
        for(Object cval : cvls){
            String sval = (String) cval;
            if(sval.isEmpty()){continue;}

            if(sval.equals("Project Mngr.")){
                String pmng = project.getDescriptorValue(Conf.Descriptors.ProjectMngr, String.class);
                if(pmng != null && !pmng.isEmpty() && !list.contains(pmng)){
                    list.add(pmng);
                }
            }
            if(sval.equals("Engineering Manager")){
                String emng = project.getDescriptorValue(Conf.Descriptors.EngMngr, String.class);
                if(emng != null && !emng.isEmpty() && !list.contains(emng)){
                    list.add(emng);
                }
            }
            if(sval.equals("DCC List")){
                List<String> dlst = project.getDescriptorValues(Conf.Descriptors.DCCList, String.class);
                for(String dccu : dlst){
                    if(dccu != null && !dccu.isEmpty() && !list.contains(dccu)){
                        list.add(dccu);
                    }
                }
            }
        }

        for(String line : list){
            IWorkbasket lwbk = bpm.getWorkbasket(line);
            if(lwbk == null){continue;}

            String wbMail = lwbk.getNotifyEMail();
            if(wbMail == null || wbMail.isEmpty()){continue;}

            if(rtrn.contains(wbMail)){continue;}
            rtrn.add(wbMail);
        }

        return rtrn.toArray(new String[rtrn.size()]);
    }
    private void executeForEscalation(IInformationObject project, String prjn, JSONObject ecfg, IDocument tplDoc, String ekey, JSONObject tasks, Integer esc) throws Exception {
        String[] cc = getPrjMails(project, ecfg, ekey + ".Mail-CC");

        for(String pkey : tasks.keySet()){
            ITask ptsk = (ITask) tasks.get(pkey);

            if(ptsk.getReadyDate() == null){continue;}

            IWorkbasket zwb = ptsk.getCurrentWorkbasket();
            String zMail = zwb.getNotifyEMail();
            if(zMail == null || zMail.isEmpty()){continue;}


            Date tbgn = ptsk.getReadyDate(), tend = new Date();
            long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());

            long durd  = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
            double durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;

            if(durd < (esc * 1L)){continue;}

            System.out.println("!!!!!!!! Send Mail " + ptsk.getDisplayName());

            sendEscalationMail(project, prjn, ecfg, tplDoc, ekey, ptsk, zMail, cc,
                    esc*1L, durd, durh);
        }
    }
    private void sendEscalationMail(IInformationObject project, String prjn, JSONObject ecfg,
                                    IDocument tplDoc, String ekey, ITask task, String to, String[] cc,
                                    long esca, long durd, double durh) throws Exception {
        String uniqueId = UUID.randomUUID().toString();
        String mailExcelPath = Utils.exportDocument(tplDoc, Conf.PrjEscalationMailPaths.MainPath, "[" + prjn + "]" + execId + "@" + uniqueId);
        JSONObject params = new JSONObject();

        IProcessInstance proi = task.getProcessInstance();
        IInformationObject main = (proi != null ? proi.getMainInformationObject() : null);

        String name = "";
        if(main != null && Utils.hasDescriptor(main, Conf.Descriptors.Name)){
            name = main.getDescriptorValue(Conf.Descriptors.Name, String.class);
        }

        params.put("Name", name);
        params.put("Title", (main != null ? main.getDisplayName() : ""));
        params.put("Task", task.getName());
        params.put("DoxisLink", mailConfig.getString("webBase") + helper.getTaskURL(task.getID()));
        params.put("EscD", Long.toString(esca));
        params.put("DurD", Long.toString(durd));
        params.put("DurH", Double.toString(durh));

        saveEscalationExcel(mailExcelPath, ecfg.getString("SheetName"), params);
        String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath,
                Conf.PrjEscalationMailPaths.MainPath + "/" + "[" + prjn + "]" + execId + "@" + uniqueId + ".html");
        JSONObject mail = new JSONObject();

        mail.put("To", to);
        mail.put("Cc", String.join(";", cc));
        mail.put("Subject",
                "Escalation Alert. For {ProjectNo}"
                        .replace("{ProjectNo}", prjn)
        );
        mail.put("BodyHTMLFile", mailHtmlPath);

        try {
            Utils.sendHTMLMail(ses, srv, mailConfig, mail);
        } catch (Exception ex){
            System.out.println("EXCP [Send-Mail] : " + ex.getMessage());
        }
    }
    private void saveEscalationExcel(String tpth, String shtNm, JSONObject pbks) throws Exception {

        FileInputStream tist = new FileInputStream(tpth);
        XSSFWorkbook twrb = new XSSFWorkbook(tist);

        Sheet tsht = twrb.getSheet(shtNm);
        for (Row trow : tsht){
            for(Cell tcll : trow){
                if(tcll.getCellType() != CellType.STRING){continue;}
                String clvl = tcll.getRichStringCellValue().getString();
                String clvv = Utils.updateCell(clvl, pbks);
                if(!clvv.equals(clvl)){
                    tcll.setCellValue(clvv);
                }

                if(clvv.indexOf("[[") != (-1) && clvv.indexOf("]]") != (-1)
                        && clvv.indexOf("[[") < clvv.indexOf("]]")){
                    String znam = clvv.substring(clvv.indexOf("[[") + "[[".length(), clvv.indexOf("]]"));
                    if(pbks.has(znam)){
                        String zval = znam;
                        if(pbks.has(znam + ".Text")){
                            zval = pbks.getString(znam + ".Text");
                        }
                        tcll.setCellValue(zval);
                        String lurl = pbks.getString(znam);
                        if(!lurl.isEmpty()) {
                            Hyperlink link = twrb.getCreationHelper().createHyperlink(HyperlinkType.URL);
                            link.setAddress(lurl);
                            tcll.setHyperlink(link);
                        }
                    }
                }
            }
        }
        FileOutputStream tost = new FileOutputStream(tpth);
        twrb.write(tost);
        tost.close();
    }
    private IInformationObject getProject(String prjn){
        if(projects.has(prjn)){return (IInformationObject) projects.get(prjn);}
        if(projects.has("!" + prjn)){return null;}

        IInformationObject iprj = Utils.getProjectWorkspace(prjn, helper);
        if(iprj == null){
            projects.put("!" + prjn, "[[ " + prjn + " ]]");
            return null;
        }
        projects.put(prjn, iprj);
        return iprj;
    }
    private IDocument getMailTplDocument(String prjn) throws Exception{
        if(mailTemplates.has(prjn)){return (IDocument) mailTemplates.get(prjn);}
        if(mailTemplates.has("!" + prjn)){return null;}

        IInformationObject prjt = getProject(prjn);
        if(prjt == null){return null;}

        IDocument dtpl = Utils.getTemplateDocument(prjt, Conf.MailTemplates.Project);
        if(dtpl == null){
            mailTemplates.put("!" + prjn, "[[ " + prjn + " ]]");
            return null;
        }
        mailTemplates.put(prjn, dtpl);
        return dtpl;
    }
    private JSONObject getExcelConfig(IDocument template, String prjn) throws Exception {
        if(excelConfigs.has(prjn)){return (JSONObject) excelConfigs.get(prjn);}
        if(excelConfigs.has("!" + prjn)){return null;}

        String excelPath = FileEvents.fileExport(template, Conf.PrjEscalationMailPaths.MainPath, "[" + prjn + "]" + execId);
        JSONObject ecfg = (FilenameUtils.getExtension(excelPath).toString().toUpperCase().equals("XLSX") ?
                Utils.getExcelConfig(excelPath) : new JSONObject());
        if(!ecfg.has("SheetName")){
            excelConfigs.put("!" + prjn, "[[ " + prjn + " ]]");
            return null;
        }

        excelConfigs.put(prjn, ecfg);
        return ecfg;
    }

}