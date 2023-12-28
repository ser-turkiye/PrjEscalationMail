package ser;

import java.util.List;

public class Conf {

    public static class MailTemplates{
        public static final String Project = "PROCESS_ESCALATION_MAIL";
    }

    public static class WBInboxMailSheetIndex {
        public static final Integer Mail = 0;
    }
    public static class WBInboxMailRowGroups {
        public static final List<Integer> MailHideCols = List.of(0);
        public static final Integer MailColInx = 0;
    }
    public static class PrjEscalationMailPaths {
        public static final String MainPath = "C:/tmp2/bulk/escalation-mails";
        public static final String WebBase = "http://localhost/webcube/";
    }
    public static class MainWFUpdateSheetIndex {
        public static final Integer Mail = 0;
    }
    public static class Databases{
        public static final String BPM = "BPM";
        public static final String Company = "D_QCON";
        public static final String ProjectFolder = "PRJ_FOLDER";
        public static final String ProjectWorkspace = "PRJ_FOLDER";
    }
    public static class ClassIDs{
        public static final String Template = "b9cf43d1-a4d3-482f-9806-44ae64c6139d";
        public static final String Project = "32e74338-d268-484d-99b0-f90187240549";
        public static final String Transmittal = "8bf0a09b-b569-4aef-984b-78cf1644ca19";
        public static final String SubReview = "629a28c4-6c36-44d0-90f7-1e5802f038e8";
        public static final String ReviewMain = "69d42aaf-6978-4b5a-8178-88a78f4b3158";
        public static final String ProjectWorkspace = "32e74338-d268-484d-99b0-f90187240549";
    }
    public static class Descriptors{
        public static final String ProjectNo = "ccmPRJCard_code";
        public static final String DrtnReviewer = "ccmPRJCard_ReviewerDrtn";
        public static final String DrtnConsalidator = "ccmPRJCard_ConsalidatorDrtn";
        public static final String DrtnDCC = "ccmPRJCard_DCCDrtn";
        public static final String ProjectMngr = "ccmPRJCard_prjmngr";
        public static final String EngMngr = "ccmPRJCard_EngMng";
        public static final String DCCList = "ccmPrjCard_DccList";
        public static final String Name = "ObjectName";
        public static final String TemplateName = "ObjectNumberExternal";

    }
    public static class DescriptorLiterals{
        public static final String PrjCardCode = "CCMPRJCARD_CODE";
        public static final String ObjectNumberExternal = "OBJECTNUMBER2";
    }
}
