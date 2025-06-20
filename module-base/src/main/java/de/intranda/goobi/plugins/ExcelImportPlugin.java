package de.intranda.goobi.plugins;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.configuration.SubnodeConfiguration;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.reloading.FileChangedReloadingStrategy;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang3.mutable.MutableInt;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.goobi.beans.Batch;
import org.goobi.beans.Process;
import org.goobi.beans.Project;
import org.goobi.beans.Step;
import org.goobi.beans.User;
import org.goobi.beans.Usergroup;
import org.goobi.managedbeans.LoginBean;
import org.goobi.plugins.datatype.Config;
import org.goobi.plugins.datatype.DataRow;
import org.goobi.plugins.datatype.MetadataMappingObject;
import org.goobi.plugins.datatype.Metadatum;
import org.goobi.plugins.datatype.UserWrapper;
import org.goobi.production.enums.LogType;
import org.goobi.production.enums.PluginType;
import org.goobi.production.flow.statistics.hibernate.FilterHelper;
import org.goobi.production.importer.Record;
import org.goobi.production.plugin.PluginLoader;
import org.goobi.production.plugin.interfaces.IOpacPlugin;
import org.goobi.production.plugin.interfaces.IPlugin;
import org.goobi.production.plugin.interfaces.IWorkflowPlugin;
import org.primefaces.event.FileUploadEvent;
import org.primefaces.model.file.UploadedFile;

import de.intranda.goobi.plugins.massuploadutils.MassUploadedFile;
import de.intranda.goobi.plugins.massuploadutils.MassUploadedProcess;
import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.config.ConfigurationHelper;
import de.sub.goobi.forms.MassImportForm;
import de.sub.goobi.helper.BeanHelper;
import de.sub.goobi.helper.Helper;
import de.sub.goobi.helper.ScriptThreadWithoutHibernate;
import de.sub.goobi.helper.enums.StepStatus;
import de.sub.goobi.helper.exceptions.DAOException;
import de.sub.goobi.helper.exceptions.ImportPluginException;
import de.sub.goobi.helper.exceptions.SwapException;
import de.sub.goobi.persistence.managers.ProcessManager;
import de.sub.goobi.persistence.managers.ProjectManager;
import de.sub.goobi.persistence.managers.StepManager;
import de.unigoettingen.sub.search.opac.ConfigOpac;
import de.unigoettingen.sub.search.opac.ConfigOpacCatalogue;
import lombok.Data;
import lombok.extern.log4j.Log4j;
import net.xeoh.plugins.base.annotations.PluginImplementation;
import ugh.dl.DigitalDocument;
import ugh.dl.DocStruct;
import ugh.dl.DocStructType;
import ugh.dl.Fileformat;
import ugh.dl.Metadata;
import ugh.dl.Prefs;
import ugh.exceptions.MetadataTypeNotAllowedException;
import ugh.exceptions.PreferencesException;
import ugh.exceptions.TypeNotAllowedForParentException;
import ugh.exceptions.WriteException;
import ugh.fileformats.mets.MetsMods;

@PluginImplementation
@Log4j
@Data
public class ExcelImportPlugin implements IWorkflowPlugin, IPlugin {

    private String title = "intranda_workflow_excelimport";
    private PluginType type = PluginType.Workflow;
    private String gui = "/uii/plugin_workflow_excelimport.xhtml";

    private String allowedTypes;
    private String filenamePart;
    private String userFolderName;
    private String filenameSeparator;
    // private String processnamePart;
    // private String processnameSeparator;
    private List<String> stepTitles;
    private List<MassUploadedFile> uploadedFiles = new ArrayList<>();
    private User user;
    private Path tempFolder;
    private HashSet<Integer> stepIDs = new HashSet<>();
    private List<MassUploadedProcess> finishedInserts = new ArrayList<>();
    private boolean copyImagesViaGoobiScript = false;

    // excel fields
    private String workflowTitle;
    private MassImportForm form;
    private Map<String, Integer> headerOrder;
    private Path excelFile;
    private Config config;
    private Prefs prefs;
    private String importFolder;
    private String ats;
    private String volumeNumber;
    private List<DataRow> rowList;

    private Process prozessVorlage = new Process();
    private BeanHelper bHelper = new BeanHelper();
    private List<Process> templateList;
    private Process processTemplate;
    private List<String> templateNames;
    private String templateName = "";
    private List<Record> recordList;
    private String batchName;
    private List<UserWrapper> users = new ArrayList<>();
    private String userName;
    private List<String> userNames;
    private UserWrapper uw;
    private boolean manualCorrection;
    private String qaStepName;

    /**
     * Constructor
     */
    public ExcelImportPlugin() {
        log.info("Excel Import Plugin started");
        allowedTypes = ConfigPlugins.getPluginConfig(title).getString("allowed-file-extensions", "/(\\.|\\/)(gif|jpe?g|png|tiff?|jp2|pdf)$/");
        filenamePart = ConfigPlugins.getPluginConfig(title).getString("filename-part", "prefix").toLowerCase();
        userFolderName = ConfigPlugins.getPluginConfig(title).getString("user-folder-name", "mass_upload").toLowerCase();
        filenameSeparator = ConfigPlugins.getPluginConfig(title).getString("filename-separator", "_").toLowerCase();
        stepTitles = Arrays.asList(ConfigPlugins.getPluginConfig(title).getStringArray("allowed-step"));
        qaStepName = ConfigPlugins.getPluginConfig(title).getString("qaStepName");
        copyImagesViaGoobiScript = ConfigPlugins.getPluginConfig(title).getBoolean("copy-images-using-goobiscript", false);
    }

    /**
     * Sets the variable templatename and resets userNames, users and userName, to be used for template selection drop down menu
     */
    public void setTemplateName(String name) {
        this.templateName = name;
        userNames = null;
        users = null;
        userName = null;
    }

    public void setBatchName(String name) {
        this.batchName = name;
    }

    /**
     * loads usernames if list does not exist
     */
    public List<String> getUserNames() {
        if (userNames == null) {
            updateUserNameList();
        }
        return userNames;
    }

    /**
     * Handle the upload and validation of a file
     * 
     * @param event
     */
    public void uploadFile(FileUploadEvent event) {

        if (user == null) {
            LoginBean login = (LoginBean) Helper.getBeanByName("LoginForm", LoginBean.class);
            if (login != null) {
                user = login.getMyBenutzer();
            }
        }

        try {
            uploadedFiles = new ArrayList<>();
            rowList = new ArrayList<>();
            recordList = new ArrayList<>();
            if (tempFolder == null) {
                tempFolder = Paths.get(ConfigurationHelper.getInstance().getTemporaryFolder(), user.getLogin());
                if (!Files.exists(tempFolder)) {
                    if (!tempFolder.toFile().mkdirs()) {
                        throw new IOException("Upload folder for user could not be created: " + tempFolder.toAbsolutePath().toString());
                    }
                }
            }
            UploadedFile upload = event.getFile();
            saveFileTemporary(upload.getFileName(), upload.getInputStream());
            excelFile = Paths.get(uploadedFiles.get(0).getFile().getAbsolutePath());
            recordList = generateRecordsFromFile();

            DataRow headerValidation = validateColumns();
            if (headerValidation.getInvalidFields() > 0) {
                rowList.add(headerValidation);
                recordList = new ArrayList<>();
                return;
            }

            rowList = validateMetadata(recordList);
            initTemplateList();
        } catch (IOException e) {
            log.error("Error while uploading files", e);
        }

    }

    private DataRow validateColumns() {
        DataRow row = new DataRow();

        for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
            if (!headerOrder.containsKey(mmo.getIdentifier())) {
                Metadatum datum = new Metadatum();
                datum.setHeadername(mmo.getIdentifier());
                // datum.setHeadername(mmo.getHeaderName());
                datum.setValue("Column '" + mmo.getIdentifier() + "' is expected, but missing.");
                datum.setValid(false);
                datum.getErrorMessages().add("Column '" + mmo.getIdentifier() + "' is expected, but missing.");
                row.setInvalidFields(row.getInvalidFields() + 1);
                row.getContentList().add(datum);
            }
        }

        return row;
    }

    /**
     * Deletes uploaded Excel file and resets internal variables
     */
    public void cancelImport() {
        try {
            Files.delete(excelFile);
        } catch (IOException e) {
            log.error("Unable to delte file at " + excelFile.toString(), e);
        }
        excelFile = null;
        recordList = null;
        rowList = null;
    }

    public boolean templateHasQaStep() {
        setTemplateFromString();
        Step step = getStepByName(processTemplate, qaStepName);
        if (step != null) {
            return true;
        }
        return false;
    }

    /**
     * Loads a list of usernames assigned to the configured qaStep, adds message to be displayed in drop down menu if step is not part of selected
     * workflow or has no users assigned
     */
    public void updateUserNameList() {
        setTemplateFromString();
        Step step = getStepByName(processTemplate, qaStepName);
        List<Usergroup> oldGroups = new ArrayList<>();
        users = new ArrayList<>();
        if (step != null) {
            for (Usergroup ug : step.getBenutzergruppen()) {
                oldGroups.add(ug);
                for (User u : ug.getBenutzer()) {
                    if (!userExistsInList(u)) {
                        users.add(new UserWrapper(u, false));
                    }
                }
            }
            // get all current users
            for (User u : step.getBenutzer()) {
                if (!userExistsInList(u)) {
                    users.add(new UserWrapper(u, false));
                }
            }
        }

        userNames = new ArrayList<>();
        // add message to userNames to display as default choice
        if (step != null) {
            if (users.isEmpty()) {
                userNames.add(Helper.getTranslation("plugin_yerushaExcelImport_noUser"));
            } else {
                userNames.add(Helper.getTranslation("plugin_yerushaExcelImport_chooseUser"));
            }
        } else {
            userNames.add(Helper.getTranslation("plugin_yerushaExcelImport_noStep"));
        }
        for (UserWrapper u : users) {
            userNames.add(u.getUser().getNachVorname());
        }
    }

    /**
     * Gets user object from list by comparing passed string to user Names (format: lastname, firstname)
     */
    private UserWrapper getUserByName(String name) {
        UserWrapper foundUser = null;
        for (UserWrapper u : users) {
            if (name != null && name.equals(u.getUser().getNachVorname())) {
                foundUser = u;
                break;
            }
        }
        return foundUser;
    }

    /**
     * Gets step object by comparing passed String to step names
     */
    private Step getStepByName(Process process, String stepName) {
        List<Step> schritteListe = process.getSchritte();
        Step schritt = null;
        for (Step s : schritteListe) {
            if (s.getTitel().equals(stepName)) {
                schritt = s;
                break;
            }
        }
        return schritt;
    }

    /**
     * Creates processes from the records in recordList according to the parameters set in batchname, processTemplate, manualCorrection
     */
    public void startImport() {
        setTemplateFromString();
        prefs = processTemplate.getRegelsatz().getPreferences();
        // if batchname is set create a new batch to add processes to
        Batch batch = null;
        if (batchName != null && !batchName.isEmpty()) {
            batch = new Batch();
            batch.setBatchName(batchName);
        }

        for (Record record : recordList) {
            try {
                // create new process
                Process process = createProcess(processTemplate, record.getId());
                if (process == null) {
                    continue;
                }
                generateFiles(record, process);
                if (batch != null) {
                    process.setBatch(batch);
                }
                if (manualCorrection) {
                    // if manual corrections are needed assign the selected user to this step,
                    // unassign everyone else
                    UserWrapper assignedUser = getUserByName(userName);
                    if (assignedUser != null && process != null) {
                        assignUserToStep(process, qaStepName, assignedUser);
                    }
                } else {
                    // delete qa step if no manual corrections are needed
                    Step step = getStepByName(process, qaStepName);
                    if (step != null) {
                        StepManager.deleteStep(step);
                    }
                }
                ProcessManager.saveProcess(process);

                // now start the first automatic step
                for (Step s : process.getSchritte()) {
                    if (StepStatus.OPEN.equals(s.getBearbeitungsstatusEnum()) && s.isTypAutomatisch()) {
                        ScriptThreadWithoutHibernate myThread = new ScriptThreadWithoutHibernate(s);
                        myThread.start();
                    }
                }

            } catch (IOException | InterruptedException | SwapException | DAOException e) {
                log.error("Error while writing metsfile of newly created process " + record.getId(), e);
            }
        }
        if (batch != null) {
            ProcessManager.saveBatch(batch);
        }

    }

    /**
     * Assigns passed user to step with passed name in process
     * 
     * @param process
     * @param stepName
     * @param assignedUser
     * @throws DAOException
     */
    private void assignUserToStep(Process process, String stepName, UserWrapper assignedUser) throws DAOException {
        //get step by name
        Step step = getStepByName(process, stepName);
        // remove all assigned users
        for (Usergroup userGroup : step.getBenutzergruppen()) {
            StepManager.removeUsergroupFromStep(step, userGroup);
        }
        for (User u : step.getBenutzer()) {
            StepManager.removeUserFromStep(step, u);
        }
        step.setBenutzer(new ArrayList<>());
        step.setBenutzergruppen(new ArrayList<>());
        // add back only the configured user
        step.getBenutzer().add(assignedUser.getUser());
        StepManager.saveStep(step);
    }

    /**
     * Check if a user exists in our internal list to avoid duplicates
     *
     * @param u
     * @return
     */
    private boolean userExistsInList(User u) {
        for (UserWrapper userWrapper : users) {
            if (userWrapper.getUser().equals(u)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Sets processTemplate by Templatename
     */
    private void setTemplateFromString() {
        for (Process process : getTemplateList()) {
            if (process.getTitel().equals(templateName)) {
                this.processTemplate = process;
            }
        }
    }

    /**
     * initializes list if it is null
     */
    private List<Process> getTemplateList() {
        if (templateList == null) {
            initTemplateList();
        }
        return templateList;
    }

    /**
     * Gets summary of validation errors
     */
    public String getInvalidFields() {
        int rows = 0;
        int fields = 0;
        if (rowList == null || rowList.isEmpty()) {
            return "";
        }
        for (DataRow a : rowList) {
            if (a.getInvalidFields() > 0) {
                rows += 1;
                fields += a.getInvalidFields();
            }
        }
        return Helper.getTranslation("plugin_yerushaExcelImport_invalidFields", String.valueOf(fields), String.valueOf(rows));
    }

    /**
     * Gets number of invalid Fields
     */
    public int getNumberInvalidFields() {
        int fields = 0;
        if (rowList == null || rowList.isEmpty()) {
            return 0;
        }
        for (DataRow a : rowList) {
            if (a.getInvalidFields() > 0) {
                fields += a.getInvalidFields();
            }
        }
        return fields;
    }

    /**
     * 
     */
    public void sortFiles() {
        Collections.sort(uploadedFiles);
    }

    /**
     * Loads list of templates
     *
     * @return
     */
    private List<Process> initTemplateList() {
        String sql = FilterHelper.criteriaBuilder("", true, null, null, null, true, false);
        List<Process> templates = ProcessManager.getProcesses(null, sql, null);
        this.templateList = templates;
        initTemplateNames();
        return templates;
    }

    /**
     * Builds list of names of templates in templateList
     */
    private void initTemplateNames() {
        List<String> lstTemplalteNames = new ArrayList<>();
        for (Process process : this.templateList) {
            lstTemplalteNames.add(process.getTitel());
        }
        this.templateName = lstTemplalteNames.get(0);
        this.templateNames = lstTemplalteNames;
    }

    /**
     * Creates new Process from prozessVorlage with title
     *
     * @param processTemplate
     * @param title
     * @return
     * @throws DAOException
     */
    private Process createProcess(Process processTemplate, String title) throws DAOException {
        String messageIdentifier = "workflowExcelImportProcessCreation";
        Process processCopy = new Process();
        String institutionIdentifier = getInstitutionIdentifier(title);
        List<Project> projectList = ProjectManager.getAllProjects();
        Project project = null;
        for (Project p : projectList) {
            String institutionIdentifierWithBrackets = "(" + institutionIdentifier + ")";
            if (p.getTitel().contains(institutionIdentifierWithBrackets)) {
                project = p;
            }
        }
        // if no project could be found report it back
        if (project == null) {
            Helper.setFehlerMeldung(messageIdentifier,
                    Helper.getTranslation("noProjectFoundForProcess") + " [" + title + " - " + institutionIdentifier + "]", "");
            return null;
        }
        // remove non-ascii characters for the sake of TIFF header limits
        String regex = ConfigurationHelper.getInstance().getProcessTitleReplacementRegex();
        String cleanedTitle = title.replaceAll(regex, "_");
        // check if process title is already in use abort creation and set notification
        // if that is the case
        if (ProcessManager.countProcessTitle(cleanedTitle, project.getInstitution()) != 0) {
            Helper.setFehlerMeldung(messageIdentifier, Helper.getTranslation("ProcessCreationErrorTitleAllreadyInUse") + " "
                    + Helper.getTranslation("process_grid_CatalogIDDigital") + ": " + cleanedTitle, "");
            return null;
            // TODO find proper exception to use here
        }
        // set basic process properties
        processCopy.setTitel(cleanedTitle);
        processCopy.setIstTemplate(false);
        processCopy.setInAuswahllisteAnzeigen(false);
        if (project != null) {
            processCopy.setProjekt(project);
        } else {
            processCopy.setProjekt(processTemplate.getProjekt());
        }
        processCopy.setRegelsatz(processTemplate.getRegelsatz());
        processCopy.setDocket(processTemplate.getDocket());
        // copy from template
        this.bHelper.SchritteKopieren(processTemplate, processCopy);
        this.bHelper.EigenschaftenKopieren(processTemplate, processCopy);

        ProcessManager.saveProcess(processCopy);
        // add message for successful process ceation
        Helper.setMeldung(messageIdentifier,
                Helper.getTranslation("process_created") + " - " + Helper.getTranslation("process_grid_CatalogIDDigital") + ": " + cleanedTitle, "");
        writeErrorsToProcessLog(title, processCopy);
        return processCopy;
    }

    private void writeErrorsToProcessLog(String title, Process processCopy) {
        List<String> processErrors = new ArrayList<>();
        //find row for this process in rowList by title
        for (DataRow row : rowList) {
            if (row.getRowIdentifier().equals(title)) {
                //if there were any validation errors add them to the list
                if (row.getInvalidFields() > 0) {
                    for (Metadatum m : row.getContentList()) {
                        if (!m.isValid()) {
                            for (String error : m.getErrorMessages()) {
                                processErrors.add(m.getHeadername() + ": " + error);
                            }
                        }
                    }
                } else {
                    break;
                }
            }
        }
        if (!processErrors.isEmpty()) {
            // ProcessManager.saveProcess(p);
            StringBuilder processlog = new StringBuilder("Validation issues during in depth data analysis: ").append("<br/>");
            processlog.append("<ul>");
            for (String s : processErrors) {
                processlog.append("<li>").append(s).append("</li>");
            }
            processlog.append("</ul>");
            Helper.addMessageToProcessLog(processCopy.getId(), LogType.DEBUG, processlog.toString());
        }
    }

    private String getInstitutionIdentifier(String title) {
        String institutionIdentifier = title.replaceAll("[^A-Za-z]+", "");
        return institutionIdentifier;
    }

    /**
     * Save the uploaded file temporary in the tmp-folder inside of goobi in a subfolder for the user
     * 
     * @param fileName
     * @param in
     * @throws IOException
     */
    private void saveFileTemporary(String fileName, InputStream in) throws IOException {
        if (user == null) {
            LoginBean login = (LoginBean) Helper.getBeanByName("LoginForm", LoginBean.class);
            if (login != null) {
                user = login.getMyBenutzer();
            }
        }
        if (tempFolder == null) {
            tempFolder = Paths.get(ConfigurationHelper.getInstance().getTemporaryFolder(), user.getLogin());
            if (!Files.exists(tempFolder)) {
                if (!tempFolder.toFile().mkdirs()) {
                    throw new IOException("Upload folder for user could not be created: " + tempFolder.toAbsolutePath().toString());
                }
            }
        }

        try (OutputStream out = Files.newOutputStream(tempFolder.resolve(fileName))) {
            int read = 0;
            byte[] bytes = new byte[1024];
            while ((read = in.read(bytes)) != -1) {
                out.write(bytes, 0, read);
            }
            out.flush();
            File file = new File(tempFolder.toFile(), fileName);
            MassUploadedFile muf = new MassUploadedFile(file, fileName);
            uploadedFiles.add(muf);
        } catch (IOException e) {
            log.error(e);
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    log.error(e);
                }
            }
        }
    }

    /**
     * Cancel the entire process and delete the uploaded files
     */
    public void cleanUploadFolder() {
        for (MassUploadedFile uploadedFile : uploadedFiles) {
            uploadedFile.getFile().delete();
        }
        uploadedFiles = new ArrayList<>();
        finishedInserts = new ArrayList<>();
        stepIDs = new HashSet<>();
    }

    /**
     * Iterates over Excel file and returns contents one Record per row. Also calls Methods which set headerOrder and the header name for mmo objects
     * in the config.
     *
     * @return list of Record objects
     */
    public List<Record> generateRecordsFromFile() {

        List<Record> lstRecords = new ArrayList<>();
        String idColumn = getConfig().getIdentifierHeaderName();
        headerOrder = new HashMap<>();

        try (InputStream fileInputStream = Files.newInputStream(excelFile); BOMInputStream in = new BOMInputStream(fileInputStream, false);
                Workbook wb = WorkbookFactory.create(in)) {

            Sheet sheet = wb.getSheetAt(0);
            int rowStart = sheet.getFirstRowNum();

            Iterator<Row> rowIterator = sheet.rowIterator();

            // get header and data row number from config first
            int rowIdentifier = getConfig().getRowIdentifier() - rowStart;
            int rowHeader = getConfig().getRowHeader() - rowStart;
            int rowDataStart = getConfig().getRowDataStart() - rowStart;
            int rowDataEnd = getConfig().getRowDataEnd() - rowStart;
            int rowCounter = 0;

            // find the identifier row
            Row identfierRow = null;
            while (rowCounter < rowIdentifier) {
                identfierRow = rowIterator.next();
                rowCounter++;
            }

            // read and validate the identifier row
            int numberOfCells = identfierRow.getLastCellNum();
            for (int i = 0; i < numberOfCells; i++) {
                readIdentifiers(identfierRow, i);
            }

            // find the header row
            if (rowIdentifier != rowHeader) {
                rowCounter = setHeaderNames(rowIterator, rowHeader, rowCounter);
            } else {
                // if Identifier and header are the same, copy them
                for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
                    mmo.setHeaderName(mmo.getIdentifier());
                }
            }

            // find out the first data row
            while (rowCounter < rowDataStart - 1) {
                identfierRow = rowIterator.next();
                rowCounter++;
            }

            // run through all the data rows
            while (rowIterator.hasNext() && rowCounter < rowDataEnd) {
                rowCounter = getContent(idColumn, rowIterator, rowCounter, lstRecords);
            }

        } catch (Exception e) {
            log.error(e);
        }

        return lstRecords;
    }

    /**
     * Gets the next row from the rowIterator, reads its contents, stores it in a map which is put into a Record object which is added to lstRecords.
     * 
     * @param idColumn name of the Identifier
     * @param rowIterator Iterator Object for rows in Excel file
     * @param rowCounter index of current row
     * @param lstRecords List of Record object to which will only be appended
     * @return updated rowCounter
     */
    private int getContent(String idColumn, Iterator<Row> rowIterator, int rowCounter, List<Record> lstRecords) {
        Map<Integer, String> map = new HashMap<>();
        Row row = rowIterator.next();
        rowCounter++;
        int lastColumn = row.getLastCellNum();
        // if row is empty, leave
        if (lastColumn == -1) {
            return rowCounter;
        }
        // read all cells in the current row
        for (int cellNumber = 0; cellNumber < lastColumn; cellNumber++) {
            String value = getCellContent(row, cellNumber).trim();
            map.put(cellNumber, value);

        }

        // just add the record if any column contains a value
        for (String v : map.values()) {
            if (v != null && !v.isEmpty()) {
                Record r = new Record();
                r.setId(map.get(headerOrder.get(idColumn)));
                r.setObject(map);
                lstRecords.add(r);
                break;
            }
        }
        return rowCounter;
    }

    /**
     * gets the content of the field at i in the passed row and stores it in headerOrder with key i
     * 
     * @param identfierRow row in excelfile where the column identifiers are
     * @param i column number
     */
    private void readIdentifiers(Row identfierRow, int i) {
        Cell cell = identfierRow.getCell(i);
        if (cell != null) {
            cell.setCellType(CellType.STRING);
            String value = cell.getStringCellValue().trim();
            for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
                if (mmo.getIdentifier().equals(value)) {
                    mmo.setColumnNumber(i);
                }
            }
            headerOrder.put(value, i);
        }
    }

    /**
     * Progresses to the row containing the Field names, sets them in the mmo objects in the config. Also checks how often a metadatum appears and
     * appends their identifier if there are multiple occurences. Returns updated row counter
     * 
     * @param rowIterator Object that iterates through the excel file
     * @param rowHeader number of the row containing the Metadata names
     * @param rowCounter number of current row
     * @return updated rowCounter corresponding to the state the rowIterator is in
     * 
     */
    private int setHeaderNames(Iterator<Row> rowIterator, int rowHeader, int rowCounter) {
        // if Identifier and header are on different rows find the header row and set
        // them for all metadata
        Row headerRow = null;
        while (rowCounter < rowHeader) {
            headerRow = rowIterator.next();
            rowCounter++;
        }
        // get Headernames from excel file
        for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
            // non existent Metadata are assigned column -1
            if (mmo.getColumnNumber() > -1) {
                Cell cell = headerRow.getCell(mmo.getColumnNumber());
                if (cell != null) {
                    cell.setCellType(CellType.STRING);
                    mmo.setHeaderName(cell.getStringCellValue());
                }
            }
        }
        // count occurances of headers
        Map<String, MutableInt> headerOccurence = new HashMap<>();
        for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
            if (mmo.getHeaderName() != null) {
                if (headerOccurence.containsKey(mmo.getHeaderName())) {
                    headerOccurence.get(mmo.getHeaderName()).increment();
                } else {
                    headerOccurence.put(mmo.getHeaderName(), new MutableInt(1));
                }
            }
        }
        // if Header name exists more than once append the identifier to separate them
        for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
            if (mmo.getHeaderName() != null && headerOccurence.get(mmo.getHeaderName()).intValue() > 1) {
                mmo.setHeaderName(mmo.getHeaderName() + " " + mmo.getIdentifier());
            }
        }
        return rowCounter;
    }

    /**
     * Gets the content of the Cell in row at position cn
     * 
     * @param row the current row in the excelfile
     * @param cn the number of the cell
     * @return String with value of cell in excel file in row and column cn
     */
    private String getCellContent(Row row, int cn) {
        Cell cell = row.getCell(cn, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        String value = "";
        switch (cell.getCellType()) {
            case BOOLEAN:
                value = cell.getBooleanCellValue() ? "true" : "false";
                break;
            case FORMULA:
                value = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                value = String.valueOf((int) cell.getNumericCellValue());
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            default:
                // none, error, blank
                value = "";
                break;
        }
        return value;
    }

    /**
     * Checks if Config Object has been loaded, if not loads it
     */
    public Config getConfig() {
        if (config == null) {
            config = loadConfig(workflowTitle);
        }
        return config;
    }

    /**
     * Initializes Config object
     *
     * @param workflowTitle
     * @return
     */
    private Config loadConfig(String workflowTitle) {
        XMLConfiguration xmlConfig = ConfigPlugins.getPluginConfig(getTitle());
        xmlConfig.setExpressionEngine(new XPathExpressionEngine());
        xmlConfig.setReloadingStrategy(new FileChangedReloadingStrategy());

        SubnodeConfiguration myconfig = null;
        try {
            myconfig = xmlConfig.configurationAt("//config");
        } catch (IllegalArgumentException e) {
            log.error("Unable to load config element from config file", e);
        }

        return new Config(myconfig);
    }

    /**
     * Iterate through all rows of the excel file and call method to check if all fields conform to their validation criteria.
     * 
     *
     * @param records
     * @return List of row elements in the excel file
     */
    private List<DataRow> validateMetadata(List<Record> records) {
        List<DataRow> rowlist = new ArrayList<>();
        for (Record record : records) {
            DataRow row = new DataRow();
            @SuppressWarnings("unchecked")
            Map<Integer, String> rowMap = (Map<Integer, String>) record.getObject();
            String rowIdentifier = rowMap.get(headerOrder.get(getConfig().getIdentifierHeaderName()));
            row.setRowIdentifier(rowIdentifier);
            for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
                Metadatum datum = validateMetadatum(rowMap, mmo);
                if (datum != null) {
                    row.getContentList().add(datum);
                    // count number of invalid fields in this row to display as statistic
                    if (!datum.isValid()) {
                        row.setInvalidFields(row.getInvalidFields() + 1);
                    }
                }
            }
            rowlist.add(row);
        }
        return rowlist;
    }

    /**
     * Checks if there are any Criteria and whether the excel field fulfills them
     * 
     * @param rowMap values from excel file are in this map
     * @param mmo this carries information on which metadatum is where and what criteria there are for it
     * @return
     */
    private Metadatum validateMetadatum(Map<Integer, String> rowMap, MetadataMappingObject mmo) {
        Metadatum datum = new Metadatum();
        datum.setHeadername(mmo.getHeaderName());
        String value = rowMap.get(headerOrder.get(mmo.getIdentifier()));
        if (value == null) {
            return null;
        }
        value = value.replace("¶", "<br/><br/>");
        value = value.replaceAll("\\u00A0|\\u2007|\\u202F", " ").trim();
        datum.setValue(value);
        // check if value is empty but required
        if (mmo.isRequired()) {
            if (value == null || value.isEmpty()) {
                datum.setValid(false);
                datum.getErrorMessages().add(mmo.getRequiredErrorMessage());
            }
        }
        // check if value matches the configured pattern
        if (mmo.getPattern() != null && value != null && !value.isEmpty()) {
            Pattern pattern = mmo.getPattern();
            Matcher matcher = pattern.matcher(value);
            if (!matcher.find()) {
                datum.setValid(false);
                datum.getErrorMessages().add(mmo.getPatternErrorMessage());
            }
        }
        // checks whether all parts of value are in the list of controlled contents
        if (!(mmo.getValidContent().isEmpty() || value == null || value.isEmpty())) {
            String[] valueList = value.split("; ");
            for (String v : valueList) {
                if (!mmo.getValidContent().contains(v)) {
                    datum.setValid(false);
                    datum.getErrorMessages().add(mmo.getValidContentErrorMessage());
                }
            }
        }
        // check if a configured requirement of either field having content is
        // fulfilled
        if (!mmo.getEitherHeader().isEmpty()) {
            if (rowMap.get(headerOrder.get(mmo.getEitherHeader())).isEmpty() && value.isEmpty()) {
                datum.setValid(false);
                datum.getErrorMessages().add(mmo.getEitherErrorMessage());
            }
        }
        // check if field has content despite required field not having content
        if (!mmo.getRequiredHeaders()[0].isEmpty()) {
            for (String requiredHeader : mmo.getRequiredHeaders()) {
                if (rowMap.get(headerOrder.get(requiredHeader)).isEmpty() && !value.isEmpty()) {
                    datum.setValid(false);
                    if (!datum.getErrorMessages().contains(mmo.getRequiredHeadersErrormessage())) {
                        datum.getErrorMessages().add(mmo.getRequiredHeadersErrormessage());
                    }
                }
            }
        }
        //check if field has the demanded wordcount
        if (mmo.getWordcount() != 0) {
            String[] wordArray = value.split(" ");
            if (wordArray.length < mmo.getWordcount()) {
                datum.setValid(false);
                datum.getErrorMessages().add(mmo.getWordcountErrormessage());
            }
        }
        return datum;
    }

    /**
     * Sets basic process properties and writes them to meta.xml
     * 
     * @param record
     * @param process
     * @throws IOException
     * @throws InterruptedException
     * @throws SwapException
     * @throws DAOException
     */
    public void generateFiles(Record record, Process process) throws IOException, InterruptedException, SwapException, DAOException {
        try {

            Object tempObject = record.getObject();
            @SuppressWarnings("unchecked")
            Map<Integer, String> rowMap = (Map<Integer, String>) tempObject;

            // generate a mets file
            DigitalDocument digitalDocument = null;
            Fileformat ff = null;
            DocStruct logical = null;
            DocStruct anchor = null;
            String catalogueIdentifier = null;
            if (!config.isUseOpac()) {
                ff = new MetsMods(prefs);
                digitalDocument = new DigitalDocument();
                ff.setDigitalDocument(digitalDocument);
                String publicationType = getConfig().getPublicationType();
                DocStructType logicalType = prefs.getDocStrctTypeByName(publicationType);
                logical = digitalDocument.createDocStruct(logicalType);
                digitalDocument.setLogicalDocStruct(logical);
            } else {
                try {
                    if (StringUtils.isBlank(config.getIdentifierHeaderName())) {
                        Helper.setFehlerMeldung("Cannot request catalogue, no identifier column defined");
                        return;
                    }

                    catalogueIdentifier = rowMap.get(headerOrder.get(config.getIdentifierHeaderName()));

                    String regex = ConfigurationHelper.getInstance().getProcessTitleReplacementRegex();
                    catalogueIdentifier = catalogueIdentifier.replaceAll(regex, "_");
                    if (StringUtils.isBlank(catalogueIdentifier)) {
                        return;
                    }
                    ff = getRecordFromCatalogue(catalogueIdentifier);
                    digitalDocument = ff.getDigitalDocument();
                    logical = digitalDocument.getLogicalDocStruct();
                    if (logical.getType().isAnchor()) {
                        anchor = logical;
                        logical = anchor.getAllChildren().get(0);
                    }
                } catch (ImportPluginException e) {
                    log.error(e);
                    return;
                }
            }

            DocStructType physicalType = prefs.getDocStrctTypeByName("BoundBook");
            DocStruct physical = digitalDocument.createDocStruct(physicalType);
            digitalDocument.setPhysicalDocStruct(physical);
            Metadata imagePath = new Metadata(prefs.getMetadataTypeByName("pathimagefiles"));
            imagePath.setValue("./images/");
            physical.addMetadata(imagePath);

            // add collections if configured
            String institutionIdentifier = null;
            if (rowMap.get(headerOrder.get(config.getIdentifierHeaderName())) != null) {
                institutionIdentifier = getInstitutionIdentifier(rowMap.get(headerOrder.get(config.getIdentifierHeaderName())));
            }
            String col = getConfig().getCollection();
            if (institutionIdentifier != null) {
                Metadata mdColl = new Metadata(prefs.getMetadataTypeByName("singleDigCollection"));
                mdColl.setValue(institutionIdentifier);
                logical.addMetadata(mdColl);
            } else if (StringUtils.isNotBlank(col)) {
                Metadata mdColl = new Metadata(prefs.getMetadataTypeByName("singleDigCollection"));
                mdColl.setValue(col);
                logical.addMetadata(mdColl);
            }

            // create importobject for massimport
            String gndURL = "http://d-nb.info/gnd/";
            for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
                addMetadatumToDocStruct(rowMap, logical, anchor, gndURL, mmo);
            }

            // write mets file into import folder
            process.writeMetadataFile(ff);
        } catch (WriteException | PreferencesException | MetadataTypeNotAllowedException | TypeNotAllowedForParentException e) {
            log.error("Error while writing Metadata file for process " + process.getTitel(), e);
        }
    }

    /**
     * Add metadata in rowMap to process represented by logical and anchor. Metadata is sorted according to pairs in class field headerOrder, the
     * relevant identifier is in mmo. DocLanguage and CatalogIDDigital get special treatment.
     * 
     * @param rowMap
     * @param logical
     * @param anchor
     * @param gndURL
     * @param mmo
     */
    private void addMetadatumToDocStruct(Map<Integer, String> rowMap, DocStruct logical, DocStruct anchor, String gndURL, MetadataMappingObject mmo) {

        String value = rowMap.get(headerOrder.get(mmo.getIdentifier()));
        if (StringUtils.isBlank(value)) {
            return;
        }
        value = value.replace("¶", "<br/><br/>");
        value = value.replaceAll("\\u00A0|\\u2007|\\u202F", " ").trim();

        String identifier = null;
        if (mmo.getNormdataHeaderName() != null) {
            identifier = rowMap.get(headerOrder.get(mmo.getNormdataHeaderName()));
        }
        if (StringUtils.isNotBlank(mmo.getRulesetName()) && StringUtils.isNotBlank(value)) {
            try {
                // DocLanguage needs special treatment as their might be several languages in the field that need to be separated
                if (mmo.isSplit()) {
                    String[] valueList = value.split("; ");
                    for (String language : valueList) {

                        Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));

                        md.setValue(language);
                        if (identifier != null) {
                            md.setAutorityFile("gnd", gndURL, identifier);
                        }
                        if (anchor != null && "anchor".equals(mmo.getDocType())) {
                            anchor.addMetadata(md);
                        } else {
                            logical.addMetadata(md);
                        }
                    }
                } else {
                    if ("DateOfOrigin".equals(mmo.getRulesetName()) && value.contains("/")) {
                        String[] splitDate = value.split("/");
                        Metadata mdStart = new Metadata(prefs.getMetadataTypeByName("DateStart"));
                        Metadata mdEnd = new Metadata(prefs.getMetadataTypeByName("DateEnd"));

                        mdStart.setValue(splitDate[0]);
                        mdEnd.setValue(splitDate[1]);
                        if (identifier != null) {
                            mdStart.setAutorityFile("gnd", gndURL, identifier);
                            mdEnd.setAutorityFile("gnd", gndURL, identifier);
                        }
                        if (anchor != null && "anchor".equals(mmo.getDocType())) {
                            anchor.addMetadata(mdStart);
                            anchor.addMetadata(mdEnd);
                        } else {
                            logical.addMetadata(mdStart);
                            logical.addMetadata(mdEnd);
                        }
                    }
                    Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                    // check if CatalogIDDigital has any disallowed characters, if so replace them with _
                    if ("CatalogIDDigital".equals(mmo.getRulesetName())) {
                        value = value.replace(' ', '-');
                    }
                    // all other Metadata are added here
                    md.setValue(value);
                    if (identifier != null) {
                        md.setAutorityFile("gnd", gndURL, identifier);
                    }
                    if (anchor != null && "anchor".equals(mmo.getDocType())) {
                        anchor.addMetadata(md);
                    } else {
                        logical.addMetadata(md);
                    }
                }
            } catch (MetadataTypeNotAllowedException e) {
                log.info(e);
                // Metadata is not known or not allowed
            }
        }
    }

    private Fileformat getRecordFromCatalogue(String identifier) throws ImportPluginException {
        ConfigOpacCatalogue coc = ConfigOpac.getInstance().getCatalogueByName(config.getOpacName());
        if (coc == null) {
            throw new ImportPluginException("Catalogue with name " + config.getOpacName() + " not found. Please check goobi_opac.xml");
        }
        IOpacPlugin myImportOpac = (IOpacPlugin) PluginLoader.getPluginByTitle(PluginType.Opac, coc.getOpacType());
        if (myImportOpac == null) {
            throw new ImportPluginException("Opac plugin " + coc.getOpacType() + " not found. Abort.");
        }
        Fileformat myRdf = null;
        try {
            myRdf = myImportOpac.search(config.getSearchField(), identifier, coc, prefs);
            if (myRdf == null) {
                throw new ImportPluginException("Could not import record " + identifier
                        + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
            }
        } catch (Exception e1) {
            throw new ImportPluginException("Could not import record " + identifier
                    + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
        }
        DocStruct ds = null;
        try {
            ds = myRdf.getDigitalDocument().getLogicalDocStruct();
            if (ds.getType().isAnchor()) {
                if (ds.getAllChildren() == null || ds.getAllChildren().isEmpty()) {
                    throw new ImportPluginException(
                            "Could not import record " + identifier + ". Found anchor file, but no children. Try to import the child record.");
                }
                ds = ds.getAllChildren().get(0);
            }
        } catch (PreferencesException e1) {
            throw new ImportPluginException("Could not import record " + identifier
                    + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
        }
        try {
            ats = myImportOpac.getAtstsl();

            List<? extends Metadata> sort = ds.getAllMetadataByType(prefs.getMetadataTypeByName("CurrentNoSorting"));
            if (sort != null && !sort.isEmpty()) {
                volumeNumber = sort.get(0).getValue();
            }

        } catch (Exception e) {
            ats = "";
        }
        return myRdf;
    }

    @Override
    public String toString() {
        return "ExcelImportPlugin [getType()=" + getType() + ", getTitle()=" + getTitle() + ", getGui()=" + getGui() + "]";
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        ExcelImportPlugin other = (ExcelImportPlugin) obj;
        if (gui == null) {
            if (other.gui != null) {
                return false;
            }
        } else if (!gui.equals(other.gui)) {
            return false;
        }
        if (title == null) {
            if (other.title != null) {
                return false;
            }
        } else if (!title.equals(other.title)) {
            return false;
        }
        if (type != other.type) {
            return false;
        }
        return true;
    }

    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result + ((gui == null) ? 0 : gui.hashCode());
        result = prime * result + ((title == null) ? 0 : title.hashCode());
        result = prime * result + ((type == null) ? 0 : type.hashCode());
        return result;
    }

}
