package de.intranda.goobi.plugins;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
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
import org.goobi.beans.Processproperty;
import org.goobi.beans.Step;
import org.goobi.beans.User;
import org.goobi.beans.Usergroup;
import org.goobi.managedbeans.LoginBean;
import org.goobi.production.enums.LogType;
import org.goobi.production.enums.PluginType;
import org.goobi.production.flow.statistics.hibernate.FilterHelper;
import org.goobi.production.importer.Record;
import org.goobi.production.plugin.PluginLoader;
import org.goobi.production.plugin.interfaces.IOpacPlugin;
import org.goobi.production.plugin.interfaces.IPlugin;
import org.goobi.production.plugin.interfaces.IValidatorPlugin;
import org.goobi.production.plugin.interfaces.IWorkflowPlugin;
import org.primefaces.event.FileUploadEvent;
import org.primefaces.model.UploadedFile;

import de.intranda.goobi.plugins.datatype.Config;
import de.intranda.goobi.plugins.datatype.DataRow;
import de.intranda.goobi.plugins.datatype.GroupMappingObject;
import de.intranda.goobi.plugins.datatype.MetadataMappingObject;
import de.intranda.goobi.plugins.datatype.Metadatum;
import de.intranda.goobi.plugins.datatype.PersonMappingObject;
import de.intranda.goobi.plugins.datatype.UserWrapper;
import de.intranda.goobi.plugins.massuploadutils.GoobiScriptCopyImages;
import de.intranda.goobi.plugins.massuploadutils.MassUploadedFile;
import de.intranda.goobi.plugins.massuploadutils.MassUploadedFileStatus;
import de.intranda.goobi.plugins.massuploadutils.MassUploadedProcess;
import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.config.ConfigurationHelper;
import de.sub.goobi.forms.MassImportForm;
import de.sub.goobi.helper.BeanHelper;
import de.sub.goobi.helper.Helper;
import de.sub.goobi.helper.HelperSchritte;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.enums.StepStatus;
import de.sub.goobi.helper.exceptions.DAOException;
import de.sub.goobi.helper.exceptions.ImportPluginException;
import de.sub.goobi.helper.exceptions.SwapException;
import de.sub.goobi.persistence.managers.ProcessManager;
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
import ugh.dl.MetadataGroup;
import ugh.dl.Person;
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

	private static final String PLUGIN_NAME = "intranda_workflow_excelimport";
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
	@SuppressWarnings("unchecked")
	public ExcelImportPlugin() {
		log.info("Mass upload plugin started");
		allowedTypes = ConfigPlugins.getPluginConfig(PLUGIN_NAME).getString("allowed-file-extensions",
				"/(\\.|\\/)(gif|jpe?g|png|tiff?|jp2|pdf)$/");
		filenamePart = ConfigPlugins.getPluginConfig(PLUGIN_NAME).getString("filename-part", "prefix").toLowerCase();
		userFolderName = ConfigPlugins.getPluginConfig(PLUGIN_NAME).getString("user-folder-name", "mass_upload")
				.toLowerCase();
		filenameSeparator = ConfigPlugins.getPluginConfig(PLUGIN_NAME).getString("filename-separator", "_")
				.toLowerCase();
		// processnamePart =
		// ConfigPlugins.getPluginConfig(this).getString("processname-part",
		// "complete").toLowerCase();
		// processnameSeparator =
		// ConfigPlugins.getPluginConfig(this).getString("processname-separator",
		// "_").toLowerCase();
		stepTitles = ConfigPlugins.getPluginConfig(PLUGIN_NAME).getList("allowed-step");
		qaStepName = ConfigPlugins.getPluginConfig(PLUGIN_NAME).getString("qaStepName");
		copyImagesViaGoobiScript = ConfigPlugins.getPluginConfig(PLUGIN_NAME)
				.getBoolean("copy-images-using-goobiscript", false);
		LoginBean login = (LoginBean) Helper.getManagedBeanValue("#{LoginForm}");
		if (login != null) {
			user = login.getMyBenutzer();
		}
	}

	@Override
	public PluginType getType() {
		return PluginType.Workflow;
	}

	@Override
	public String getTitle() {
		return PLUGIN_NAME;
	}

	@Override
	public String getGui() {
		return "/uii/plugin_workflow_massupload.xhtml";
	}

	public void setTemplateName(String name) {
		this.templateName = name;
		userNames = null;
		users = null;
		userName = null;
	}

	public void setBatchName(String name) {
		this.batchName = name;
	}

	public List<String> getUserNames() {
		if (userNames == null) {
			updateUserNameList();
		}
		return userNames;
	}

	/**
	 * Handle the upload of a file
	 * 
	 * @param event
	 */
	public void uploadFile(FileUploadEvent event) {
		try {
			uploadedFiles = new ArrayList<>();
			rowList = new ArrayList<>();
			recordList = new ArrayList<>();
			if (tempFolder == null) {
				tempFolder = Paths.get(ConfigurationHelper.getInstance().getTemporaryFolder(), user.getLogin());
				if (!Files.exists(tempFolder)) {
					if (!tempFolder.toFile().mkdirs()) {
						throw new IOException("Upload folder for user could not be created: "
								+ tempFolder.toAbsolutePath().toString());
					}
				}
			}
			UploadedFile upload = event.getFile();
			saveFileTemporary(upload.getFileName(), upload.getInputstream());
			excelFile = Paths.get(uploadedFiles.get(0).getFile().getAbsolutePath());
			recordList = generateRecordsFromFile();
			rowList = validationTest(recordList);
			initTemplateList();
		} catch (IOException e) {
			log.error("Error while uploading files", e);
		}

	}

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
		if (step != null) {
			if (users.isEmpty()) {
				userNames.add("No users assigned to step");
			} else {
				userNames.add("Choose user");
			}
		} else {
			userNames.add("No step with configured name found");
		}
		for (UserWrapper u : users) {
			userNames.add(u.getUser().getNachVorname());
		}
	}

	private UserWrapper getUserByName(String name) {
		UserWrapper foundUser = null;
		for (UserWrapper u : users) {
			if (name.equals(u.getUser().getNachVorname())) {
				foundUser = u;
				break;
			}
		}
		return foundUser;
	}

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

	public void startImport() {
		setTemplateFromString();
		prefs = processTemplate.getRegelsatz().getPreferences();
		Batch batch = null;
		if (!batchName.isEmpty()) {
			batch = new Batch();
			batch.setBatchName(batchName);
		}

		for (Record record : recordList) {
			try {
				Process process = createProcess(processTemplate, record.getId());
				generateFiles(record, process);
				if (batch != null) {
					process.setBatch(batch);
				}
				if (manualCorrection) {
					UserWrapper assignedUser = getUserByName(userName);
					if (assignedUser != null) {
						assignUserToStep(process, qaStepName, assignedUser);
					}
				} else {
					Step step = getStepByName(process, qaStepName);
					if (step != null) {
						StepManager.deleteStep(step);
					}
				}
				ProcessManager.saveProcess(process);

			} catch (DAOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				log.error("Error while writing metsfile of newly created process " + record.getId(), e);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SwapException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		if (batch != null) {
			ProcessManager.saveBatch(batch);
		}

	}

	private void assignUserToStep(Process process, String stepName, UserWrapper assignedUser) throws DAOException {
		Step step = getStepByName(process, stepName);
		for (Usergroup ug : step.getBenutzergruppen()) {
			StepManager.removeUsergroupFromStep(step, ug);
		}
		for (User u : step.getBenutzer()) {
			StepManager.removeUserFromStep(step, u);
		}
		step.setBenutzer(new ArrayList<User>());
		step.setBenutzergruppen(new ArrayList<Usergroup>());
		step.getBenutzer().add(assignedUser.getUser());
		StepManager.saveStep(step);
	}

	/**
	 * check if a user exists in our internal list to avoid duplicates
	 */
	private boolean userExistsInList(User u) {
		for (UserWrapper userWrapper : users) {
			if (userWrapper.getUser().equals(u)) {
				return true;
			}
		}
		return false;
	}

	private void setTemplateFromString() {
		for (Process process : getTemplateList()) {
			if (process.getTitel().equals(templateName)) {
				this.processTemplate = process;
			}
		}
	}

	private List<Process> getTemplateList() {
		if (templateList == null) {
			initTemplateList();
		}
		return templateList;
	}

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
		return String.valueOf(fields) + " invalid fields in " + rows + " rows";
	}

	public void sortFiles() {
		Collections.sort(uploadedFiles);
	}

	private List<Process> initTemplateList() {
		String sql = FilterHelper.criteriaBuilder("", true, null, null, null, true, false);
		List<Process> templates = ProcessManager.getProcesses(null, sql);
		this.templateList = templates;
		initTemplateNames();
		return templates;
	}

	private void initTemplateNames() {
		List<String> templateNamesInit = new ArrayList<>();
		for (Process process : this.templateList) {
			templateNamesInit.add(process.getTitel());
		}
		this.templateName = templateNamesInit.get(0);
		this.templateNames = templateNamesInit;
	}

	private Process createProcess(Process prozessVorlage, String title) throws DAOException {
		Process prozessKopie = new Process();
		prozessKopie.setTitel(title);
		prozessKopie.setIstTemplate(false);
		prozessKopie.setInAuswahllisteAnzeigen(false);
		prozessKopie.setProjekt(prozessVorlage.getProjekt());
		prozessKopie.setRegelsatz(prozessVorlage.getRegelsatz());
		prozessKopie.setDocket(prozessVorlage.getDocket());

		this.bHelper.SchritteKopieren(prozessVorlage, prozessKopie);
		this.bHelper.ScanvorlagenKopieren(prozessVorlage, prozessKopie);
		this.bHelper.WerkstueckeKopieren(prozessVorlage, prozessKopie);
		this.bHelper.EigenschaftenKopieren(prozessVorlage, prozessKopie);

		ProcessManager.saveProcess(prozessKopie);

		return prozessKopie;
	}

	/**
	 * Save the uploaded file temporary in the tmp-folder inside of goobi in a
	 * subfolder for the user
	 * 
	 * @param fileName
	 * @param in
	 * @throws IOException
	 */
	private void saveFileTemporary(String fileName, InputStream in) throws IOException {
		if (tempFolder == null) {
			tempFolder = Paths.get(ConfigurationHelper.getInstance().getTemporaryFolder(), user.getLogin());
			if (!Files.exists(tempFolder)) {
				if (!tempFolder.toFile().mkdirs()) {
					throw new IOException(
							"Upload folder for user could not be created: " + tempFolder.toAbsolutePath().toString());
				}
			}
		}

		try (OutputStream out = Files.newOutputStream(tempFolder.resolve(fileName));) {
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
	 * do not upload the images from web UI, use images of subfolder in user home
	 * directory instead, usually called 'mass_upload'
	 */
	public void readFilesFromUserHomeFolder() {
		uploadedFiles = new ArrayList<>();
		finishedInserts = new ArrayList<>();
		stepIDs = new HashSet<>();
		try {
			Path folder = Paths.get(user.getHomeDir(), userFolderName);
//			File folder = new File(user.getHomeDir(), userFolderName);
			if (Files.exists(folder) && Files.isReadable(folder)) {
				// we use the Files API intentionally, as we expect folders with many files in
				// them.
				// The nio DirectoryStream initializes the Path objects lazily, so we don't have
				// as many objects in memory and to create
				try (DirectoryStream<Path> files = Files.newDirectoryStream(folder)) {
					Map<String, List<Process>> searchCache = new HashMap<>();
					for (Path file : files) {
						if (!file.getFileName().toString().equals(".DS_Store")) {
							MassUploadedFile muf = new MassUploadedFile(file.toFile(), file.getFileName().toString());
							assignProcess(muf, searchCache);
							uploadedFiles.add(muf);
						}
					}
				}
			} else {
				Helper.setFehlerMeldung(
						"Folder " + folder.toAbsolutePath().toString() + " does not exist or is not readable.");
			}
		} catch (Exception e) {
			log.error("Error while reading files from users home directory for mass upload", e);
			Helper.setFehlerMeldung("Error while reading files from users home directory for mass upload", e);
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
	 * All uploaded files shall now be moved to the correct processes
	 */
	public void startInserting() {

		if (copyImagesViaGoobiScript) {
			GoobiScriptCopyImages gsci = new GoobiScriptCopyImages();
			gsci.setUploadedFiles(uploadedFiles);
			gsci.setUser(user);
			gsci.execute();
			Helper.setMeldung("plugin_massupload_insertionStartedViaGoobiScript");

		} else {
			for (MassUploadedFile muf : uploadedFiles) {
				if (muf.getStatus() == MassUploadedFileStatus.OK) {
					Path src = Paths.get(muf.getFile().getAbsolutePath());
					Path target = Paths.get(muf.getProcessFolder(), muf.getFilename());
					try {
						StorageProvider.getInstance().copyFile(src, target);
					} catch (IOException e) {
						muf.setStatus(MassUploadedFileStatus.ERROR);
						muf.setStatusmessage("File could not be copied to: " + target.toString());
						log.error("Error while copying file during mass upload", e);
						Helper.setFehlerMeldung("Error while copying file during mass upload", e);
					}
					muf.getFile().delete();
				} else {
					Helper.setFehlerMeldung("File could not be matched and gets skipped: " + muf.getFilename());
				}
			}

			// all images are uploaded, so we close the workflow step now
			// first remove all stepIds which had errors
			for (MassUploadedFile muf : uploadedFiles) {
				if (muf.getStatus() != MassUploadedFileStatus.OK) {
					stepIDs.remove(muf.getStepId());
				}
			}

			// all others can be finished now
			for (Integer id : stepIDs) {
				Step so = StepManager.getStepById(id);
				if (so.getValidationPlugin() != null && so.getValidationPlugin().length() > 0) {
					IValidatorPlugin ivp = (IValidatorPlugin) PluginLoader.getPluginByTitle(PluginType.Validation,
							so.getValidationPlugin());
					ivp.setStep(so);
					if (!ivp.validate()) {
						log.error("Error while closing the step " + so.getTitel() + " for process "
								+ so.getProzess().getTitel());
						Helper.setFehlerMeldung("Error while closing the step " + so.getTitel() + " for process "
								+ so.getProzess().getTitel());
					}
				}
				Helper.addMessageToProcessLog(so.getProcessId(), LogType.DEBUG,
						"Images uploaded and step " + so.getTitel() + " finished using Massupload Plugin.");
				HelperSchritte hs = new HelperSchritte();
				so.setBearbeitungsbenutzer(user);
				hs.CloseStepObjectAutomatic(so);
				finishedInserts.add(new MassUploadedProcess(so));
			}

			Helper.setMeldung("plugin_massupload_allFilesInserted");
		}
	}

	/**
	 * check for uploaded file if a correct process can be found and assigned
	 * 
	 * @param uploadedFile
	 */
	private void assignProcess(MassUploadedFile uploadedFile, Map<String, List<Process>> searchCache) {
		// get the relevant part of the file name
		String matchFile = uploadedFile.getFilename().substring(0, uploadedFile.getFilename().lastIndexOf('.'));
		if (filenamePart.equals("prefix") && matchFile.contains(filenameSeparator)) {
			matchFile = matchFile.substring(0, matchFile.lastIndexOf(filenameSeparator));
		}
		if (filenamePart.equals("suffix") && matchFile.contains(filenameSeparator)) {
			matchFile = matchFile.substring(matchFile.lastIndexOf(filenameSeparator) + 1, matchFile.length());
		}

		// get all matching processes
		// first try to get this from the cache
		String filter = FilterHelper.criteriaBuilder(matchFile, false, null, null, null, true, false);
		List<Process> hitlist = searchCache == null ? null : searchCache.get(filter);
		if (hitlist == null) {
			// there was no result in the cache. Get result from the DB and then add it to
			// the cache.
			hitlist = ProcessManager.getProcesses("prozesse.titel", filter, 0, 5);
			if (searchCache != null) {
				searchCache.put(filter, hitlist);
			}
		}

		// if list is empty
		if (hitlist == null || hitlist.isEmpty()) {
			uploadedFile.setStatus(MassUploadedFileStatus.ERROR);
			uploadedFile.setStatusmessage("No matching process found for this image.");
		} else {
			// if list is bigger then one hit
			if (hitlist.size() > 1) {
				StringBuilder processtitles = new StringBuilder();
				for (Process process : hitlist) {
					processtitles.append(process.getTitel());
					processtitles.append(", ");
				}
				uploadedFile.setStatus(MassUploadedFileStatus.ERROR);
				uploadedFile.setStatusmessage(
						"More than one matching process where found for this image: " + processtitles.toString());
			} else {
				// we have just one hit and take it
				Process p = hitlist.get(0);
				uploadedFile.setProcessId(p.getId());
				uploadedFile.setProcessTitle(p.getTitel());
				try {
					uploadedFile.setProcessFolder(p.getImagesOrigDirectory(true));
				} catch (IOException | InterruptedException | SwapException | DAOException e) {
					uploadedFile.setStatus(MassUploadedFileStatus.ERROR);
					uploadedFile.setStatusmessage("Error getting the master folder: " + e.getMessage());
				}

				if (uploadedFile.getStatus() != MassUploadedFileStatus.ERROR) {
					// check if one of the open workflow steps is named as expected
					boolean workflowStepAsExpected = false;

					for (Step s : p.getSchritte()) {
						if (s.getBearbeitungsstatusEnum() == StepStatus.OPEN) {
							for (String st : stepTitles) {
								if (st.equals(s.getTitel())) {
									workflowStepAsExpected = true;
									uploadedFile.setStepId(s.getId());
									stepIDs.add(s.getId());
								}
							}
						}
					}

					// if correct open step was found, remember it
					if (workflowStepAsExpected) {
						uploadedFile.setStatus(MassUploadedFileStatus.OK);
					} else {
						uploadedFile.setStatus(MassUploadedFileStatus.ERROR);
						uploadedFile.setStatusmessage(
								"Process could be found, but there is no open workflow step with correct naming that could be accepted.");
					}
				}
			}
		}
	}

	public List<Record> generateRecordsFromFile() {

//		if (StringUtils.isBlank(workflowTitle)) {
//			workflowTitle = form.getTemplate().getTitel();
//		}

		List<Record> lstRecords = new ArrayList<>();
		String idColumn = getConfig().getIdentifierHeaderName();
		headerOrder = new HashMap<>();
//		List<Column> columnList=new ArrayList<>();

		try (InputStream fileInputStream = Files.newInputStream(excelFile);
//				try (InputStream fileInputStream = new FileInputStream(excelFile);
				BOMInputStream in = new BOMInputStream(fileInputStream, false);
				Workbook wb = WorkbookFactory.create(in);) {

			Sheet sheet = wb.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.rowIterator();

			// get header and data row number from config first
			int rowIdentifier = getConfig().getRowIdentifier();
			int rowHeader = getConfig().getRowHeader();
			int rowDataStart = getConfig().getRowDataStart();
			int rowDataEnd = getConfig().getRowDataEnd();
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

			// find the header row
			if (rowIdentifier != rowHeader) {
				// if Identifier and header are on different rows find the header row and set
				// them for all metadata
				Row headerRow = null;
				while (rowCounter < rowHeader) {
					headerRow = rowIterator.next();
					rowCounter++;
				}
				for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
					if (mmo.getColumnNumber() > -1) {
						Cell cell = headerRow.getCell(mmo.getColumnNumber());
						if (cell != null) {
							cell.setCellType(CellType.STRING);
							mmo.setHeaderName(cell.getStringCellValue());
						}
					}
				}
				Map<String, MutableInt> headerOccurence = new HashMap<>();
				for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
					if (mmo.getHeaderName() != null) {
						if (headerOccurence.containsKey(mmo.getHeaderName())) {
							headerOccurence.get(mmo.getHeaderName()).increment();;
						} else {
							headerOccurence.put(mmo.getHeaderName(), new MutableInt(1));
						}
					}
				}
				for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
					if (mmo.getHeaderName() != null && headerOccurence.get(mmo.getHeaderName()).intValue() > 1) {
						mmo.setHeaderName(mmo.getHeaderName() +" "+ mmo.getIdentifier());
					}
				}
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
				Map<Integer, String> map = new HashMap<>();
				Row row = rowIterator.next();
				rowCounter++;
				int lastColumn = row.getLastCellNum();
				if (lastColumn == -1) {
					continue;
				}
				for (int cellNumber = 0; cellNumber < lastColumn; cellNumber++) {
					String value = getCellContent(row, cellNumber);
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
			}

		} catch (Exception e) {
			log.error(e);
		}

		return lstRecords;
	}

	private String getCellContent(Row row, int cn) {
		// while (cellIterator.hasNext()) {
		// Cell cell = cellIterator.next();
		Cell cell = row.getCell(cn, MissingCellPolicy.CREATE_NULL_AS_BLANK);
		String value = "";
		switch (cell.getCellTypeEnum()) {
		case BOOLEAN:
			value = cell.getBooleanCellValue() ? "true" : "false";
			break;
		case FORMULA:
			// value = cell.getCellFormula();
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
	 * checks if Config Object has been loaded, if not loads it
	 */
	public Config getConfig() {
		if (config == null) {
			config = loadConfig(workflowTitle);
		}
		return config;
	}

	/**
	 * initializes Config object
	 */
	private Config loadConfig(String workflowTitle) {
		XMLConfiguration xmlConfig = ConfigPlugins.getPluginConfig(getTitle());
		xmlConfig.setExpressionEngine(new XPathExpressionEngine());
		xmlConfig.setReloadingStrategy(new FileChangedReloadingStrategy());

		SubnodeConfiguration myconfig = null;
		try {

//			myconfig = xmlConfig.configurationAt("/opt/digiverso/goobi/config/plugin_intranda_import_excel_read_headerdata.xml");
			myconfig = xmlConfig.configurationAt("//config[./template = '" + workflowTitle + "']");
		} catch (IllegalArgumentException e) {
			myconfig = xmlConfig.configurationAt("//config[./template = '*']");
		}

		return new Config(myconfig);
	}

	private List<DataRow> validationTest(List<Record> records) {
		List<DataRow> rowlist = new ArrayList<>();
		for (Record record : records) {
			DataRow row = new DataRow();
			@SuppressWarnings("unchecked")
			Map<Integer, String> rowMap = (Map<Integer, String>) record.getObject();
			String rowIdentifier = rowMap.get(headerOrder.get(getConfig().getIdentifierHeaderName()));
			row.setRowIdentifier(rowIdentifier);
			for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
				Metadatum datum = new Metadatum();
				datum.setHeadername(mmo.getHeaderName());
				String value = rowMap.get(headerOrder.get(mmo.getIdentifier()));
				datum.setValue(value);
				if (mmo.isRequired()) {
					if (value == null || value.isEmpty()) {
						datum.setValid(false);
					}
				}
				if (!mmo.getPattern().isEmpty() && value != null && !value.isEmpty()) {
					Pattern pattern = Pattern.compile(mmo.getPattern());
					Matcher matcher = pattern.matcher(value);
					if (!matcher.find()) {
						datum.setValid(false);
					}
				}
				if (!(mmo.getValidContent().isEmpty() || value == null || value.isEmpty())) {
					String[] valueList = value.split("; ");
					for (String v : valueList) {
						if (!mmo.getValidContent().contains(v)) {
							datum.setValid(false);
						}
					}
				}
				if (!mmo.getEitherHeader().isEmpty()) {
					if (rowMap.get(headerOrder.get(mmo.getEitherHeader())).isEmpty() && value.isEmpty()) {
						datum.setValid(false);
					}
				}
				if (!mmo.getRequiredHeaders()[0].isEmpty()) {
					for (String requiredHeader : mmo.getRequiredHeaders()) {
						if (rowMap.get(headerOrder.get(requiredHeader)).isEmpty() && !value.isEmpty()) {
							datum.setValid(false);
						}
					}
				}
				row.getContentList().add(datum);
				if (!datum.isValid()) {
					row.setInvalidFields(row.getInvalidFields() + 1);
				}

			}
			rowlist.add(row);
		}
		return rowlist;
	}

	public void generateFiles(Record record, Process process)
			throws IOException, InterruptedException, SwapException, DAOException {

		try {

			Object tempObject = record.getObject();

			@SuppressWarnings("unchecked")
			Map<Integer, String> rowMap = (Map<Integer, String>) tempObject;

			// generate a mets file
			DigitalDocument digitalDocument = null;
			Fileformat ff = null;
			DocStruct logical = null;
			DocStruct anchor = null;
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

					String catalogueIdentifier = rowMap.get(headerOrder.get(config.getIdentifierHeaderName()));
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
			String col = getConfig().getCollection();
			if (StringUtils.isNotBlank(col)) {
				Metadata mdColl = new Metadata(prefs.getMetadataTypeByName("singleDigCollection"));
				mdColl.setValue(col);
				logical.addMetadata(mdColl);
			}
			// and add all collections that where selected
//			for (String colItem : form.getDigitalCollections()) {
//				if (!colItem.equals(col.trim())) {
//					Metadata mdColl = new Metadata(prefs.getMetadataTypeByName("singleDigCollection"));
//					mdColl.setValue(colItem);
//					logical.addMetadata(mdColl);
//				}
//			}
			// create file name for mets file

			// create importobject for massimport

			for (MetadataMappingObject mmo : getConfig().getMetadataList()) {

				String value = rowMap.get(headerOrder.get(mmo.getIdentifier()));
				String identifier = null;
				if (mmo.getNormdataHeaderName() != null) {
					identifier = rowMap.get(headerOrder.get(mmo.getNormdataHeaderName()));
				}
				if (StringUtils.isNotBlank(mmo.getRulesetName()) && StringUtils.isNotBlank(value)) {
					try {
						Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
						md.setValue(value);
						if (identifier != null) {
							md.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);

						}
						if (anchor != null && "anchor".equals(mmo.getDocType())) {
							anchor.addMetadata(md);
						} else {
							logical.addMetadata(md);
						}
					} catch (MetadataTypeNotAllowedException e) {
						log.info(e);
						// Metadata is not known or not allowed
					}
				}

				if (StringUtils.isNotBlank(mmo.getPropertyName())) {
					Processproperty p = new Processproperty();
					p.setTitel(mmo.getPropertyName());
					p.setWert(value);
				}
			}

			for (PersonMappingObject mmo : getConfig().getPersonList()) {
				String firstname = "";
				String lastname = "";
				if (mmo.isSplitName()) {
					String name = rowMap.get(headerOrder.get(mmo.getHeaderName()));
					if (StringUtils.isNotBlank(name)) {
						if (name.contains(mmo.getSplitChar())) {
							if (mmo.isFirstNameIsFirst()) {
								firstname = name.substring(0, name.lastIndexOf(mmo.getSplitChar()));
								lastname = name.substring(name.lastIndexOf(mmo.getSplitChar()));
							} else {
								lastname = name.substring(0, name.lastIndexOf(mmo.getSplitChar())).trim();
								firstname = name.substring(name.lastIndexOf(mmo.getSplitChar()) + 1).trim();
							}
						} else {
							lastname = name;
						}
					}
				} else {
					firstname = rowMap.get(headerOrder.get(mmo.getFirstnameHeaderName()));
					lastname = rowMap.get(headerOrder.get(mmo.getLastnameHeaderName()));
				}

				String identifier = null;
				if (mmo.getNormdataHeaderName() != null) {
					identifier = rowMap.get(headerOrder.get(mmo.getNormdataHeaderName()));
				}
				if (StringUtils.isNotBlank(mmo.getRulesetName())) {
					try {
						Person p = new Person(prefs.getMetadataTypeByName(mmo.getRulesetName()));
						p.setFirstname(firstname);
						p.setLastname(lastname);

						if (identifier != null) {
							p.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);
						}
						if (anchor != null && "anchor".equals(mmo.getDocType())) {
							anchor.addPerson(p);
						} else {
							logical.addPerson(p);
						}

						// logical.addPerson(p);
					} catch (MetadataTypeNotAllowedException e) {
						log.info(e);
						// Metadata is not known or not allowed
					}
				}
			}

			for (GroupMappingObject gmo : getConfig().getGroupList()) {
				try {
					MetadataGroup group = new MetadataGroup(prefs.getMetadataGroupTypeByName(gmo.getRulesetName()));
					for (MetadataMappingObject mmo : gmo.getMetadataList()) {
						String value = rowMap.get(headerOrder.get(mmo.getIdentifier()));
						Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
						md.setValue(value);
						if (mmo.getNormdataHeaderName() != null) {
							md.setAutorityFile("gnd", "http://d-nb.info/gnd/",
									rowMap.get(headerOrder.get(mmo.getNormdataHeaderName())));
						}
						group.addMetadata(md);
					}
					for (PersonMappingObject pmo : gmo.getPersonList()) {
						Person p = new Person(prefs.getMetadataTypeByName(pmo.getRulesetName()));
						String firstname = "";
						String lastname = "";
						if (pmo.isSplitName()) {
							String name = rowMap.get(headerOrder.get(pmo.getHeaderName()));
							if (StringUtils.isNotBlank(name)) {
								if (name.contains(pmo.getSplitChar())) {
									if (pmo.isFirstNameIsFirst()) {
										firstname = name.substring(0, name.lastIndexOf(pmo.getSplitChar()));
										lastname = name.substring(name.lastIndexOf(pmo.getSplitChar()));
									} else {
										lastname = name.substring(0, name.lastIndexOf(pmo.getSplitChar()));
										firstname = name.substring(name.lastIndexOf(pmo.getSplitChar()));
									}
								} else {
									lastname = name;
								}
							}
						} else {
							firstname = rowMap.get(headerOrder.get(pmo.getFirstnameHeaderName()));
							lastname = rowMap.get(headerOrder.get(pmo.getLastnameHeaderName()));
						}

						p.setFirstname(firstname);
						p.setLastname(lastname);

						if (pmo.getNormdataHeaderName() != null) {
							p.setAutorityFile("gnd", "http://d-nb.info/gnd/",
									rowMap.get(headerOrder.get(pmo.getNormdataHeaderName())));
						}
						group.addMetadata(p);
					}
					if (anchor != null && "anchor".equals(gmo.getDocType())) {
						anchor.addMetadataGroup(group);
					} else {
						logical.addMetadataGroup(group);
					}

					// logical.addMetadataGroup(group);

				} catch (MetadataTypeNotAllowedException e) {
					log.info(e);
					// Metadata is not known or not allowed
				}
			}

			// write mets file into import folder
			process.writeMetadataFile(ff);
		} catch (WriteException | PreferencesException | MetadataTypeNotAllowedException
				| TypeNotAllowedForParentException e) {
		}

	}

	private Fileformat getRecordFromCatalogue(String identifier) throws ImportPluginException {
		ConfigOpacCatalogue coc = ConfigOpac.getInstance().getCatalogueByName(config.getOpacName());
		if (coc == null) {
			throw new ImportPluginException(
					"Catalogue with name " + config.getOpacName() + " not found. Please check goobi_opac.xml");
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
//		DocStruct anchor = null;
		try {
			ds = myRdf.getDigitalDocument().getLogicalDocStruct();
			if (ds.getType().isAnchor()) {
//				anchor = ds;
				if (ds.getAllChildren() == null || ds.getAllChildren().isEmpty()) {
					throw new ImportPluginException("Could not import record " + identifier
							+ ". Found anchor file, but no children. Try to import the child record.");
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

//	private void createMetadata(Record r, Process process ) {
//		Map<Integer, String> rowMap = (Map<Integer, String>) r;
//		String fileName=process.getMetadataFilePath();
//		for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
//
//			String value = rowMap.get(headerOrder.get(mmo.getHeaderName()));
////		if (mmo.isRequired()) {
////			if (value.isEmpty()) {
////				System.out.println("field " + mmo.getHeaderName() + " in "
////						+ config.getIdentifierHeaderName() + "is required but empty");
////			}
////			// TODO check if value is empty and collect result
////		}
////		if (!mmo.getPattern().isEmpty()) {
////			Pattern pattern = Pattern.compile(mmo.getPattern());
////			Matcher matcher = pattern.matcher(value);
////			if (!matcher.find()) {
////				System.out.println("field " + mmo.getHeaderName() + " in "
////						+ config.getIdentifierHeaderName() + "does not match pattern");
////			}
////			// TODO collect result
////
////		}
//			String identifier = null;
//			if (mmo.getNormdataHeaderName() != null) {
//				identifier = rowMap.get(headerOrder.get(mmo.getNormdataHeaderName()));
//			}
//			if (StringUtils.isNotBlank(mmo.getRulesetName()) && StringUtils.isNotBlank(value)) {
//				try {
//					Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
//					md.setValue(value);
//					if (identifier != null) {
//						md.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);
//
//					}
//					if (anchor != null && "anchor".equals(mmo.getDocType())) {
//						anchor.addMetadata(md);
//					} else {
//						logical.addMetadata(md);
//					}
//				} catch (MetadataTypeNotAllowedException e) {
//					log.info(e);
//					// Metadata is not known or not allowed
//				}
//
//				if (mmo.getRulesetName().equalsIgnoreCase("CatalogIDDigital") && !"anchor".equals(mmo.getDocType())) {
//					fileName = getImportFolder() + File.separator + value + ".xml";
//					io.setProcessTitle(value);
//					io.setMetsFilename(fileName);
//				}
//			}
//
//			if (StringUtils.isNotBlank(mmo.getPropertyName())) {
//				Processproperty p = new Processproperty();
//				p.setTitel(mmo.getPropertyName());
//				p.setWert(value);
//				io.getProcessProperties().add(p);
//			}
//		}
//
//		for (PersonMappingObject mmo : getConfig().getPersonList()) {
//			String firstname = "";
//			String lastname = "";
//			if (mmo.isSplitName()) {
//				String name = rowMap.get(headerOrder.get(mmo.getHeaderName()));
//				if (StringUtils.isNotBlank(name)) {
//					if (name.contains(mmo.getSplitChar())) {
//						if (mmo.isFirstNameIsFirst()) {
//							firstname = name.substring(0, name.lastIndexOf(mmo.getSplitChar()));
//							lastname = name.substring(name.lastIndexOf(mmo.getSplitChar()));
//						} else {
//							lastname = name.substring(0, name.lastIndexOf(mmo.getSplitChar())).trim();
//							firstname = name.substring(name.lastIndexOf(mmo.getSplitChar()) + 1).trim();
//						}
//					} else {
//						lastname = name;
//					}
//				}
//			} else {
//				firstname = rowMap.get(headerOrder.get(mmo.getFirstnameHeaderName()));
//				lastname = rowMap.get(headerOrder.get(mmo.getLastnameHeaderName()));
//			}
//
//			String identifier = null;
//			if (mmo.getNormdataHeaderName() != null) {
//				identifier = rowMap.get(headerOrder.get(mmo.getNormdataHeaderName()));
//			}
//			if (StringUtils.isNotBlank(mmo.getRulesetName())) {
//				try {
//					Person p = new Person(prefs.getMetadataTypeByName(mmo.getRulesetName()));
//					p.setFirstname(firstname);
//					p.setLastname(lastname);
//
//					if (identifier != null) {
//						p.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);
//					}
//					if (anchor != null && "anchor".equals(mmo.getDocType())) {
//						anchor.addPerson(p);
//					} else {
//						logical.addPerson(p);
//					}
//
//					// logical.addPerson(p);
//				} catch (MetadataTypeNotAllowedException e) {
//					log.info(e);
//					// Metadata is not known or not allowed
//				}
//			}
//		}
//	}
}
