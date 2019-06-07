package de.intranda.goobi.plugins.datatype;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.SubnodeConfiguration;
import org.jfree.util.Log;

import lombok.Data;

@Data
public class Config {

	private String publicationType;
	private String collection;
	private int firstLine;
	private int identifierColumn;
	private int conditionalColumn;
	private int rowHeader;
	private int rowDataStart;
	private int rowDataEnd;
	private List<MetadataMappingObject> metadataList = new ArrayList<>();
	private List<PersonMappingObject> personList = new ArrayList<>();
	private List<GroupMappingObject> groupList = new ArrayList<>();
	private String identifierHeaderName;

	private boolean useOpac = false;
	private String opacName;
	private String searchField;

	/**
	 * loads the &lt;config&gt; block from xml file
	 * 
	 * @param xmlConfig
	 */

	@SuppressWarnings("unchecked")
	public Config(SubnodeConfiguration xmlConfig) {

		publicationType = xmlConfig.getString("/publicationType", "Monograph");
		collection = xmlConfig.getString("/collection", "");
		firstLine = xmlConfig.getInt("/firstLine", 1);
		identifierColumn = xmlConfig.getInt("/identifierColumn", 1);
		conditionalColumn = xmlConfig.getInt("/conditionalColumn", identifierColumn);
		identifierHeaderName = xmlConfig.getString("/identifierHeaderName", null);
		rowHeader = xmlConfig.getInt("/rowHeader", 1);
		rowDataStart = xmlConfig.getInt("/rowDataStart", 2);
		rowDataEnd = xmlConfig.getInt("/rowDataEnd", 20000);

		List<HierarchicalConfiguration> mml = xmlConfig.configurationsAt("//metadata");
		for (HierarchicalConfiguration md : mml) {
			metadataList.add(getMetadata(md));
		}

		List<HierarchicalConfiguration> pml = xmlConfig.configurationsAt("//person");
		for (HierarchicalConfiguration md : pml) {
			personList.add(getPersons(md));
		}

		List<HierarchicalConfiguration> gml = xmlConfig.configurationsAt("//group");
		for (HierarchicalConfiguration md : gml) {
			String rulesetName = md.getString("@ugh");
			GroupMappingObject grp = new GroupMappingObject();
			grp.setRulesetName(rulesetName);

			String docType = md.getString("@docType", "child");
			grp.setDocType(docType);
			List<HierarchicalConfiguration> subList = md.configurationsAt("//person");
			for (HierarchicalConfiguration sub : subList) {
				PersonMappingObject pmo = getPersons(sub);
				grp.getPersonList().add(pmo);
			}
			subList = md.configurationsAt("//metadata");
			for (HierarchicalConfiguration sub : subList) {
				MetadataMappingObject pmo = getMetadata(sub);
				grp.getMetadataList().add(pmo);
			}

			groupList.add(grp);

		}
		useOpac = xmlConfig.getBoolean("/useOpac", false);
		if (useOpac) {
			opacName = xmlConfig.getString("/opacName", "ALMA WUW");
			searchField = xmlConfig.getString("/searchField", "12");
		}
	}

	private MetadataMappingObject getMetadata(HierarchicalConfiguration md) {
		String rulesetName = md.getString("@ugh");
		String propertyName = md.getString("@name");
		Integer columnNumber = md.getInteger("@column", null);
		Integer identifierColumn = md.getInteger("@identifier", null);
		String headerName = md.getString("@headerName", null);
		String normdataHeaderName = md.getString("@normdataHeaderName", null);
		String docType = md.getString("@docType", "child");
		boolean required = md.getBoolean("@required", false);
		String pattern = md.getString("@pattern", "");
		String listPath = md.getString("@list");
		ArrayList<String> validContent = new ArrayList<>();
		if (listPath != null && !listPath.isEmpty()) {
			try {
				validContent = readFileToList(listPath);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
		}

		MetadataMappingObject mmo = new MetadataMappingObject();
		mmo.setExcelColumn(columnNumber);
		mmo.setIdentifierColumn(identifierColumn);
		mmo.setPropertyName(propertyName);
		mmo.setRulesetName(rulesetName);
		mmo.setHeaderName(headerName);
		mmo.setNormdataHeaderName(normdataHeaderName);
		mmo.setDocType(docType);
		mmo.setRequired(required);
		mmo.setPattern(pattern);
		mmo.setValidContent(validContent);
		return mmo;
	}

	private ArrayList<String> readFileToList(String listPath) throws FileNotFoundException {
		Scanner s = new Scanner(new File(listPath));
		ArrayList<String> validContent = new ArrayList<>();
		while (s.hasNext()) {
			validContent.add(s.next());
		}
		s.close();
		return validContent;
	}

	private PersonMappingObject getPersons(HierarchicalConfiguration md) {
		String rulesetName = md.getString("@ugh");
		Integer firstname = md.getInteger("firstname", null);
		Integer lastname = md.getInteger("lastname", null);
		Integer identifier = md.getInteger("identifier", null);
		String headerName = md.getString("nameFieldHeader", null);
		String firstnameHeaderName = md.getString("firstnameFieldHeader", null);
		String lastnameHeaderName = md.getString("lastnameFieldHeader", null);
		String normdataHeaderName = md.getString("@normdataHeaderName", null);
		boolean splitName = md.getBoolean("splitName", false);
		String splitChar = md.getString("splitChar", " ");
		boolean firstNameIsFirstPart = md.getBoolean("splitName/@firstNameIsFirstPart", false);
		String docType = md.getString("@docType", "child");

		boolean required = md.getBoolean("@required", false);
		String pattern = md.getString("@pattern", "");

		PersonMappingObject pmo = new PersonMappingObject();
		pmo.setFirstnameColumn(firstname);
		pmo.setLastnameColumn(lastname);
		pmo.setIdentifierColumn(identifier);
		pmo.setRulesetName(rulesetName);
		pmo.setHeaderName(headerName);
		pmo.setNormdataHeaderName(normdataHeaderName);

		pmo.setFirstnameHeaderName(firstnameHeaderName);
		pmo.setLastnameHeaderName(lastnameHeaderName);
		pmo.setSplitChar(splitChar);
		pmo.setSplitName(splitName);
		pmo.setFirstNameIsFirst(firstNameIsFirstPart);
		pmo.setDocType(docType);

		pmo.setRequired(required);
		pmo.setPattern(pattern);

		return pmo;

	}

}
