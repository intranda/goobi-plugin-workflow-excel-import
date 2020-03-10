package org.goobi.plugins.datatype;

import java.util.ArrayList;
import java.util.List;
import lombok.Data;

@Data
public class DataRow {
	private String rowIdentifier;
	private List<Metadatum> contentList = new ArrayList<>();
	private int invalidFields=0;

}
