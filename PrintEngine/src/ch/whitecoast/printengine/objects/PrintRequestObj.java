package ch.whitecoast.printengine.objects;

import java.io.Serializable;

/**
 * @author tkaler
 * @since 15.04.2015
 */
public class PrintRequestObj implements Serializable {

	private static final long serialVersionUID = 5564588454242457810L;

	private byte[] fileData;
	private ValueMap[] valueMaps;
	private String fileType;
	private String responseType;
	
	public PrintRequestObj(){
		
	}
	
	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param fileData
	 * @param fieldMap
	 * @param fileType
	 */
	public PrintRequestObj(byte[] fileData, ValueMap[] valueMaps, String fileType, String responseType){
		this.fileData = fileData;
		this.valueMaps = valueMaps;
		this.fileType = fileType;
		this.responseType = responseType;
	}


	
	/* - - - - - - - - - - - - - - - - - - - - - GETTERS & SETTERS - - - - - - - - - - - - - - - - - - - - - - - */
	
	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @return the fileData
	 */
	public byte[] getFileData() {
		return fileData;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param fileData the fileData to set
	 */
	public void setFileData(byte[] fileData) {
		this.fileData = fileData;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @return the fieldMap
	 */
	public ValueMap[] getValueMaps() {
		return valueMaps;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param fieldMap the fieldMap to set
	 */
	public void setValueMaps(ValueMap[] valueMaps) {
		this.valueMaps = valueMaps;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @return the fileType
	 */
	public String getFileType() {
		return fileType;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param fileType the fileType to set
	 */
	public void setFileType(String fileType) {
		this.fileType = fileType;
	}

	/**
	 * @TimKaler
	 * @since 24.04.2015
	 * @return the responseType
	 */
	public String getResponseType() {
		return responseType;
	}

	/**
	 * @TimKaler
	 * @since 24.04.2015
	 * @param responseType the responseType to set
	 */
	public void setResponseType(String responseType) {
		this.responseType = responseType;
	}
	
	
	
}
