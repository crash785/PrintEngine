package ch.whitecoast.printengine.objects;

import java.io.Serializable;

/**
 * @author tkaler
 * @since 15.04.2015
 */
public class PrintResponseObj implements Serializable {

	private static final long serialVersionUID = -6482018840089928443L;

	private byte[] fileData;
	private String fileType;
	
	public PrintResponseObj(){
		
	}
	
	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param fileData
	 * @param fileType
	 */
	public PrintResponseObj(byte[] fileData, String fileType){
		this.fileData = fileData;
		this.fileType = fileType;
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
	
	
}
