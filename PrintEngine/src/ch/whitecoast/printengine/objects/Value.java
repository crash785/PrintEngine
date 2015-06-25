package ch.whitecoast.printengine.objects;

import java.io.Serializable;

/**
 * @author tkaler
 * @since 15.04.2015
 */
public class Value implements Serializable {

	private static final long serialVersionUID = -6900657573302269558L;

	private String key;
	private String[] values;
	private boolean asList;
	private String prefix;
	private String suffix;
	
	public Value(){
		
	}
	
	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param key
	 * @param values
	 * @param asList
	 * @param prefix
	 * @param suffix
	 */
	public Value(String key, String[] values, boolean asList, String prefix, String suffix){
		this.key = key;
		this.values = values;
		this.asList = asList;
		this.prefix = prefix;
		this.suffix = suffix;
	}
	
	/* - - - - - - - - - - - - - - - - - - - - - GETTERS & SETTERS - - - - - - - - - - - - - - - - - - - - - - - */
	
	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @return the key
	 */
	public String getKey() {
		return key;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param key the key to set
	 */
	public void setKey(String key) {
		this.key = key;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @return the values
	 */
	public String[] getValues() {
		return values;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param values the values to set
	 */
	public void setValues(String[] values) {
		this.values = values;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @return the asList
	 */
	public boolean isAsList() {
		return asList;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param asList the asList to set
	 */
	public void setAsList(boolean asList) {
		this.asList = asList;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @return the prefix
	 */
	public String getPrefix() {
		return prefix;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param prefix the prefix to set
	 */
	public void setPrefix(String prefix) {
		this.prefix = prefix;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @return the suffix
	 */
	public String getSuffix() {
		return suffix;
	}

	/**
	 * @TimKaler
	 * @since 15.04.2015
	 * @param suffix the suffix to set
	 */
	public void setSuffix(String suffix) {
		this.suffix = suffix;
	}
}
