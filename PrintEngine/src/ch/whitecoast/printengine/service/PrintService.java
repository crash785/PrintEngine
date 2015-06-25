package ch.whitecoast.printengine.service;

import java.io.ByteArrayInputStream;


import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import ch.whitecoast.printengine.handler.WordHandler;
import ch.whitecoast.printengine.handler.WordXHandler;
import ch.whitecoast.printengine.objects.PrintRequestObj;
import ch.whitecoast.printengine.objects.PrintResponseObj;
import ch.whitecoast.printengine.objects.Value;
import ch.whitecoast.printengine.objects.ValueMap;

// Testing
//import java.io.ByteArrayOutputStream;
//import java.io.IOException;

/**
 * @author tkaler
 * @since 15.04.2015
 */
public class PrintService {
	
	public PrintService(){
		
	}
		
	/**
	 * This webservice handles incoming requests and directs them to the relevant handler.
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 29.04.2015
	 * @param request an instance of the PrintRequestObj, filled with the request parameters.
	 * 		it holds:
	 * 			- byte[] fileData: the contents of the original Word file
				- ValueMap[] valueMaps: a List of all values to be filled in the bookmarks
				- String fileType: the file type of the original document (should be .doc or .docx)
				- String responseType: the requested file type of the returned document (can be .doc, .docx, or .pdf)
				
	 * @return an instance of the PrintResponseObj, which contains the finished document.
	 */
	public PrintResponseObj getDocument(PrintRequestObj request){
		PrintResponseObj response = new PrintResponseObj();
		
		//We have to handle the file types .doc and .docx differently
		switch(request.getFileType()){
			case ".docx": 
				//create a new WordXHandler and call the replaceBookmarks method
				WordXHandler wordXHandler = new WordXHandler();
				wordXHandler.replaceBookmarks(convertByteArrayToInputStream(request.getFileData()), convertValueMaps(request.getValueMaps()));
				
				//Now we set the response document to .pdf or to .docx, whichever was defined in the call
				if((".pdf").equals(request.getResponseType())){
					response.setFileData(wordXHandler.getPDFByteArray());
				}
				else{
					response.setFileData(wordXHandler.getWordByteArray());
				}
				response.setFileType(request.getResponseType());
				break;
			case ".doc": 
				//create a new WordHandler and call the replaceBookmarks method
				WordHandler wordHandler = new WordHandler();
				wordHandler.replaceBookmarks(convertByteArrayToInputStream(request.getFileData()), convertValueMaps(request.getValueMaps()));
				
				//Now we set the response document to .pdf or to .doc, whichever was defined in the call
				if((".pdf").equals(request.getResponseType())){
					response.setFileData(wordHandler.getPDFByteArray());
				}
				else{
					response.setFileData(wordHandler.getWordByteArray());
				}
				response.setFileType(request.getResponseType());
				
				break;
		}
		
		//Return the PrintResponseObj to the client
		return response;
	}
	
	
	/**
	 * Converts the array of ValueMaps to a more useable instance of ArrayList<HashMap<String, Value>>
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 29.04.2015
	 * @param valueMaps: an array of ValueMaps, each instance stands for a data record to be inserted in the document
	 * @return an instance of ArrayList<HashMap<String, Value>>, containing all the data to insert in the document
	 */
	private ArrayList<HashMap<String, Value>> convertValueMaps(ValueMap[] valueMaps){
		//Initiate the new object
		ArrayList<HashMap<String, Value>> list = new ArrayList<HashMap<String, Value>>();
		if(valueMaps != null){
			// Loop through all valueMaps, generating the new ArrayList
			for(int x = 0; x < valueMaps.length; x++){								
				Value[] valueMap = valueMaps[x].getValueMap();
				HashMap<String, Value> map = new HashMap<String, Value>();
				//Loop trough each valueMap, generating the new Map
				for(int v = 0; v < valueMap.length; v++){
					map.put(valueMap[v].getKey(), valueMap[v]);
				}
				list.add(map);
			}
		}
		return list;
	}
	
	
	/**
	 * Converts the byte-Array to an InputStream
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 29.04.2015
	 * @param byteArray: the byteArray to convert
	 * @return an instance of the InputStream class
	 */
	private InputStream convertByteArrayToInputStream(byte[] byteArray){
		return  new ByteArrayInputStream(byteArray);
	}
	
}
