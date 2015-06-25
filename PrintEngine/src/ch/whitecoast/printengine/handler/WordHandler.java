package ch.whitecoast.printengine.handler;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.sax.SAXResult;

import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.Fop;
import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.FormattingResults;
import org.apache.fop.apps.MimeConstants;
import org.apache.fop.apps.PageSequenceResults;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToFoConverter;
import org.apache.poi.hwpf.model.PAPX;
import org.apache.poi.hwpf.sprm.SprmBuffer;
import org.apache.poi.hwpf.usermodel.Bookmark;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.ParagraphProperties;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import ch.whitecoast.printengine.objects.Value;

/**
 * @author Tim Kaler / Whitecoast Solutions AG 
 * @since 15.04.2015
 */
public class WordHandler {
	HWPFDocument document;
	HWPFDocument tmpDoc;
	
	private final FopFactory fopFactory = FopFactory.newInstance();


	/**
	 * This method loops through all data elements and calls the corresponding methods to replace each bookmark.
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 15.04.2015
	 * @param inStream the InputStream of the original .doc-File. This will be used as the template, and the bookmarks in it
	 * 			will be replaced with each iteration of datas.
	 * 
	 * @param data
	 * 
	 * 	ArrayList<					Each item represents a data row, each item will create a new instance of the document.
	 * 		HashMap<				Contains all informations about the replacements of the bookmarks for this data item
	 * 			String,				Holds the bookmark name and is the id of this value
	 * 			Value(=ArrayList<	Contains the needed information to replace a bookmark (prefix, suffix, asList and the values)
	 * 				String>)		
	 * 			>
	 * 		>
	 */
	public void replaceBookmarks(InputStream inStream, ArrayList<HashMap<String, Value>> data) {
		try{
		
			//Loop through all elements in "data". This means, we do the following
			//for every instance of a serial letter
			for(int x = 0; x < data.size(); x++){
				
				//First, set the tmpDoc, in which we will replace the bookmarks, to be the same as the template.
				inStream.mark(0);
				tmpDoc = new HWPFDocument(inStream);
				inStream.reset();
				
				//Then loop trough each bookmark in the document
				for(int i=0;i<tmpDoc.getBookmarks().getBookmarksCount();i++){
					//Get the current bookmark
					Bookmark bm = tmpDoc.getBookmarks().getBookmark(i);
					if(bm != null){
						
						//Try to get the corresponding Value object for this bookmark
						Value value = data.get(x).get(bm.getName());
						if(value != null){
							
							//If we have a match, create the new text for the bookmark
							String bmText = "";
							
							for(int v = 0; v < value.getValues().length; v++){
								//add linebreak if asList
								if(value.isAsList()){
									bmText += "\r";
								}
								
								//add the prefix, if there is one
								if(value.getPrefix() != null && ! value.getPrefix().isEmpty())
									bmText += value.getPrefix() + " ";
								
								//add the actual new value
								bmText += value.getValues()[v];
								
								//add the suffix, if there is one
								if(value.getSuffix() != null && ! value.getSuffix().isEmpty())
									bmText += " " + value.getSuffix();
								
								//separate with commas, if not as list
								if(! value.isAsList()){
									if(v < value.getValues().length - 1){
										bmText += ", ";
									}
								}
								else{
									if(v == value.getValues().length - 1){
										bmText += "\r";
									}
								}
							}
							
							//get the range of the bookmark
							Range range = new Range(bm.getStart(), bm.getEnd(), tmpDoc); 
							if (range.text().length()>0){
								//replace the bookmark text
								range.replaceText(bmText,false); 
							}else{ 
								//if the bookmark contains no text, insert the new text after
		                        range.insertAfter(bmText);
							}
						}
					}
				}
				
				//if in the first iteration, create the new document
				if(x==0){					
					ByteArrayOutputStream bo2 = new ByteArrayOutputStream();
					tmpDoc.write(bo2);
					
					ByteArrayInputStream bi2 = new ByteArrayInputStream(bo2.toByteArray());
					
					document = new HWPFDocument(bi2);
				}
				else{
					
					//otherwise, append the contents of tmpDoc into document
					//note that the following code does not work correctly, this is still experimental.
					
					ParagraphProperties pp = new ParagraphProperties();
					pp.setPageBreakBefore(true);
					
					document.getRange().insertAfter(pp, 0);
				
					CharacterRun run = null;
					
					for(int pIndex = 0; pIndex < tmpDoc.getRange().numParagraphs(); pIndex++){
						Paragraph origPara = tmpDoc.getRange().getParagraph(pIndex);
						Paragraph p = null;
						if(run == null){
							p = document.getRange().insertAfter(origPara.cloneProperties(), 0);
						}
						else{
							p = run.insertAfter(origPara.cloneProperties(), 0);
						}
						
						for(int cIndex = 0; cIndex < origPara.numCharacterRuns(); cIndex++){
							run = p.insertAfter(origPara.getCharacterRun(cIndex).text().replace("\r", ""), origPara.getCharacterRun(cIndex).cloneProperties()) ;
						}
						
						
					}
					
				}
				
			}
		}catch(Exception e){
			System.out.println("Error in replaceBookmarks: " + e.toString());
			e.printStackTrace();
		}
	}
	
	
	/**
	 * Returns the byte array of the document
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 15.04.2015
	 * @return the document as ByteArray
	 */
	public byte[] getWordByteArray(){
		try{		
			ByteArrayOutputStream os = new ByteArrayOutputStream();
			document.write(os);
			byte[] ret = os.toByteArray();
			os.close();
			
			return ret;
		}catch(Exception e){
			System.out.println("Error in getByteArray " + e.toString());
			return null;
		}
	}


	/**
	 * Converts the document to PDF and returns its byteArray
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 29.04.2015
	 * @return the pdf-document as ByteArray
	 */
	public byte[] getPDFByteArray(){
		ByteArrayOutputStream out = null;
		try{
			WordToFoConverter wordToFoConverter = new WordToFoConverter(XMLHelper.getDocumentBuilderFactory().newDocumentBuilder().newDocument() );
	        wordToFoConverter.processDocument(document);
	        
	        FOUserAgent foUserAgent = fopFactory.newFOUserAgent();
	        
	        out = new ByteArrayOutputStream();
	        
            // Construct fop with desired output format
            Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, foUserAgent, out);

            // Setup JAXP using identity transformer
            TransformerFactory factory = TransformerFactory.newInstance();
            Transformer transformer = factory.newTransformer(); // identity transformer

            // Setup input stream
            Source src = new DOMSource(wordToFoConverter.getDocument());
            
            // Resulting SAX events (the generated FO) must be piped through to FOP
            Result res = new SAXResult(fop.getDefaultHandler());

            // Start XSLT transformation and FOP processing
            transformer.transform(src, res);
            
	        out.close();
			return out.toByteArray();
		}
		catch(Exception e){
			System.out.println("Error in getPDFByteArray " + e.toString());
		}
		return null;
	}
}
