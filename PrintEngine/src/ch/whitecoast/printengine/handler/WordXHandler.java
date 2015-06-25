package ch.whitecoast.printengine.handler;

import java.io.*;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.w3c.dom.Node;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import ch.whitecoast.printengine.objects.Value;

/**
 * This class provides the logic to replace bookmarks in a given docx-file with the requested values.
 * 
 * @author Tim Kaler / Whitecoast Solutions AG 
 * @since 15.04.2015
 */
public class WordXHandler {
	
	private XWPFDocument document;
	private XWPFDocument tmpDoc;
	private List<XWPFParagraph> emptyParas = new ArrayList<XWPFParagraph>();
	private List<Integer> emptyParaPositions = new ArrayList<Integer>();
	

	/**
	 * This method loops through all data elements and calls the corresponding methods to replace each bookmark.
	 * 
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 15.04.2015
	 * @param inStream the InputStream of the original .docx-File. This will be used as the template, and the bookmarks in it
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
				tmpDoc = new XWPFDocument(inStream);
				inStream.reset();
				//Then get the hashmap (-> list of key (bookmarkName) and value (text to be inserted) pairs) from 
				//the current iteration
				HashMap<String, Value> map = data.get(x);
				
				//Now we want to loop through the map, to handle each bookmark seperately.
				for(Map.Entry<String, Value> entry : map.entrySet()){
					
					Value value = entry.getValue();
					String bmText = "";
					
					//Loop through each value, and put them together according to the settings that were defined
					for(int v = 0; v < value.getValues().length; v++){
						
						//Insert the defined prefix before the actual value, if defined
						if(value.getPrefix() != null && ! value.getPrefix().isEmpty())
							bmText += value.getPrefix() + " ";
						
						bmText += value.getValues()[v];
						
						//Append the defined suffix after the actual value, if defined
						if(value.getSuffix() != null && ! value.getSuffix().isEmpty())
							bmText += " " + value.getSuffix();
						
						//Put the values together as a comma seperated string of as a list with new lines.
						if(! value.isAsList()){
							if(v < value.getValues().length - 1){
								bmText += ", ";
							}
						}
						else{
							bmText += "\n";
						}
					}
					//Call insertAtBookmark with the name of the bookmark and the text to be inserted
					insertAtBookmark(entry.getKey(), bmText);
				}
				
				//Now that we have replaced the bookmarks in tmoDoc, we want to add a new page to our serial letter
				//or create a new document if we are at the first iteration.
				if(x == 0){
					
					ByteArrayOutputStream bo2 = new ByteArrayOutputStream();
					tmpDoc.write(bo2);
					
					ByteArrayInputStream bi2 = new ByteArrayInputStream(bo2.toByteArray());
					
					document = new XWPFDocument(bi2);
				}
				else{
					document = merge(document, tmpDoc);
				}
				
			}
		}catch(Exception e){
			System.out.println("Error in replaceBookmarks: " + e.toString());
			e.printStackTrace();
		}
	}
	
	
	
	
	
    /**
     * Inserts a value at a location within the Word document specified by a
     * named bookmark.
     *
     * @param bookmarkName An instance of the String class that encapsulates the
     * name of the bookmark. Note that case is important and the case of the
     * bookmarks name within the document and that of the value passed to this
     * parameter must match.
     * @param bookmarkValue An instance of the String class that encapsulates
     * the value that should be inserted into the document at the location
     * specified by the bookmark.
     * 
     * @author Mark Beardsley / Tim Kaler
     * @since 25.06.2015
     * @throws XmlException
     */
    public final void insertAtBookmark(String bookmarkName,
            String bookmarkValue) throws XmlException {
        List<XWPFTable> tableList = null;
        Iterator<XWPFTable> tableIter = null;
        List<XWPFTableRow> rowList = null;
        Iterator<XWPFTableRow> rowIter = null;
        List<XWPFTableCell> cellList = null;
        Iterator<XWPFTableCell> cellIter = null;
        XWPFTable table = null;
        XWPFTableRow row = null;
        XWPFTableCell cell = null;

        // Firstly, deal with any paragraphs in the body of the document.
        this.procParaList(this.tmpDoc.getParagraphs(), bookmarkName, bookmarkValue);
        removeEmptyParagraphs(null);
        
        // Then check to see if there are any bookmarks in table cells. To do this
        // it is necessary to get at the list of paragraphs 'stored' within the
        // individual table cell, hence this code which get the tables from the
        // document, the rows from each table, the cells from each row and the
        // paragraphs from each cell.
        tableList = this.tmpDoc.getTables();
        tableIter = tableList.iterator();
        while (tableIter.hasNext()) {
            table = tableIter.next();
            rowList = table.getRows();
            rowIter = rowList.iterator();
            while (rowIter.hasNext()) {
                row = rowIter.next();
                cellList = row.getTableCells();
                cellIter = cellList.iterator();
                while (cellIter.hasNext()) {
                    cell = cellIter.next();
                    this.procParaList(cell.getParagraphs(),
                            bookmarkName,
                            bookmarkValue);
                    removeEmptyParagraphs(cell);
                }
            }
        }
    }
    
    /**
     * Remove Empty Paragraphs.
     * 
     * Removes all Paragraphs, that do not contain a value for the bookmark.
     * With this we are able to have optional bookmarks in the address of a letter, whitout the resulting gaps.
     *
	 * @since 19.05.2015
	 * @author Henrick Biercher / Whitecoast Solutions AG 
     */
    private void removeEmptyParagraphs(XWPFTableCell cell) {
    	if (cell != null) {
        	for(Integer paraToRemove : emptyParaPositions){
        		if (cell.getParagraphs().size() > 1) {
        			cell.removeParagraph(paraToRemove);
        		} else {
        			XWPFParagraph para =  cell.getParagraphs().get(paraToRemove);
        			int size = para.getRuns().size();
        			for (int i=0; i<size; i++) {
        				para.removeRun(0);	
        			} 
        		}
            }
    		
    	} else {
    		for(XWPFParagraph paraToRemove : emptyParas){
            	this.tmpDoc.removeBodyElement(this.tmpDoc.getPosOfParagraph(paraToRemove));
            }
    	}
		this.emptyParas = new ArrayList<XWPFParagraph>();
		this.emptyParaPositions = new ArrayList<Integer>();
    }
    

    /**
     * Inserts text into the document at the position indicated by a specific
     * bookmark. Note that the current implementation does not take account of
     * nested bookmarks, that is bookmarks that contain other bookmarks. Note
     * also that any text contained within the bookmark itself will be removed.
     *
     * @param paraList An instance of a class that implements the List interface
     * and which encapsulates references to one or more instances of the
     * XWPFParagraph class.
     * @param bookmarkName An instance of the String class that encapsulates the
     * name of the bookmark that identifies the position within the document
     * some text should be inserted.
     * @param bookmarkValue An instance of the String class that encapsulates
     * the text that should be inserted at the location specified by the
     * bookmark.
     */
    /**
     * @author Mark Beardsley / Tim Kaler
     * @since 25.06.2015
     * @param paraList
     * @throws XmlException
     */
    private final void procParaList(List<XWPFParagraph> paraList,
            String bookmarkName, String bookmarkValue) throws XmlException {
        Iterator<XWPFParagraph> paraIter = null;
        XWPFParagraph para = null;
        List<CTBookmark> bookmarkList = null;
        Iterator<CTBookmark> bookmarkIter = null;
        CTBookmark bookmark = null;
        XWPFRun run = null;
        Integer i = 0;

        // Get an Iterator for the XWPFParagraph object and step through them
        // one at a time.
        paraIter = paraList.iterator();
        while (paraIter.hasNext()) {
            para = paraIter.next();

            // Get a List of the CTBookmark object sthat the paragraph
            // 'contains' and step through these one at a time.
            bookmarkList = para.getCTP().getBookmarkStartList();
            bookmarkIter = bookmarkList.iterator();
            while (bookmarkIter.hasNext()) {
                bookmark = bookmarkIter.next();

                // If the name of the CTBookmakr object matches the value
                // encapsulated within the argumnet passed to the bookmarkName
                // parameter then this is where the text should be inserted.
                if (bookmark.getName().equals(bookmarkName)) {

                	if(bookmarkValue.isEmpty()){
                		this.emptyParas.add(para);
                		this.emptyParaPositions.add(i);
                	}
                	else{
	                    // Create a new character run to hold the value encapsulated
	                    // within the argument passed to the bookmarkValue parameter
	                    // and then test whether this new run shouold be inserted
	                    // into the document before or after the bookmark.
	                    run = para.createRun();
	                    
	                    if(bookmarkValue.indexOf("\n") != -1){
	                    	for(String subStr : bookmarkValue.split("\\n")){
	                    		run.addBreak();
	                    		run.setText(subStr);
	                    	}
	                    }
	                    else{
	                    	run.setText(bookmarkValue);
	                    }
	                                        
	                    this.replaceBookmark(bookmark, run, para);
	                }
                }
            }
            
            
            
        }
        
        
    }

    /**
     * Inserts some text into a Word document in a position that is immediately
     * after a named bookmark.
     *
     * Bookmarks can take two forms, they can either simply mark a location
     * within a document or they can do this but contain some text. The
     * difference is obvious from looking at some XML markup. The simple
     * placeholder bookmark will look like this;
     *
     * <pre>
     *
     * <w:bookmarkStart w:name="AllAlone" w:id="0"/><w:bookmarkEnd w:id="0"/>
     *
     * </pre>
     *
     * Simply a pair of tags where one tag has the name bookmarkStart, the other
     * the name bookmarkEnd and both share matching id attributes. In this case,
     * the text will simply be inserted into the document at a point immediately
     * after the bookmarkEnd tag. No styling will be applied to the text, it
     * will simply inherit the documents defaults.
     *
     * The more complex case looks like this;
     *
     * <pre>
     *
     * <w:bookmarkStart w:name="InStyledText" w:id="3"/>
     *   <w:r w:rsidRPr="00DA438C">
     *     <w:rPr>
     *       <w:rFonts w:hAnsi="Engravers MT" w:ascii="Engravers MT" w:cs="Arimo"/>
     *       <w:color w:val="FF0000"/>
     *     </w:rPr>
     *     <w:t>text</w:t>
     *   </w:r>
     * <w:bookmarkEnd w:id="3"/>
     *
     * </pre>
     *
     * Here, the user has selected the word 'text' and chosen to insert a
     * bookmark into the document at that point. So, the bookmark tags 'contain'
     * a character run that is styled. Inserting any text after this bookmark,
     * it is important to ensure that the styling is preserved and copied over
     * to the newly inserted text.
     *
     * The approach taken to dealing with both cases is similar but slightly
     * different. In both cases, the code simply steps along the document nodes
     * until it finds the bookmarkEnd tag whose ID matches that of the
     * bookmarkStart tag. Then, it will look to see if there is one further node
     * following the bookmarkEnd tag. If there is, it will insert the text into
     * the paragraph immediately in front of this node. If, on the other hand,
     * there are no more nodes following the bookmarkEnd tag, then the new run
     * will simply be positioned at the end of the paragraph.
     *
     * Styles are dealt with by 'looking' for a 'w:rPr' element whilst iterating
     * through the nodes. If one is found, its details will be captured and
     * applied to the run before the run is inserted into the paragraph. If
     * there are multiple runs between the bookmarkStart and bookmarkEnd tags
     * and these have different styles applied to them, then the style applied
     * to the last run before the bookmarkEnd tag - if any - will be cloned and
     * applied to the newly inserted text.
     *
     * @param bookmark An instance of the CTBookmark class that encapsulates
     * information about the bookmark.
     * @param run An instance of the XWPFRun class that encapsulates the text
     * that is to be inserted into the document following the bookmark.
     * @param para An instance of the XWPFParagraph class that encapsulates that
     * part of the document, a paragraph, into which the run will be inserted.
     *
     * @author Mark Beardsley / Tim Kaler
     * @since 25.06.2015
     */
    private void replaceBookmark(CTBookmark bookmark, XWPFRun run,
            XWPFParagraph para) {
        Node nextNode = null;
        Node insertBeforeNode = null;
        Node styleNode = null;
        int bookmarkStartID = 0;
        int bookmarkEndID = -1;

        // Capture the id of the bookmarkStart tag. The code will step through
        // the document nodes 'contained' within the start and end tags that have
        // matching id numbers.
        bookmarkStartID = bookmark.getId().intValue();

        // Get the node for the bookmark start tag and then enter a loop that
        // will step from one node to the next until the bookmarkEnd tag with
        // a matching id is fouind.
        nextNode = bookmark.getDomNode();

        boolean nodeRemoved = false;
        
        while (bookmarkStartID != bookmarkEndID) {

            // Get the next node along and check to see if it is a bookmarkEnd
            // tag. If it is, get its id so that the containing while loop can
            // be terminated once the correct end tag is found. Note that the
            // id will be obtained as a String and must be converted into an
            // integer. This has been coded to fail safely so that if an error
            // is encuntered converting the id to an int value, the while loop
            // will still terminate.
        	if(! nodeRemoved){
        		nextNode = nextNode.getNextSibling();
        	}
        	else{
        		nodeRemoved = false;
        	}
            if (nextNode.getNodeName().contains("bookmarkEnd")) {
                try {
                    bookmarkEndID = Integer.parseInt(
                            nextNode.getAttributes().getNamedItem("w:id").getNodeValue());
                } catch (NumberFormatException nfe) {
                    bookmarkEndID = bookmarkStartID;
                }
            } // If we are not dealing with a bookmarkEnd node, are we dealing
            // with a run node that MAY contains styling information. If so,
            // then get that style information from the run.
            else {
                if (nextNode.getNodeName().equals("w:r")) {
                    styleNode = this.getStyleNode(nextNode);
                }
                nextNode = nextNode.getNextSibling();
                para.getCTP().getDomNode().removeChild(nextNode.getPreviousSibling());
                nodeRemoved = true;
            }
        }

        // After the while loop completes, it should have located the correct
        // bookmarkEnd tag but we cannot perform an insert after only an insert
        // before operation and must, therefore, get the next node.
        insertBeforeNode = nextNode.getNextSibling();

        
        // Style the newly inserted text. Note that the code copies or clones
        // the style it found in another run, failure to do this would remove the
        // style from one node and apply it to another.
        if (styleNode != null) {
            run.getCTR().getDomNode().insertBefore(
                    styleNode.cloneNode(true), run.getCTR().getDomNode().getFirstChild());
        }

        // Finally, check to see if there was a node after the bookmarkEnd
        // tag. If there was, then this code will insert the run in front of
        // that tag. If there was no node following the bookmarkEnd tag then the
        // run will be inserted at the end of the paragarph and this was taken
        // care of at the point of creation.
        if (insertBeforeNode != null) {
            para.getCTP().getDomNode().insertBefore(
                    run.getCTR().getDomNode(), insertBeforeNode);
        }
        
    }

    /**
     * Recover styling information - if any - from another document node. Note
     * that it is only possible to accomplish this if the node is a run (w:r)
     * and this could be tested for in the code that calls this method. However,
     * a check is made in the calling code as to whether a style has been found
     * and only if a style is found is it applied. This method always returns
     * null if it does nto find a style making that checking process easier.
     *
     * @param parentNode An instance of the Node class that encapsulates a
     * reference to a document node.
     * @return An instance of the Node class that encapsulates the styling
     * information applied to a character run. Note that if no styling
     * information is found in the run OR if the node passed as an argument to
     * the parentNode parameter is NOT a run, then a null value will be
     * returned.
     *
     * @author Mark Beardsley
     * @since 25.06.2015
     * @param parentNode
     * @return
     */
    private Node getStyleNode(Node parentNode) {
        Node childNode = null;
        Node styleNode = null;
        if (parentNode != null) {

            // If the node represents a run and it has child nodes then
            // it can be processed further. Note, whilst testing the code, it
            // was observed that although it is possible to get a list of a nodes
            // children, even when a node did have children, trying to obtain this
            // list would often return a null value. This is the reason why the
            // technique of stepping from one node to the next is used here.
            if (parentNode.getNodeName().equalsIgnoreCase("w:r")
                    && parentNode.hasChildNodes()) {

                // Get the first node and catch it's reference for return if
                // the first child node is a style node (w:rPr).
                childNode = parentNode.getFirstChild();
                if (childNode.getNodeName().equals("w:rPr")) {
                    styleNode = childNode;
                } else {
                    // If the first node was not a style node and there are other
                    // child nodes remaining to be checked, then step through
                    // the remaining child nodes until either a style node is
                    // found or until all child nodes have been processed.
                    while ((childNode = childNode.getNextSibling()) != null) {
                        if (childNode.getNodeName().equals("w:rPr")) {
                            styleNode = childNode;
                            // Note setting to null here if a style node is
                            // found in order to terminate any further
                            // checking
                            childNode = null;
                        }
                    }
                }
            }
        }
        return (styleNode);
    }
    
	
	/**
	 * Converts the document into a Byte-Array
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 15.04.2015
	 * @return
	 */
	public byte[] getWordByteArray(){
		try{		
			ByteArrayOutputStream os = new ByteArrayOutputStream();
			
			document.write(os);
			byte[] ret = os.toByteArray();
			os.close();
			
			return ret;
		}catch(Exception e){
			System.out.println("Error in getWordByteArray " + e.toString());
			return null;
		}
	}
	
	
	/**
	 * Merges the two documents into one
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 17.06.2015
	 * @param doc1 The first document, this document will hold both afterwards
	 * @param doc2 The second document, this will merge into doc1
	 * @return The merged document
	 * @throws Exception
	 */
	public static XWPFDocument merge(XWPFDocument doc1, XWPFDocument doc2) throws Exception {
      
	    CTBody src1Body = doc1.getDocument().getBody();
	    CTBody src2Body = doc2.getDocument().getBody();        
	    appendBody(src1Body, src2Body);
	    
	    return doc1;
	}

	/**
	 * Appends the contents of the second body into the first one
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 17.06.2015
	 * @param src The first CTBody, will be appended by append
	 * @param append The second CTBody, will insert into src
	 * @throws Exception
	 */
	private static void appendBody(CTBody src, CTBody append) throws Exception {
	    XmlOptions optionsOuter = new XmlOptions();
	    optionsOuter.setSaveOuter();
	    String appendString = append.xmlText(optionsOuter);
	    String srcString = src.xmlText();
	    String prefix = srcString.substring(0,srcString.indexOf(">")+1);
	    String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
	    String sufix = srcString.substring( srcString.lastIndexOf("<") );
	    String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
	    CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
	    src.set(makeBody);
	}
	
	
	/**
	 * Converts the document to a PDF-File and returns its Byte[]
	 * 
	 * @author Tim Kaler / Whitecoast Solutions AG 
	 * @since 29.04.2015
	 * @return the converted PDF file as byte array
	 */
	public byte[] getPDFByteArray(){
		try{
			PdfOptions options = PdfOptions.getDefault();

			ByteArrayOutputStream out = new ByteArrayOutputStream();
			PdfConverter.getInstance().convert(document, out, options);
			
			return out.toByteArray();
		}
		catch(Exception e){
			System.out.println("Error in getPDFByteArray: " + e.toString());
			e.printStackTrace();
		}
		return null;
	}
	
}
