import java.io.File;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class CreateLayoutFromExcel {

	
	public static Map makeMap(){
		Map fieldsMap = new HashMap();
		fieldsMap.put("Label Name","api_name__c");
		

		return fieldsMap;

	}
	public static void main(String argv[]){
		try {
			
			String fileName = "C:\\layout.xlsx";
			String objectName = "LIN02_carr_master__c";
			String layoutName = "NameOfLayout";
			Map fieldsMap = makeMap();
				
				/**XML file initialization START*/
				DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
				DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
	
				// root elements
				Document doc = docBuilder.newDocument();
				Element rootElement = doc.createElement("Layout");
				rootElement.setAttribute("xmlns", "http://soap.sforce.com/2006/04/metadata");
				doc.appendChild(rootElement);
				
				//to exclude Submit for approval Button
				Element excludeButtons = doc.createElement("excludeButtons");
				excludeButtons.setTextContent("Submit");
				rootElement.appendChild(excludeButtons);
				
				
				Element layoutSections =  doc.createElement("layoutSections");
				rootElement.appendChild(layoutSections);
				
					Element customLabel =  doc.createElement("customLabel");
					customLabel.setTextContent("true");
					layoutSections.appendChild(customLabel);
					
					
					Element detailHeading =  doc.createElement("detailHeading");
					detailHeading.setTextContent("false");
					layoutSections.appendChild(detailHeading);
					
					Element editHeading =  doc.createElement("editHeading");
					editHeading.setTextContent("false");
					layoutSections.appendChild(editHeading);
					
					Element label =  doc.createElement("label");
					label.setTextContent("Fields");
					layoutSections.appendChild(label);
					
					/**Left column*/
					Element layoutColumns =  doc.createElement("layoutColumns");
					layoutSections.appendChild(layoutColumns);
					
						Element layoutItems =  doc.createElement("layoutItems");
						layoutColumns.appendChild(layoutItems);
						
							Element behavior =  doc.createElement("behavior");
							behavior.setTextContent("Required");
							layoutItems.appendChild(behavior);
							
							Element field =  doc.createElement("field");
							field.setTextContent("LIN01__c");
							layoutItems.appendChild(field);
					
					/**Right column*/
					layoutColumns =  doc.createElement("layoutColumns");
					layoutSections.appendChild(layoutColumns);		
							
					Element style =  doc.createElement("style");
					style.setTextContent("TwoColumnsLeftToRight");
					layoutSections.appendChild(style);
				
				Element layoutColumnsLeft = null;
	            Element layoutColumnsRight = null;
	            
				/**XML file initialization END*/
			
	            Boolean toggleFlag = false;
	            
			
			    /**Excel initialization START*/
			    Workbook wb = WorkbookFactory.create(new File(fileName));
			    Sheet sheet = wb.getSheetAt(0);

			    //Iterate through each rows from first sheet
			    Iterator<Row> rowIterator = sheet.iterator();
			    /**Excel initialization END*/
			    
			    while(rowIterator.hasNext()) {
			        Row row = rowIterator.next();

			        //For each row, iterate through each columns
			        Iterator<Cell> cellIterator = row.cellIterator();
			        while(cellIterator.hasNext()) {

			            Cell cell = cellIterator.next();
			            
			            switch(cell.getCellType()) {
			                case Cell.CELL_TYPE_BOOLEAN:
			                    System.out.print(cell.getBooleanCellValue() + "\t\t");
			                    break;
			                case Cell.CELL_TYPE_NUMERIC:
			                    System.out.print(cell.getNumericCellValue() + "\t\t");
			                    break;
			                case Cell.CELL_TYPE_STRING:
			                	switch(cell.getStringCellValue()){
			                		
			                	case "pbTitle":
			                		//Probably not required
			                		break;
			                		
			                	case "pbsTitle":
			                		
			                		//create layoutSections bundle 
			    					layoutSections =  doc.createElement("layoutSections");
			    					rootElement.appendChild(layoutSections);
			    					
			    						customLabel =  doc.createElement("customLabel");
			    						customLabel.setTextContent("true");
			    						layoutSections.appendChild(customLabel);
			    						
			    						
			    						detailHeading =  doc.createElement("detailHeading");
			    						detailHeading.setTextContent("true");
			    						layoutSections.appendChild(detailHeading);
			    						
			    						editHeading =  doc.createElement("editHeading");
			    						editHeading.setTextContent("true");
			    						layoutSections.appendChild(editHeading);
			    						
			    						style =  doc.createElement("style");
			    						style.setTextContent("TwoColumnsLeftToRight");
			    						layoutSections.appendChild(style);

			                		
			                		label =  doc.createElement("label");
									if(cellIterator.hasNext()){
			                			cell = cellIterator.next();
			                			label.setTextContent(layoutName);
										layoutSections.appendChild(label);
			                			System.out.print(cell.getStringCellValue() + "\t\t\n");
			                			
			                			layoutColumnsLeft =  doc.createElement("layoutColumns");
			        					layoutSections.appendChild(layoutColumnsLeft);
			        					
			        					layoutColumnsRight =  doc.createElement("layoutColumns");
			        					layoutSections.appendChild(layoutColumnsRight);
			        					
			        					toggleFlag = false;
			                		}
			                		break;
			                		
			                	case "BLANKSPACE":
			                		//TODO
			                		
			                	default:
			                		
			                		if(toggleFlag == false){
			                			layoutItems =  doc.createElement("layoutItems");
				                		layoutColumnsLeft.appendChild(layoutItems);
										
											behavior =  doc.createElement("behavior");
											behavior.setTextContent("Edit");
											layoutItems.appendChild(behavior);
											
											field =  doc.createElement("field");
											field.setTextContent(""+fieldsMap.get(cell.getStringCellValue())); // to get API name
											layoutItems.appendChild(field);	
											toggleFlag = true;
			                		}
			                		else{
			                			layoutItems =  doc.createElement("layoutItems");
				                		layoutColumnsRight.appendChild(layoutItems);
										
											behavior =  doc.createElement("behavior");
											behavior.setTextContent("Edit");
											layoutItems.appendChild(behavior);
											
											field =  doc.createElement("field");
											field.setTextContent(""+fieldsMap.get(cell.getStringCellValue())); // to get API name
											layoutItems.appendChild(field);	
											toggleFlag = false;
			                		}
			                		System.out.print(cell.getStringCellValue() + "\t\t");
			                		break;
			                	}
			                    break;
			            }
			            
			        }System.out.println("");//End of cell iterator while loop
			        
			    }//End of row iterator while loop
			    
			    
			    /**Footer START*/
			    layoutSections =  doc.createElement("layoutSections");
				rootElement.appendChild(layoutSections);
				
					customLabel =  doc.createElement("customLabel");
					customLabel.setTextContent("true");
					layoutSections.appendChild(customLabel);
					
					
					detailHeading =  doc.createElement("detailHeading");
					detailHeading.setTextContent("true");
					layoutSections.appendChild(detailHeading);
					
					editHeading =  doc.createElement("editHeading");
					editHeading.setTextContent("false");
					layoutSections.appendChild(editHeading);
					
					label =  doc.createElement("label");
					label.setTextContent("Custom Links");
					layoutSections.appendChild(label);
					
					layoutColumns =  doc.createElement("layoutColumns");
					layoutSections.appendChild(layoutColumns);
					
					layoutColumns =  doc.createElement("layoutColumns");
					layoutSections.appendChild(layoutColumns);
					
					layoutColumns =  doc.createElement("layoutColumns");
					layoutSections.appendChild(layoutColumns);
					
					style =  doc.createElement("style");
					style.setTextContent("CustomLinks");
					layoutSections.appendChild(style);
				
				Element showEmailCheckbox =  doc.createElement("showEmailCheckbox");
				showEmailCheckbox.setTextContent("false");
				rootElement.appendChild(showEmailCheckbox);
				
				Element showHighlightsPanel =  doc.createElement("showHighlightsPanel");
				showHighlightsPanel.setTextContent("false");
				rootElement.appendChild(showHighlightsPanel);
				
				Element showInteractionLogPanel =  doc.createElement("showInteractionLogPanel");
				showInteractionLogPanel.setTextContent("false");
				rootElement.appendChild(showInteractionLogPanel);
				
				Element showRunAssignmentRulesCheckbox =  doc.createElement("showRunAssignmentRulesCheckbox");
				showRunAssignmentRulesCheckbox.setTextContent("false");
				rootElement.appendChild(showRunAssignmentRulesCheckbox);
				
				Element showSubmitAndAttachButton =  doc.createElement("showSubmitAndAttachButton");
				showSubmitAndAttachButton.setTextContent("false");
				rootElement.appendChild(showSubmitAndAttachButton);
				/**Footer END*/
			    

				// write the content into xml file
				TransformerFactory transformerFactory = TransformerFactory.newInstance();
				Transformer transformer = transformerFactory.newTransformer();
				transformer.setOutputProperty(OutputKeys.INDENT, "yes");
				transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
				//transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
				DOMSource source = new DOMSource(doc);
				StreamResult result = new StreamResult(new File("C:\\layout.xml"));

				// Output to console for testing
				// StreamResult result = new StreamResult(System.out);

				transformer.transform(source, result);

				System.out.println("File saved!");

		} catch(Exception ioe) {
		    ioe.printStackTrace();
		}
	}

}
