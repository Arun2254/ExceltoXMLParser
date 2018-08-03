package exceltoxmlparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadValuesFromExcel {
	
	String destfilepath=null;
	public void generatePOXML(List<String> arrList,String targetfilename) throws InterruptedException, IOException {
		PostXMLtoMQ pxmlmq = new PostXMLtoMQ();
		try {
			String filepath = "support/PO.xml";
			destfilepath="support/Generated_POfiles/PO.xml";
			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
			Document doc = docBuilder.parse(filepath);

			// Get the root element
			Node Message = doc.getFirstChild();
			// Get the order element by tag name directly
			Node order = doc.getElementsByTagName("Order").item(0);
			System.out.println(order.getNodeName());
			
			// loop the staff child node
			NodeList list = order.getChildNodes();

			for (int i = 0; i < list.getLength(); i++) {
				
	                   Node node = list.item(i);
			   if ("OrderId".equals(node.getNodeName())) {
				   node.setTextContent(arrList.get(0));
			   }
			   if("DestinationFacilityAliasId".equals(node.getNodeName())){
					node.setTextContent(arrList.get(1));
				}
			   if("PoDueDate".equals(node.getNodeName())){
					node.setTextContent(arrList.get(2));
				}
			   if("PODate".equals(node.getNodeName())){
					node.setTextContent(arrList.get(3));
				}
			   if("LineItem".equals(node.getNodeName())){
				   Node lineitem= doc.getElementsByTagName("LineItem").item(0);
				   NodeList lineitemlist = lineitem.getChildNodes();
				   for(int j=0; j< lineitemlist.getLength(); j++){
					   Node lineitemnode = lineitemlist.item(j);
					   if("LineItemId".equals(lineitemnode.getNodeName())){
						   lineitemnode.setTextContent(arrList.get(4));
						}
					   if("ItemName".equals(lineitemnode.getNodeName())){
						   lineitemnode.setTextContent(arrList.get(5));
						}
					   if("Quantity".equals(lineitemnode.getNodeName())){
						   Node quantity= doc.getElementsByTagName("Quantity").item(0);
						   NodeList quantitylist = quantity.getChildNodes();
						   for(int k=0; k< quantitylist.getLength(); k++){
							   Node quantitynode = quantitylist.item(k);
							   if("OrderQty".equals(quantitynode.getNodeName())){
								   quantitynode.setTextContent(arrList.get(6));
								}
						   }
						}
				   }
			   }
			}


			// write the content into xml file
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			DOMSource source = new DOMSource(doc);
			StreamResult result = new StreamResult(new File(destfilepath));
			transformer.transform(source, result);
			
			System.out.println("Done");

		   } catch (ParserConfigurationException pce) {
			pce.printStackTrace();
		   } catch (TransformerException tfe) {
			tfe.printStackTrace();
		   } catch (IOException ioe) {
			ioe.printStackTrace();
		   } catch (SAXException sae) {
			sae.printStackTrace();
		   }
		String content = new String(Files.readAllBytes(Paths.get(destfilepath)), "UTF-8");
		content=content.replaceAll("\\r\\n|\\r|\\n", " ");
		content=content.replaceAll("\\s", "");
		//content=content.replaceAll("whitespace", "");
		content=content.replaceAll("><", "> <");
		String path1="<?xmlversion=\"1.0\"encoding=\"UTF-8\"standalone=\"no\"?> <tXML>";
		String path2="<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><tXML>";
		content=content.replace(path1,path2);
		content=content.replace("00:00", " 00:00");
		System.out.println(content);
		//pxmlmq.postmsgtoMQ(content);
	}
}
