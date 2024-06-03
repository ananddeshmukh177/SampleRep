
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;



public class XmlToExcelConverterClass {
    private static Workbook workbook;
    private static int rowNum;

    private final static Date DATE = 0;
    private final static int TRANSACTION_TYPE = 1;
    private final static int VOUCHER_NO = 2;
    private final static int REF_NO = 3;
    private final static int REF_TYPE = 4;
    private final static int DEBTER = 5;
    private final static int REF_AMOUNT = 6;
	private final static int VOUCHER_TYPE = 8;




    public static void main(String[] args) throws Exception {
        getAndReadXml();
    }


    /**
     *
     * Downloads a XML file, reads the substance and product values and then writes them to rows on an excel file.
     *
     * @throws Exception
     */
    private static void getAndReadXml() throws Exception {
        System.out.println("getAndReadXml");

        

        /* File xmlFile = File.createTempFile("substances", "tmp");
        String xmlFileUrl = "";
        URL url = new URL(xmlFileUrl);
        System.out.println("downloading file from " + xmlFileUrl + " ...");
        FileUtils.copyURLToFile(url, xmlFile);
        System.out.println("downloading finished, parsing...");
        */
		
		File xmlFile = new File("C:/Temp/Input.xml");


        //initXls();

        Sheet sheet = workbook.getSheetAt(0);

        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        Document doc = dBuilder.parse(xmlFile);

        NodeList nList = doc.getElementsByTagName("REQUESTDATA");
        for (int i = 0; i < nList.getLength(); i++) {
            System.out.println("Processing element " + (i+1) + "/" + nList.getLength());
            Node node = nList.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE && VOUCHER_TYPE == "Receipt") {
                Element element = (Element) node;
				
                String transactionType = element.getElementsByTagName("TRANSACTION_TYPE").item(0).getTextContent();
                String voucherNo = element.getElementsByTagName("VOUCHER_NO").item(0).getTextContent();
                String ref_No = element.getElementsByTagName("REF_NO").item(0).getTextContent();
				String refType = element.getElementsByTagName("REF_TYPE").item(0).getTextContent();
                String debter = element.getElementsByTagName("DEBTER").item(0).getTextContent();
                String refAmount = element.getElementsByTagName("REF_AMOUNT").item(0).getTextContent();    
                String voucherType = element.getElementsByTagName("VCHTYPE").item(0).getTextContent();

                        Row row = sheet.createRow(rowNum++);
                        Cell cell = row.createCell(TRANSACTION_TYPE);
                        cell.setCellValue(transactionType);

                        cell = row.createCell(VOUCHER_NO);
                        cell.setCellValue(voucherNo);

                        cell = row.createCell(REF_NO);
                        cell.setCellValue(ref_No);
						
						
                        cell = row.createCell(REF_TYPE);
                        cell.setCellValue(refType);
						
                        cell = row.createCell(DEBTER);
                        cell.setCellValue(debter);
                        
						cell = row.createCell(REF_AMOUNT);
                        cell.setCellValue(refAmount);
                        
						cell = row.createCell(VOUCHER_TYPE);
                        cell.setCellValue(voucherType);

                    }
                }
            }
        }


        FileOutputStream fileOut = new FileOutputStream("C:/Temp/Result.xlsx");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();

        if (xmlFile.exists()) {
            System.out.println("delete file-> " + xmlFile.getAbsolutePath());
            if (!xmlFile.delete()) {
                System.out.println("file '" + xmlFile.getAbsolutePath() + "' was not deleted!");
            }
        }

        System.out.println("getAndReadXml finished, processed " + nList.getLength() + " REQUESTDATA!");
    }

	 
