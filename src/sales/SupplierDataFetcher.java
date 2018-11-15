package sales;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;


public class SupplierDataFetcher {

    static final String BASE_URL = "https://www.digitalmarketplace.service.gov.uk/g-cloud/suppliers";
    static List<String> pages;
    static Map<String, String[]> detailsMap;
    static Document doc;
    static XSSFWorkbook workbook;
    static XSSFSheet spreadsheet;


    public static void main(String[] args) throws Exception {
    	fetchData();
    }
    
    private static void fetchData() {
        try {

            long start = System.currentTimeMillis();
            detailsMap = new TreeMap<>();
            doc = Jsoup.connect(BASE_URL).get();
            getNavLinks();
            List<String> firstPage = getSuppliersLinks(doc.location());
            getDetails(firstPage);
            for (String page: pages) {
                List <String> pageList = getSuppliersLinks(page);
                getDetails(pageList);
            }
            createSpreadsheet();
            long end = System.currentTimeMillis();
            System.out.println("Time to complete: " + ((end - start) / 1000) / 60 + "min");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Method to get the URLs for each character in the navigation bar (e.g. 'A', 'B', 'C')
     */
    private static void getNavLinks() {
        System.out.println("Fetching navigation URLs...");
        Element navLinks = doc.getElementById("global-atoz-navigation");
        Elements links = navLinks.getElementsByTag("a");
        pages = new ArrayList<>();
        //Iterate over each page URL and add it to pages
        for (Element link: links) {
            pages.add(link.attr("abs:href"));
        }
        System.out.println("Navigation URLs have been collected.");
    }

    /**
     * Method to get a list of URLs of each supplier page for a given URL
     * @param URL - URL of the page containing the supplier pages 
     * @return a list of URLs of each supplier page 
     * @throws IOExeception due to Jsoup connect
     */
    private static List <String> getSuppliersLinks(String URL) throws IOException {
        System.out.println("Fetching supplier URLs for page " + URL);
        List <String> linkList = new ArrayList < > ();
        try {

            doc = Jsoup.connect(URL).get();
            Elements elements = doc.getElementsByClass("search-result-title");
            //Iterate over each hyperlink element and add it to linkList
            for (Element element: elements) {
                linkList.add(element.select("a[href]").attr("abs:href"));
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return linkList;
    }

    /**
     * Method to get the details of each supplier and store them into detailsMap.
     * @param linkList - List containing the URLs of each supplier page
     * @throws IOExeception due to Jsoup connect
     */
    private static void getDetails(List<String> linkList) throws IOException {

        try {
            //Get the current URL
            String currentURL = doc.location();
            //Iterate over each URL in linkList and store them into detailsMap
            for (int i = 0; i < linkList.size(); i++) {
                System.out.println("Fetching supplier details for " + linkList.get(i));
                doc = Jsoup.connect(linkList.get(i)).get();
                String supplierName = doc.select("#content > header > h1").text();
                String contactName = null;
                String contactTel = null;
                String contactEmail = null;
                String description = null;
                //Get the supplier description and contact name, if it is present     
                if (existsElement("p.supplier-description")) {
                    description = doc.select("p.supplier-description").first().text();
                }
                if (existsElement("#meta > div > p:nth-child(2) > span > span")) {
                    contactName = doc.select("#meta > div > p:nth-child(2) > span > span").first().text();
                }
                //Get each of the contact details
                //Check each span element as some supplier pages contain less information (e.g. no telephone number)
                Elements contactDetails = doc.getElementsByClass("contact-details-block");

                for (int x = 1; x < contactDetails.size(); x++) {
                    String prop = (contactDetails.get(x).getElementsByTag("span").first().attr("itemprop"));
                    if (prop.equals("telephone")) {
                        contactTel = contactDetails.get(x).text();
                    }
                    if (prop.equals("email")) {
                        contactEmail = contactDetails.get(x).text();
                    }
                    //Add supplier details to detailsMap
                    detailsMap.put(supplierName, new String[] {
                        contactName,
                        contactTel,
                        contactEmail,
                        description
                    });
                }


            }
            //Navigate back to the original URL and check if there are any more suppliers for a given character
            System.out.println("Supplier details have been fetched for current list.");
            doc = Jsoup.connect(currentURL).get();
            if (existsElement(".next")) {
                Element nextLink = doc.selectFirst(".next");
                String nextURL = nextLink.select("a[href]").attr("abs:href");
                //Recursively call method containing list of URLs for next page
                getDetails(getSuppliersLinks(nextURL));

            }
        } catch (Exception e) {
            e.printStackTrace();

        }
        System.out.println("All Supplier details have been fetched.");

    }

    /**
     * Method to create a xlsx spreadsheet containing each of the suppliers details.
     * @throws Exception
     */
    private static void createSpreadsheet() throws Exception {
        System.out.println("Creating spreadsheet...");
        try {

            //Create blank workbook
            workbook = new XSSFWorkbook();
            //Create a blank sheet
            spreadsheet = workbook.createSheet("Supplier information");
            //Create row headings
            XSSFRow headerRow = spreadsheet.createRow(0);
            headerRow.createCell(0).setCellValue("Company Name");
            headerRow.createCell(1).setCellValue("Contact Name");
            headerRow.createCell(2).setCellValue("Contact Number");
            headerRow.createCell(3).setCellValue("Contact Email");
            headerRow.createCell(4).setCellValue("Supplier Description");
            //Style headings
            CellStyle headingStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            headingStyle.setFont(font);
            for (int i = 0; i < 5; i++) {
                headerRow.getCell(i).setCellStyle(headingStyle);
                spreadsheet.autoSizeColumn(i);
            }

            int rowId = 1;
            XSSFRow row;
            // Iterate over company details and write to sheet	      
            for (String key: detailsMap.keySet()) {
                row = spreadsheet.createRow(rowId++);
                String[] detailsArr = detailsMap.get(key);
                row.createCell(0).setCellValue(key);

                int cellId = 1;
                for (String detail: detailsArr) {
                    Cell cell = row.createCell(cellId++);
                    cell.setCellValue(detail);
                }
            }
            //Create xlsx file
            FileOutputStream out = new FileOutputStream(new File("output/G-Cloud-Suppliers-List.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("G-Cloud-Suppliers-List.xlsx written successfully");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * Helper method used to identify if an element exists.
     * @param value - element value to search for
     * @return true if value exists, false otherwise
     */
    private static boolean existsElement(String value) {

        Element test = doc.select(value).first();
        if (test != null) {
            return true;
        }
        return false;
    }
}