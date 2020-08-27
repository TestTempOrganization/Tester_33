package com.alation.lambda.s3.metadata;

import com.alation.lambda.s3.http.HtmlTable;

import org.apache.log4j.Logger;
import org.apache.poi.hdgf.extractor.VisioTextExtractor;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hpsf.extractor.HPSFPropertiesExtractor;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.Ffn;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ooxml.extractor.POIXMLPropertiesTextExtractor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlink;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.metadata.TikaCoreProperties;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.microsoft.OfficeParser;
import org.apache.tika.sax.BodyContentHandler;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;
import org.xml.sax.SAXException;

import java.io.*;
import java.net.ContentHandler;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MicrosoftFilesMetadataHandler extends MetadataHandle {

	private String originalFileName;
	public static Logger logger = Logger.getLogger(MicrosoftFilesMetadataHandler.class);
    @Override
    public String getBody() {

        logger.info("getBody().MicrosoftOctetStreamMetadataHandler " + getS3Object().getKey());

        originalFileName = getS3Object().getKey();
		if(getS3Object().getKey().toLowerCase().endsWith(".gz")) {
			getS3Object().setKey(getS3Object().getKey().substring(0,getS3Object().getKey().lastIndexOf('.')));
		}
        
        // Octet could carry varying paylods. Check for pptx file format
        if (getS3Object().getKey().endsWith("pptx")) {
            return getPptxMetadata();
        } else if (getS3Object().getKey().endsWith("docx")) {
            return getDocxMetadata();
        } else if (getS3Object().getKey().endsWith("xlsx")) {
            return getXlsxMetadata();
        } else if (getS3Object().getKey().endsWith("doc")) {
            return getDocMetadata();
        } else if (getS3Object().getKey().endsWith("xls")) {
            return getXlsMetadata();
        } else if (getS3Object().getKey().endsWith("vsd")) {
            return getVisioMetadata();
        } else if (getS3Object().getKey().endsWith("pub")) {
            return getPublisherMetadata();
        } else if (getS3Object().getKey().endsWith("ppt")) {
            return getPptMetadata();
        } else if (getS3Object().getKey().endsWith("xls")) {
            return getXlsMetadata();
        } else {
            logger.info("File format not supported");
        }

        return HtmlTable.getEmptyTableHTML();
    }

    private String getPptxMetadata() {
        Map<String, String> pptxMetadata = scanPptxMetadata();
        return formatMetadata(pptxMetadata);
    }

    private String getDocxMetadata() {
        Map<String, String> docxMetadata = scanDocxMetadata();
        return formatMetadata(docxMetadata);
    }

    private String getXlsxMetadata() {
        Map<String, String> xlsxMetadata = scanXlsxMetadata();
        return formatMetadata(xlsxMetadata);
    }

    private String getDocMetadata() {
        Map<String, String> docMetadata = scanDocMetadata();
        return formatMetadata(docMetadata);
    }

    private String getXlsMetadata() {
        Map<String, String> xlsMetadata = scanXlsMetadata();
        return formatMetadata(xlsMetadata);
    }

    private String getVisioMetadata() {
        Map<String, String> visioMetadata = scanVisioMetadata();
        return formatMetadata(visioMetadata);
    }

    private String getPublisherMetadata() {
        Map<String, String> visioMetadata = scanPublisherMetadata();
        return formatMetadata(visioMetadata);
    }

    private String getPptMetadata() {
        Map<String, String> pptMetadata = scanPptMetadata();
        return formatMetadata(pptMetadata);
    }

    /**
     * This method is used to format the metadata in html format with a heading and data
     * @param metadata map of metadata
     * @return String that displays on UI
     */
    private String formatMetadata(Map<String, String> metadata) {
        logger.info("Metadata map count: " + metadata.size());

        if (metadata != null && metadata.size() > 0) {
            StringBuffer buff = new StringBuffer();
            for (String key : metadata.keySet()) {
                String val = metadata.get(key);
                if (val != null) {
                    buff.append("<h4>" + key + "</h4>");
                    buff.append("<div>");
                    buff.append("<pre>");
                    buff.append((val));
                    buff.append("</pre>");
                    buff.append("</div>");
                }
            }

            return buff.toString();

        } else {
            logger.info("MetadataHandler.getBody got an Empty Map");
        }

        return HtmlTable.getEmptyTableHTML();
    }

    /**
     * Extract PPTX file metadata
     * @return return Metadata map
     */
    private Map<String, String> scanPptxMetadata() {
        Map<String, String> metadata = new HashMap<>();
        try {
            @SuppressWarnings("resource")
            InputStream inputStream = new ByteArrayInputStream(getFileContent());
            XMLSlideShow xmlSlideShow = new XMLSlideShow(inputStream);
            POIXMLPropertiesTextExtractor metadataTextExtractor = xmlSlideShow.getMetadataTextExtractor();
            metadata.put("numberOfSlides", String.valueOf(xmlSlideShow.getSlides().size()));
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < xmlSlideShow.getSlides().size(); i++) {
                sb.append(xmlSlideShow.getSlides().get(i).getTitle());
                if (i != xmlSlideShow.getSlides().size() - 1) {
                    sb.append(", ");
                }
            }
            metadata.put("titlesOfSlides", sb.toString());
            POIXMLProperties.CoreProperties coreProperties = metadataTextExtractor.getCoreProperties();
            metadata.put("title", coreProperties.getTitle());
            metadata.put("creator", coreProperties.getCreator());
            if (coreProperties.getCreated() != null) {
                metadata.put("createdDate", coreProperties.getCreated().toString());
            }
            metadata.put("lastModifiedByUser", coreProperties.getLastModifiedByUser());
            if (coreProperties.getLastPrinted() != null) {
                metadata.put("lastPrinted", coreProperties.getLastPrinted().toString());
            }
            metadata.put("description", coreProperties.getDescription());
            metadata.put("subject", coreProperties.getSubject());
            metadata.put("company", xmlSlideShow.getProperties().getExtendedProperties().getCompany());

        } catch (IOException e) {
            logger.error(e.getMessage(),e);
        }
        return metadata;
    }

    /**
     * Extract PPT file metadata
     * @return return Metadata map
     */
    private Map<String, String> scanPptMetadata() {
        Map<String, String> pptFileData = new HashMap<>();
        try (InputStream inputStream = new ByteArrayInputStream(getFileContent())) {
            HSLFSlideShow ppt = new HSLFSlideShow(inputStream);
            HPSFPropertiesExtractor metadataTextExtractor = ppt.getMetadataTextExtractor();
            pptFileData.put("numberOfSlides", String.valueOf(ppt.getSlides().size()));
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < ppt.getSlides().size(); i++) {
                sb.append(ppt.getSlides().get(i).getTitle());
                if (i != ppt.getSlides().size() - 1) {
                    sb.append(", ");
                }
            }
            pptFileData.put("titlesOfSlides", sb.toString());
            SummaryInformation summaryInformation = metadataTextExtractor.getSummaryInformation();
            pptFileData.put("author", summaryInformation.getAuthor());
            pptFileData.put("title", summaryInformation.getTitle());
            pptFileData.put("lastAuthor", summaryInformation.getLastAuthor());
            if (summaryInformation.getLastPrinted() != null) {
                pptFileData.put("lastPrinted", summaryInformation.getLastPrinted().toString());
            }
            if (summaryInformation.getLastSaveDateTime() != null) {
                pptFileData.put("lastEditDateTime", summaryInformation.getLastSaveDateTime().toString());
            }
            if (summaryInformation.getCreateDateTime() != null) {
                pptFileData.put("createDateTime", summaryInformation.getCreateDateTime().toString());
            }
            pptFileData.put("comments", summaryInformation.getComments());
            pptFileData.put("applicationName", summaryInformation.getApplicationName());

            DocumentSummaryInformation docSummaryInformation = metadataTextExtractor.getDocSummaryInformation();
            pptFileData.put("company ", docSummaryInformation.getCompany());
        } catch (IOException e) {
            logger.error(e.getMessage(),e);
        }
        return pptFileData;
    }


    /**
     * Extract DOC file metadata
     * @return return Metadata map
     */
    private Map<String, String> scanDocMetadata() {
        Map<String, String> docFileData = new HashMap<>();
        try {
//            			File file = new File("D:\\alation\\S3Connector\\src\\main\\java\\com\\alation\\extractors\\samples\\file-sample_500kB.doc");

            InputStream inputStream = new ByteArrayInputStream(getFileContent());
//            InputStream inputStream = new FileInputStream(file);
            HWPFDocument document = new HWPFDocument(inputStream);

            if (document.getSummaryInformation().getLastSaveDateTime() != null) {
                docFileData.put("lastEditDateTime", document.getSummaryInformation().getLastSaveDateTime().toString());
            }
            if (document.getSummaryInformation().getCreateDateTime() != null) {
                docFileData.put("creationDate", document.getSummaryInformation().getCreateDateTime().toString());
            }

            docFileData.put("Author", document.getSummaryInformation().getAuthor());
            docFileData.put("Last Author", document.getSummaryInformation().getLastAuthor());
            docFileData.put("Comments", document.getSummaryInformation().getComments());
            docFileData.put("Subject", document.getSummaryInformation().getSubject());
            docFileData.put("Application Name", document.getSummaryInformation().getApplicationName());
            docFileData.put("Char Count", String.valueOf(document.getSummaryInformation().getCharCount()));
            docFileData.put("Page Count", String.valueOf(document.getSummaryInformation().getPageCount()));
            docFileData.put("Revision Number", String.valueOf(document.getSummaryInformation().getRevNumber()));

            // Getting fonts in word doc
            Ffn[] fontNames = document.getFontTable().getFontNames();
            List<String> fontNamesList = new ArrayList<>();
            for (int i=0; i<fontNames.length; i++) {
                fontNamesList.add(fontNames[i].getMainFontName());
            }
            if (fontNamesList.size() > 0) {
                docFileData.put("Fonts", fontNamesList.toString());
            }

            if (document.getText() != null) {
                String data = document.getText().toString();
                if (data.length() > 1024) {
                    docFileData.put("Sample Content", data.substring(0, 1024));
                } else {
                    docFileData.put("Sample Content", data);
                }
            }

        } catch (IOException e) {
           logger.error(e.getMessage(),e);
        }
        return docFileData;

    }

    /**
     * Extract DOCX file metadata
     * @return return Metadata map
     */
    private Map<String, String> scanDocxMetadata() {
        Map<String, String> metadata = new HashMap<>();
        try {
            @SuppressWarnings("resource")
            InputStream inputStream = new ByteArrayInputStream(getFileContent());
            XWPFDocument document = new XWPFDocument(inputStream);
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            metadata.put("Number of Paragraphs", String.valueOf(paragraphs.size()));
            POIXMLProperties properties = document.getProperties();

            List<XWPFPictureData> pictureData = document.getAllPictures();
            metadata.put("Number of Pictures", String.valueOf(pictureData.size()));
            POIXMLProperties.CoreProperties coreProperties = properties.getCoreProperties();
            metadata.put("author", coreProperties.getCreator());
            POIXMLProperties.ExtendedProperties extendedProperties = document.getProperties().getExtendedProperties();
            metadata.put("application", extendedProperties.getApplication());
            metadata.put("appVersion", extendedProperties.getAppVersion());
            metadata.put("company", extendedProperties.getCompany());

            if (coreProperties.getCreated() != null) {
                metadata.put("creationDate", coreProperties.getCreated().toString());
            }

            if (coreProperties.getLastPrinted() != null) {
                metadata.put("lastPrinted", coreProperties.getLastPrinted().toString());
            }
            metadata.put("lastModifiedByUser", coreProperties.getLastModifiedByUser());


        } catch (Exception e) {
            logger.warn(e.getMessage(),e);
        }
        return metadata;
    }

    /**
     * Extract XLXS file metadata
     * @return return Metadata map
     */
    private Map<String, String> scanXlsxMetadata() {
        Map<String, String> excelData = new HashMap<>();
        try {
            InputStream inputStream = new ByteArrayInputStream(getFileContent());
            @SuppressWarnings("resource")
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            excelData.put("numberOfSheets", String.valueOf(workbook.getNumberOfSheets()));
            List<String> workSheetNames = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                workSheetNames.add(workbook.getSheetName(i));
            }
            excelData.put("workSheetNames", workSheetNames.toString());
            POIXMLProperties properties = workbook.getProperties();
            POIXMLProperties.CoreProperties coreProperties = properties.getCoreProperties();
            excelData.put("author", coreProperties.getCreator());
            if (coreProperties.getCreated() != null) {
                excelData.put("creationDate", coreProperties.getCreated().toString());
            }

            if (coreProperties.getLastPrinted() != null) {
                excelData.put("lastPrinted", coreProperties.getLastPrinted().toString());
            }
            excelData.put("lastModifiedByUser", coreProperties.getLastModifiedByUser());

            POIXMLProperties.ExtendedProperties extendedProperties = workbook.getProperties().getExtendedProperties();
            excelData.put("application", extendedProperties.getApplication());
            excelData.put("appVersion", extendedProperties.getAppVersion());
            excelData.put("company", extendedProperties.getCompany());

//            List<CTProperty> customProperties = properties.getCustomProperties().getUnderlyingProperties().getPropertyList();
//            for (CTProperty ctProperty : customProperties) {
//                properties.getCustomProperties().getProperty(ctProperty.getName());
//            }
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
        }
        return excelData;
    }

    /**
     * Extract XLS file metadata
     * @return return Metadata map
     */
    private Map<String, String> scanXlsMetadata() {
        Map<String, String> metadata = new HashMap<>();
        try (InputStream inputStream = new ByteArrayInputStream(getFileContent())) {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            int numberOfWorkSheets = workbook.getNumberOfSheets();
            metadata.put("numberOfWorkSheets", String.valueOf(numberOfWorkSheets));
            List<String> workSheetNames = new ArrayList<>();
            for (int i = 0; i < numberOfWorkSheets; i++) {
                workSheetNames.add(workbook.getSheetName(i));
            }
            metadata.put("workSheetNames", workSheetNames.toString());
            SummaryInformation summaryInformation = workbook.getSummaryInformation();
            if(summaryInformation.getLastSaveDateTime() != null) {
                metadata.put("lastEditDateTime", summaryInformation.getLastSaveDateTime().toString());
            }
            if (summaryInformation.getCreateDateTime() != null) {
                metadata.put("createDateTime", summaryInformation.getCreateDateTime().toString());
            }
            if (summaryInformation.getCreateDateTime() != null) {
                metadata.put("lastPrinted", summaryInformation.getLastPrinted().toString());
            }
            metadata.put("subject", summaryInformation.getSubject());
            metadata.put("author", summaryInformation.getAuthor());
            metadata.put("comments", summaryInformation.getComments());
            metadata.put("lastAuthor", summaryInformation.getLastAuthor());
            metadata.put("title", summaryInformation.getTitle());
            for (String key : metadata.keySet()) {
                if (metadata.get(key) == null) {
                    metadata.replace(key, "Unknown");
                }
            }
        } catch (IOException e) {
            logger.error(e.getMessage(),e);
        }
        return metadata;
    }

    /**
     * Extract Visio file metadata
     * @return return Metadata map
     */
    private Map<String, String> scanVisioMetadata() {

        Map<String, String> metadata = new HashMap<>();
        try (InputStream inputStream = new ByteArrayInputStream(getFileContent())) {
//            VisioTextExtractor textExtractor = new VisioTextExtractor(inputStream);
            HPSFPropertiesExtractor hpsfPropertiesExtractor = new HPSFPropertiesExtractor(new POIFSFileSystem(inputStream));
            //following two lines contain some metadata of the file
            SummaryInformation summaryInformation = hpsfPropertiesExtractor.getSummaryInformation();
            if(summaryInformation.getLastSaveDateTime() != null) {
                metadata.put("lastEditDateTime", summaryInformation.getLastSaveDateTime().toString());
            }
            if (summaryInformation.getCreateDateTime() != null) {
                metadata.put("createDateTime", summaryInformation.getCreateDateTime().toString());
            }
            if (summaryInformation.getCreateDateTime() != null) {
                metadata.put("lastPrinted", summaryInformation.getLastPrinted().toString());
            }
            metadata.put("subject", summaryInformation.getSubject());
            metadata.put("author", summaryInformation.getAuthor());
            metadata.put("comments", summaryInformation.getComments());
            metadata.put("lastAuthor", summaryInformation.getLastAuthor());
            metadata.put("title", summaryInformation.getTitle());
            for (String key : metadata.keySet()) {
                if (metadata.get(key) == null) {
                    metadata.replace(key, "Unknown");
                }
            }
        } catch (Exception e) {
           logger.error(e.getMessage(),e);
        }
        return metadata;
    }

    /**
     * Extract Publisher file metadata
     * @return return Metadata map
     */
    private Map<String, String> scanPublisherMetadata() {
        Map<String, String> metadata = new HashMap<>();
        try (InputStream inputStream = new ByteArrayInputStream(getFileContent())) {
            HPSFPropertiesExtractor hpsfPropertiesExtractor = new HPSFPropertiesExtractor(new POIFSFileSystem(inputStream));
            SummaryInformation summaryInformation = hpsfPropertiesExtractor.getSummaryInformation();
            if(summaryInformation.getLastSaveDateTime() != null) {
                metadata.put("lastEditDateTime", summaryInformation.getLastSaveDateTime().toString());
            }
            if (summaryInformation.getCreateDateTime() != null) {
                metadata.put("createDateTime", summaryInformation.getCreateDateTime().toString());
            }
            if (summaryInformation.getCreateDateTime() != null) {
                metadata.put("lastPrinted", summaryInformation.getLastPrinted().toString());
            }
            metadata.put("subject", summaryInformation.getSubject());
            metadata.put("author", summaryInformation.getAuthor());
            metadata.put("comments", summaryInformation.getComments());
            metadata.put("lastAuthor", summaryInformation.getLastAuthor());
            metadata.put("title", summaryInformation.getTitle());
            for (String key : metadata.keySet()) {
                if (metadata.get(key) == null) {
                    metadata.replace(key, "Unknown");
                }
            }
        } catch (FileNotFoundException e) {
          logger.error(e.getMessage(),e);
        } catch (IOException e) {
            logger.error(e.getMessage(),e);
        }
        return metadata;
    }


    public static void main(String[] args) {
        MicrosoftFilesMetadataHandler microsoftFilesMetadataHandler = new MicrosoftFilesMetadataHandler();
        microsoftFilesMetadataHandler.scanDocMetadata();
    }

	@Override
	public String getSchema() {
		// TODO Auto-generated method stub
		return null;
	}

}
