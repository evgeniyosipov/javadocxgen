package ru.blogspot.evgeniyosipov.javadocxgen;

import java.io.File;

/**
 *
 * @author Evgeniy Osipov
 */
public class Model {
    
    //Columns, rows, text
    private static int columns;
    private static int rows;
    private static String text;

    //Directories
    private static File dirRels;
    private static File dirDocProps;
    private static File dirWord;
    private static File dirWordRels;

    //Files
    private static File fileRelsRels;
    private static File fileDocPropAppXml;
    private static File fileDocPropCoreXml;
    private static File fileWordRelsDocXmlRels;
    private static File fileWordDocXml;
    private static File fileContentTypesXml;

    //An array of strings (path and file name)
    private static String[] fileNameArr;

    //StringBuilder for document.xml
    private static StringBuilder partialDocXmlBuilder;

    //Content for document.xml
    private static String partialDocXmlString;

    //Below is the content of the necessary files
    private static String bodyXmlBegin = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            + "<w:document xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
            + "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
            + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">\n "
            + "<w:body>\n"
            + "  <w:tbl>\n";

    private static String bodyXmlEnd = "</w:tbl>\n"
            + "	</w:body>\n"
            + "</w:document>";

    private static String rowXmlBegin = "\t<w:tr>\n";
    private static String rowXmlEnd = "\t</w:tr>\n";

    private static String columnXmlBegin = "\t\t<w:tc>\n"
            + "					<w:tcPr>\n"
            + "						<w:tcBorders>\n"
            + "							<w:top w:val=\"single\" w:color=\"000000\"/>\n"
            + "							<w:left w:val=\"single\" w:color=\"000000\"/>\n"
            + "							<w:bottom w:val=\"single\" w:color=\"000000\"/>\n"
            + "							<w:right w:val=\"single\" w:color=\"000000\"/>\n"
            + "						</w:tcBorders>\n"
            + "					</w:tcPr>\n"
            + "					<w:p>\n"
            + "						<w:r>\n";

    private static String columnXmlEnd = "</w:r>\n"
            + "					</w:p>\n"
            + "				</w:tc>\n";

    private static String cellTextXmlBegin = "\t\t<w:t>";
    private static String cellTextXmlEnd = "</w:t>";

    private static String relsRelsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            + "	<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\n"
            + "	<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\n"
            + "	<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\n"
            + "</Relationships>";

    private static String docPropAppXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n"
            + "	<TotalTime>0</TotalTime>\n"
            + "</Properties>";

    private static String docPropCoreXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            + "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n"
            + "	<dcterms:created xsi:type=\"dcterms:W3CDTF\">2014-04-09T15:23:25Z</dcterms:created>\n"
            + "	<dc:language>ru</dc:language>\n"
            + "	<cp:revision>0</cp:revision>\n"
            + "</cp:coreProperties>";

    private static String wordRelsDocXmlRels = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            + "	<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n"
            + "</Relationships>";

    private static String contentTypesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
            + "	<Override PartName=\"/_rels/.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
            + "	<Override PartName=\"/word/_rels/document.xml.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
            + "	<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n"
            + "	<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n"
            + "	<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n"
            + "</Types>";

    public static int getColumns() {
        return columns;
    }

    public static void setColumns(int aColumns) {
        columns = aColumns;
    }

    public static int getRows() {
        return rows;
    }

    public static void setRows(int aRows) {
        rows = aRows;
    }

    public static String getText() {
        return text;
    }

    public static void setText(String aText) {
        text = aText;
    }

    public static File getDirRels() {
        return dirRels;
    }

    public static void setDirRels(File aDirRels) {
        dirRels = aDirRels;
    }

    public static File getDirDocProps() {
        return dirDocProps;
    }

    public static void setDirDocProps(File aDirDocProps) {
        dirDocProps = aDirDocProps;
    }

    public static File getDirWord() {
        return dirWord;
    }

    public static void setDirWord(File aDirWord) {
        dirWord = aDirWord;
    }

    public static File getDirWordRels() {
        return dirWordRels;
    }

    public static void setDirWordRels(File aDirWordRels) {
        dirWordRels = aDirWordRels;
    }

    public static File getFileRelsRels() {
        return fileRelsRels;
    }

    public static void setFileRelsRels(File aFileRelsRels) {
        fileRelsRels = aFileRelsRels;
    }

    public static File getFileDocPropAppXml() {
        return fileDocPropAppXml;
    }

    public static void setFileDocPropAppXml(File aFileDocPropAppXml) {
        fileDocPropAppXml = aFileDocPropAppXml;
    }

    public static File getFileDocPropCoreXml() {
        return fileDocPropCoreXml;
    }

    public static void setFileDocPropCoreXml(File aFileDocPropCoreXml) {
        fileDocPropCoreXml = aFileDocPropCoreXml;
    }

    public static File getFileWordRelsDocXmlRels() {
        return fileWordRelsDocXmlRels;
    }

    public static void setFileWordRelsDocXmlRels(File aFileWordRelsDocXmlRels) {
        fileWordRelsDocXmlRels = aFileWordRelsDocXmlRels;
    }

    public static File getFileWordDocXml() {
        return fileWordDocXml;
    }

    public static void setFileWordDocXml(File aFileWordDocXml) {
        fileWordDocXml = aFileWordDocXml;
    }

    public static File getFileContentTypesXml() {
        return fileContentTypesXml;
    }

    public static void setFileContentTypesXml(File aFileContentTypesXml) {
        fileContentTypesXml = aFileContentTypesXml;
    }

    public static String[] getFileNameArr() {
        return fileNameArr;
    }

    public static void setFileNameArr(String[] aFileNameArr) {
        fileNameArr = aFileNameArr;
    }

    public static StringBuilder getPartialDocXmlBuilder() {
        return partialDocXmlBuilder;
    }

    public static void setPartialDocXmlBuilder(StringBuilder aPartialDocXmlBuilder) {
        partialDocXmlBuilder = aPartialDocXmlBuilder;
    }

    public static String getPartialDocXmlString() {
        return partialDocXmlString;
    }

    public static void setPartialDocXmlString(String aPartialDocXmlString) {
        partialDocXmlString = aPartialDocXmlString;
    }

    public static String getBodyXmlBegin() {
        return bodyXmlBegin;
    }

    public static String getBodyXmlEnd() {
        return bodyXmlEnd;
    }

    public static String getRowXmlBegin() {
        return rowXmlBegin;
    }

    public static String getRowXmlEnd() {
        return rowXmlEnd;
    }

    public static String getColumnXmlBegin() {
        return columnXmlBegin;
    }

    public static String getColumnXmlEnd() {
        return columnXmlEnd;
    }

    public static String getCellTextXmlBegin() {
        return cellTextXmlBegin;
    }

    public static String getCellTextXmlEnd() {
        return cellTextXmlEnd;
    }

    public static String getRelsRelsXml() {
        return relsRelsXml;
    }

    public static String getDocPropAppXml() {
        return docPropAppXml;
    }

    public static String getDocPropCoreXml() {
        return docPropCoreXml;
    }

    public static String getWordRelsDocXmlRels() {
        return wordRelsDocXmlRels;
    }

    public static String getContentTypesXml() {
        return contentTypesXml;
    }
}
