package ru.blogspot.evgeniyosipov.javadocxgen;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.zip.Deflater;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 *
 * @author Evgeniy Osipov
 */
public class Controller {

    //Creating table
    public static void createTable(int rows, int columns, String text) {
        Model.setPartialDocXmlBuilder(new StringBuilder());

        Model.getPartialDocXmlBuilder().append(Model.getBodyXmlBegin());
        for (int r = 1; r <= rows; r++) {
            Model.getPartialDocXmlBuilder().append(Model.getRowXmlBegin());
            for (int c = 1; c <= columns; c++) {
                Model.getPartialDocXmlBuilder().append(Model.getColumnXmlBegin());
                Model.getPartialDocXmlBuilder().append("\t<w:t>" + text + "</w:t>");
                Model.getPartialDocXmlBuilder().append(Model.getColumnXmlEnd());
            }
            Model.getPartialDocXmlBuilder().append(Model.getRowXmlEnd());
        }
        Model.getPartialDocXmlBuilder().append(Model.getBodyXmlEnd());

        Model.setPartialDocXmlString(Model.getPartialDocXmlBuilder().toString());
    }

    //Creating directories
    public static void createDir() {
        Model.setDirRels(new File("_rels"));
        Model.getDirRels().mkdir();
        Model.setDirDocProps(new File("docProps"));
        Model.getDirDocProps().mkdir();
        Model.setDirWord(new File("word"));
        Model.getDirWord().mkdir();
        Model.setDirWordRels(new File("word/_rels"));
        Model.getDirWordRels().mkdir();
    }

    //Creating files (UTF-8)
    public static void createFiles() {
        try {
            
            Model.setFileRelsRels(new File("_rels/.rels"));
            Model.getFileRelsRels().createNewFile();
            BufferedWriter bwRelsRels = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(Model.getFileRelsRels()), "UTF-8"));
            bwRelsRels.write(Model.getRelsRelsXml());
            bwRelsRels.flush();
            bwRelsRels.close();

            Model.setFileDocPropAppXml(new File("docProps/app.xml"));
            Model.getFileDocPropAppXml().createNewFile();
            BufferedWriter bwDocPropAppXml = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(Model.getFileDocPropAppXml()), "UTF-8"));
            bwDocPropAppXml.write(Model.getDocPropAppXml());
            bwDocPropAppXml.flush();
            bwDocPropAppXml.close();

            Model.setFileDocPropCoreXml(new File("docProps/core.xml"));
            Model.getFileDocPropCoreXml().createNewFile();
            BufferedWriter bwDocPropCoreXml = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(Model.getFileDocPropCoreXml()), "UTF-8"));
            bwDocPropCoreXml.write(Model.getDocPropCoreXml());
            bwDocPropCoreXml.flush();
            bwDocPropCoreXml.close();

            Model.setFileWordRelsDocXmlRels(new File("word/_rels/document.xml.rels"));
            Model.getFileWordRelsDocXmlRels().createNewFile();
            BufferedWriter bwWordRelsDocXmlRels = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(Model.getFileWordRelsDocXmlRels()), "UTF-8"));
            bwWordRelsDocXmlRels.write(Model.getWordRelsDocXmlRels());
            bwWordRelsDocXmlRels.flush();
            bwWordRelsDocXmlRels.close();

            Model.setFileWordDocXml(new File("word/document.xml"));
            Model.getFileWordDocXml().createNewFile();
            BufferedWriter bwWordDocXml = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(Model.getFileWordDocXml()), "UTF-8"));
            bwWordDocXml.write(Model.getPartialDocXmlString());
            bwWordDocXml.flush();
            bwWordDocXml.close();

            Model.setFileContentTypesXml(new File("[Content_Types].xml"));
            Model.getFileContentTypesXml().createNewFile();
            BufferedWriter bwContentTypesXml = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(Model.getFileContentTypesXml()), "UTF-8"));
            bwContentTypesXml.write(Model.getContentTypesXml());
            bwContentTypesXml.flush();
            bwContentTypesXml.close();

            //An array of strings (path and file name) 
            Model.setFileNameArr(new String[6]);
            Model.getFileNameArr()[0] = Model.getFileRelsRels().toString();
            Model.getFileNameArr()[1] = Model.getFileDocPropAppXml().toString();
            Model.getFileNameArr()[2] = Model.getFileDocPropCoreXml().toString();
            Model.getFileNameArr()[3] = Model.getFileWordRelsDocXmlRels().toString();
            Model.getFileNameArr()[4] = Model.getFileWordDocXml().toString();
            Model.getFileNameArr()[5] = Model.getFileContentTypesXml().toString();

        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    //Creating zip/docx file
    public static void createZipDocx(String zipFileName, String[] ToCompressFiles) {
        try {
            String[] fileNames = ToCompressFiles;

            FileInputStream inStream;

            FileOutputStream outStream = new FileOutputStream(zipFileName);
            ZipOutputStream zipOStream = new ZipOutputStream(outStream);

            zipOStream.setLevel(Deflater.BEST_COMPRESSION);

            for (int loop = 0; loop < fileNames.length; loop++) {
                inStream = new FileInputStream(fileNames[loop]);

                zipOStream.putNextEntry(new ZipEntry(fileNames[loop]));

                int i = 0;
                while ((i = inStream.read()) != -1) {
                    zipOStream.write(i);
                }

                zipOStream.closeEntry();
                inStream.close();
            }
            zipOStream.flush();
            zipOStream.close();
        } catch (IllegalArgumentException iae) {
            iae.printStackTrace();
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    //Save all files and directories, or just the end .docx file
    public static void saveAll(boolean flag) {
        if (flag == true) {
            return;
        } else {

            Model.getFileRelsRels().delete();
            Model.getFileDocPropAppXml().delete();
            Model.getFileDocPropCoreXml().delete();
            Model.getFileWordRelsDocXmlRels().delete();
            Model.getFileWordDocXml().delete();
            Model.getFileContentTypesXml().delete();

            Model.getDirRels().deleteOnExit();
            Model.getDirDocProps().deleteOnExit();
            Model.getDirDocProps().deleteOnExit();
            Model.getDirWord().deleteOnExit();
            Model.getDirWordRels().deleteOnExit();

            try {
                Thread.sleep(100);
            } catch (InterruptedException ex) {
                ex.printStackTrace();
            }

        }
    }
}
