package ru.blogspot.evgeniyosipov.javadocxgen;

/**
 *
 * @author Evgeniy Osipov
 */
public class View {

    public static void main(String[] args) {

        //Parameters: rows, columns, text
        Model.setRows(Integer.parseInt(args[0]));
        Model.setColumns(Integer.parseInt(args[1]));
        Model.setText(args[2]);

        //Creating a table in document.xml
        Controller.createTable(Model.getRows(), Model.getColumns(), Model.getText());

        //Creating directories, files and their wrapping zip/docx
        Controller.createDir();
        Controller.createFiles();
        Controller.createZipDocx("javadocxgen.docx", Model.getFileNameArr());

        //Save all files and directories or only the final .docx file
        if (args.length == 4 && args[3].equalsIgnoreCase("saveAll")) {
            Controller.saveAll(true);
        } else {
            Controller.saveAll(false);
        }

        System.out.println("OK!");
    }

}
