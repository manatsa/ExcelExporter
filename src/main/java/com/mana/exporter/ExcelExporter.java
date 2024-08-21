package com.mana.exporter;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelExporter<T> {

    private static final Logger LOG = LogManager.getLogger(ExcelExporter.class);


    /**
     * A method to export the data set to excel
     * @param objects - the data that is to be exported to excel
     * @param outputStream - where to write the rxcel
     * @param workbook - the workbook to fill the data
     * @throws Exception
     */
    public void exportToExcel(List<T> objects, OutputStream outputStream, XSSFWorkbook workbook) throws Exception {

        LOG.info("\n\n*************** Initializing Excel Exporter ******************************\n\n");
        /**
         * get class name as object type
         */
        Class type = objects.get(0).getClass();


        /**
         * Create a sheet with class name as its name
         */
        XSSFSheet sheet=workbook.createSheet(type.getSimpleName().toLowerCase());


        /**
         * get field and store in a fields arraylist
         */
        List<Field> fields = getAllFields(type);


        /**
         * get headers from field names
         */
        List<String> headers=fields.stream().map(field -> field.getName().toUpperCase()).toList();


        /**
         * Create a header row on the worksheet
         */
        XSSFRow headerRow=sheet.createRow(0);


        LOG.info("\n************* Writing Excel Headers ***********\n");
        /**
         * Create headers on the Excel sheet
         */
        for(int i=0; i<headers.size();i++){
            XSSFCell cell=headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }

        LOG.info("\n*********** writing data rows ***************\n");

        /**
         * Loop through the list and get values
         */
        int rowNum=0;
        for (T object : objects) {
//            System.err.println(object);
            /**
             * Create a cell
             */
            XSSFRow row=sheet.createRow(++rowNum);
            for (int i=0; i<fields.size();i++) {
                fields.get(i).setAccessible(true);
                XSSFCell cell=row.createCell(i);
                if(fields.get(i).getType()==String.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(fields.get(i).get(object).toString());
                    }
                }else if(fields.get(i).getType()==Integer.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(Integer.parseInt(fields.get(i).get(object).toString()));
                    }
                }else if(fields.get(i).getType()==Double.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(Double.parseDouble(fields.get(i).get(object).toString()));
                    }
                }else if(fields.get(i).getType()==Long.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(Long.parseLong(fields.get(i).get(object).toString()));
                    }
                }else if(fields.get(i).getType()==Boolean.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(Boolean.parseBoolean(fields.get(i).get(object).toString()));
                    }
                }else if(fields.get(i).getType()== Date.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(fields.get(i).get(object).toString());
                    }
                }else if(fields.get(i).getType().isEnum()){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(fields.get(i).get(object).toString());
                    }
                }else if(fields.get(i).getType().isArray()){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(flattenCollection(Arrays.asList((Object[]) fields.get(i).get(object))));
                    }
                }else if(Collection.class.isAssignableFrom(fields.get(i).getType())){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(flattenCollection(Arrays.asList((Collection<Object>) fields.get(i).get(object))));
                    }
                }else if(Map.class.isAssignableFrom(fields.get(i).getType())){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(flattenMap((Map<Object, Object>) fields.get(i).get(object)));
                    }
                }else if(fields.get(i).getType().isEnum()){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(fields.get(i).get(object).toString());
                    }
                }
                else{

                    try {
                        if(fields.get(i).get(object)!=null) {
                            cell.setCellValue(flattenInstanceObject(fields.get(i).get(object)));
                        }
                    }catch (Exception e){
                        e.printStackTrace();
                        //throw new RuntimeException("Unsupported field type");
                    }

                }
//                System.err.println("NAME::"+field.getName()+"\tVALUE::"+field.get(object)+"\t TYPE::"+field.getType().getSimpleName());

            }

        }

        LOG.info("\n************* Done creating the in-memory workbook ***********\n");
        LOG.info("\n************* Now, writing workbook to outputStream ***********\n");
        /**
         * write workbook to output stream
         */

        try {
            workbook.write(outputStream);
            workbook.close();
        }catch (Exception e){
            e.printStackTrace();
        }

        LOG.info("\n************* ExcelExporter is done ***********\n");
    }


    /**
     * A method to create an in-memory excel workbook
     * @param objects
     * @return -an excel workbook ready to be written to an outputStream
     * @throws Exception
     */
    public Workbook createWorkbook(List<T> objects) throws Exception {

        LOG.info("\n\n*************** Initializing Excel Exporter ******************************\n\n");
        /**
         * get class name as object type
         */
        Class type = objects.get(0).getClass();


        /**
         * Create an excel workbook
         */
        XSSFWorkbook workbook=new XSSFWorkbook();

        /**
         * Create a sheet with class name as its name
         */
        XSSFSheet sheet=workbook.createSheet(type.getSimpleName().toLowerCase());


        /**
         * get field and store in a fields arraylist
         */
        List<Field> fields = getAllFields(type);


        /**
         * get headers from field names
         */
        List<String> headers=fields.stream().map(field -> field.getName().toUpperCase()).toList();


        /**
         * Create a header row on the worksheet
         */
        XSSFRow headerRow=sheet.createRow(0);


        LOG.info("\n************* Writing Excel Headers ***********\n");
        /**
         * Create headers on the Excel sheet
         */
        for(int i=0; i<headers.size();i++){
            XSSFCell cell=headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }

        LOG.info("\n*********** writing data rows ***************\n");

        /**
         * Loop through the list and get values
         */
        int rowNum=0;
        for (T object : objects) {
//            System.err.println(object);
            /**
             * Create a cell
             */
            XSSFRow row=sheet.createRow(++rowNum);
            for (int i=0; i<fields.size();i++) {
                fields.get(i).setAccessible(true);
                XSSFCell cell=row.createCell(i);
                if(fields.get(i).getType()==String.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(fields.get(i).get(object).toString());
                    }
                }else if(fields.get(i).getType()==Integer.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(Integer.parseInt(fields.get(i).get(object).toString()));
                    }
                }else if(fields.get(i).getType()==Double.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(Double.parseDouble(fields.get(i).get(object).toString()));
                    }
                }else if(fields.get(i).getType()==Long.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(Long.parseLong(fields.get(i).get(object).toString()));
                    }
                }else if(fields.get(i).getType()==Boolean.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(Boolean.parseBoolean(fields.get(i).get(object).toString()));
                    }
                }else if(fields.get(i).getType()== Date.class){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(fields.get(i).get(object).toString());
                    }
                }else if(fields.get(i).getType().isEnum()){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(fields.get(i).get(object).toString());
                    }
                }else if(fields.get(i).getType().isArray()){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(flattenCollection(Arrays.asList((Object[]) fields.get(i).get(object))));
                    }
                }else if(Collection.class.isAssignableFrom(fields.get(i).getType())){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(flattenCollection(Arrays.asList((Collection<Object>) fields.get(i).get(object))));
                    }
                }else if(Map.class.isAssignableFrom(fields.get(i).getType())){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(flattenMap((Map<Object, Object>) fields.get(i).get(object)));
                    }
                }else if(fields.get(i).getType().isEnum()){
                    if(fields.get(i).get(object)!=null) {
                        cell.setCellValue(fields.get(i).get(object).toString());
                    }
                }
                else{

                    try {
                        if(fields.get(i).get(object)!=null) {
                            cell.setCellValue(flattenInstanceObject(fields.get(i).get(object)));
                        }
                    }catch (Exception e){
                        e.printStackTrace();
                        //throw new RuntimeException("Unsupported field type");
                    }

                }
//                System.err.println("NAME::"+field.getName()+"\tVALUE::"+field.get(object)+"\t TYPE::"+field.getType().getSimpleName());

            }

        }

        LOG.info("\n************* Done creating the in-memory workbook ***********\n");

        return workbook;
    }

    /**
     * flatten a collection into a comma separated string of items
     */
    private String flattenCollection(Collection<Object> list){
        return "["+list.stream().map(item->item.toString()).collect(Collectors.joining(", "))+"]";
    }


    /**
     * flatten a collection into a comma separated string of items
     */
    private String flattenMap(Map<Object,Object> list){
        Set<Object> set=list.keySet();
        return "["+set.stream().map(item->"{"+item.toString()+" : "+list.get(item).toString()+"} ").collect(Collectors.joining(","))+"]";
    }

    /**
     * flatten a collection into a comma separated string of items
     */
    private String flattenInstanceObject(Object object){
        List<Field> fields=getAllFields(object.getClass());

        String result= "["+fields.stream().map(field -> {
            field.setAccessible(true);
            try {
                return " { "+field.getName()+" : "+field.get(object).toString()+" } ";
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }).collect(Collectors.joining(" , "))+"]";



        return result;
    }

    /**
     * Get all fields of a class
     * @param clazz
     * @return
     */
    List<Field> getAllFields(Class clazz) {
        if (clazz == null) {
            return Collections.emptyList();
        }

        List<Field> result = new ArrayList<>(getAllFields(clazz.getSuperclass()));
        List<Field> filteredFields = Arrays.stream(clazz.getDeclaredFields())
                .filter(f -> Modifier.isPublic(f.getModifiers()) || Modifier.isProtected(f.getModifiers()) || Modifier.isPrivate(f.getModifiers()))
                .toList();
        result.addAll(filteredFields);
        result.stream().map(Field::getName).forEach(System.out::println);
        return result;
    }



}
