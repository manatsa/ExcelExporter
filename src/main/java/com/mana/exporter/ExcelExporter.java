package com.mana.exporter;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelExporter<T> {

    public void exportToExcel(List<T> objects, OutputStream outputStream, XSSFWorkbook workbook) throws Exception {

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
        List<Field> fields = new ArrayList<Field>(List.of(type.getDeclaredFields()));


        /**
         * get headers from field names
         */
        List<String> headers=fields.stream().map(field -> field.getName().toUpperCase()).toList();


        /**
         * Create a header row on the worksheet
         */
        XSSFRow headerRow=sheet.createRow(0);


        /**
         * Create headers on the Excel sheet
         */
        for(int i=0; i<headers.size();i++){
            XSSFCell cell=headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }


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
                System.err.println("NAME::"+fields.get(i).getName()+"\tVALUE::"+fields.get(i).get(object)+"\t TYPE::"+fields.get(i).getType().getSimpleName());
                XSSFCell cell=row.createCell(i);
                if(fields.get(i).getType()==String.class){
                    cell.setCellValue(fields.get(i).get(object).toString());
                }else if(fields.get(i).getType()==Integer.class){
                    cell.setCellValue(Integer.parseInt(fields.get(i).get(object).toString()));
                }else if(fields.get(i).getType()==Double.class){
                    cell.setCellValue(Double.parseDouble(fields.get(i).get(object).toString()));
                }else if(fields.get(i).getType()==Boolean.class){
                    cell.setCellValue(Boolean.parseBoolean(fields.get(i).get(object).toString()));
                }else if(fields.get(i).getType()== Date.class){
                    cell.setCellValue(fields.get(i).get(object).toString());
                }else if(fields.get(i).getType().isEnum()){
                    cell.setCellValue(fields.get(i).get(object).toString());
                }else if(fields.get(i).getType().isArray()){
                    cell.setCellValue(flattenCollection(Arrays.asList((Object[])fields.get(i).get(object))));
                }else if(Collection.class.isAssignableFrom(fields.get(i).getType())){
                    cell.setCellValue(flattenCollection(Arrays.asList((Collection<Object>)fields.get(i).get(object))));
                }else if(Map.class.isAssignableFrom(fields.get(i).getType())){
                    cell.setCellValue(flattenMap((Map<Object, Object>)fields.get(i).get(object)));
                }else if(fields.get(i).getType().isEnum()){
                    cell.setCellValue(fields.get(i).get(object).toString());
                }
                else{

                    throw new RuntimeException("Unsupported field type");
                }
//                System.err.println("NAME::"+field.getName()+"\tVALUE::"+field.get(object)+"\t TYPE::"+field.getType().getSimpleName());

            }

        }

        /**
         * write workbook to output stream
         */

        try {
            workbook.write(outputStream);
            workbook.close();
        }catch (Exception e){
            e.printStackTrace();
        }
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
        return "["+set.stream().map(item->"{"+item.toString()+","+list.get(item).toString()+"} ").collect(Collectors.joining(","))+"]";
    }
}
