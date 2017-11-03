/*
 */
package com.iesvdc.acceso.excelapi.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import static org.apache.poi.hssf.usermodel.HeaderFooter.file;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Esta clase almacena información de libros para generar ficheros de Excel.
 * Un libro se compone de hojas.
 * @author AndresAFB
 */
public class Libro {
    private List<Hoja> hojas;
    private String nombreArchivo;
    
    public Libro(){
        this.hojas = new ArrayList<>();
        this.nombreArchivo = "nuevo.xlsx";
    }
    
    public Libro(String nombreArchivo) {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = nombreArchivo;
    }
    
    public String getNombreArchivo(){
        return nombreArchivo;
    }
    
    public void setNombreArchivo(String nombreArchivo) {
        this.nombreArchivo = nombreArchivo;
    }
    
    public boolean addHoja(Hoja hoja) {
        return this.hojas.add(hoja);
    }
    
    public Hoja removeHoja(int index) throws ExcelAPIException {
        if(index<0 || index>this.hojas.size()) {
            throw new ExcelAPIException("Libro:removeHoja(): Posición no válida.");
        }
        return this.hojas.remove(index);
    }
    
    public Hoja indexHoja(int index) throws ExcelAPIException {
        if(index<0 || index>this.hojas.size()) {
            throw new ExcelAPIException("Libro:indexHoja(): Posición no válida.");
        }
        return this.hojas.get(index);
    }
    
    public void load() throws ExcelAPIException, FileNotFoundException, IOException{
        FileInputStream file = new FileInputStream (nombreArchivo);
        XSSFWorkbook wb = new XSSFWorkbook(nombreArchivo);
        XSSFWorkbook libro = new XSSFWorkbook();
        int nFilas = 0; 
        int nColumnas = 0;
        Libro hojas = new Libro(nombreArchivo);
        for (Hoja hoja: this.hojas){       
            Sheet sh = wb.createSheet(libro.getSheetName(nFilas));
            
            for (int i = 0; i < hoja.getFilas(); i++) {
                Row row = sh.createRow(i);
                
                for (int j = 0; j < hoja.getColumnas(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(hoja.getDato(i, j));                
                }
            }
        }
        
    }
    
   /* public void load(String fileName){
        this.nombreArchivo = fileName;
        this.load();
    }
    */
    public void save() throws ExcelAPIException {
        SXSSFWorkbook wb = new SXSSFWorkbook();
         
        for (Hoja hoja: this.hojas){       
            Sheet sh = wb.createSheet(hoja.getNombre());           
            for (int i = 0; i < hoja.getFilas(); i++) {
                Row row = sh.createRow(i);
                for (int j = 0; j < hoja.getColumnas(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(hoja.getDato(i, j));                
                }
            }
        }
        //Try para java 8
        try (FileOutputStream out = new FileOutputStream(this.nombreArchivo)){
            //(FileOutputStream out = new FileOutputStream(this.nombreArchivo));
            wb.write(out);
            //out.close();                        
        } catch (IOException ex) {
            throw new ExcelAPIException("Error al guardar el archivo");
        } finally {
            wb.dispose();
        }
    }
    
    public void save(String fileName) throws ExcelAPIException{
        this.nombreArchivo = fileName;
        this.save();
    }
    
    private void testExtension() {
        //try {
        //    wb (new FileOutputStream(nombreArchivo + ".xlsx"));
       // } catch (FileNotFoundException ex) {
       //     Logger.getLogger(Libro.class.getName()).log(Level.SEVERE, null, ex);
       // }
    }
}
