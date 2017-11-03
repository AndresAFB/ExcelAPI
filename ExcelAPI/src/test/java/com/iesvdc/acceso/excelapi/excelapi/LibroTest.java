/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.iesvdc.acceso.excelapi.excelapi;

import java.io.IOException;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author matinal
 */
public class LibroTest {
    
    public LibroTest() {
    }
    
    @BeforeClass
    public static void setUpClass() {
    }
    
    @AfterClass
    public static void tearDownClass() {
    }
    
    @Before
    public void setUp() {
    }
    
    @After
    public void tearDown() {
    }

    /**
     * Test of getNombreArchivo method, of class Libro.
     */
    @Test
    public void testGetNombreArchivo() {
        System.out.println("getNombreArchivo");
        Libro instance = new Libro();
        String expResult = "nuevo.xlsx";
        String result = instance.getNombreArchivo();
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        // fail("The test case is a prototype.");
    }

    /**
     * Test of setNombreArchivo method, of class Libro.
     */
    @Test
    public void testSetNombreArchivo() {
        System.out.println("setNombreArchivo");
        String nombreArchivo = "unNombre.xlsx";
        Libro instance = new Libro();
        instance.setNombreArchivo(nombreArchivo);
        assertEquals(nombreArchivo, instance.getNombreArchivo());
        // TODO review the generated test code and remove the default call to fail.
        // fail("The test case is a prototype.");
    }

    /**
     * Test of addHoja method, of class Libro.
     */
    @Test
    public void testAddHoja() {
        System.out.println("addHoja");
        
        int filas=20, columnas=30;
        Hoja hoja = new Hoja("pepe", filas, columnas);
        
        for (int i=0; i<filas; i++){
            for (int j=0; j<columnas; j++){
                hoja.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        Libro instance = new Libro();
        instance.addHoja(hoja);
        try {
            assertEquals(instance.indexHoja(0).compare(hoja), false);
        } catch (ExcelAPIException ex) {
            // TODO review the generated test code and remove the default call to fail.
            fail("The test case is a prototype.");            
        }
    }

    /**
     * Test of removeHoja method, of class Libro.
     * @throws java.lang.Exception
     */
    @Test
    public void testRemoveHoja() throws Exception {
        System.out.println("removeHoja");
        int index = 0;
        Libro instance = new Libro("nuevo_libro.xlsx");
        Hoja expResult = new Hoja("Hoja 1", 5, 5);
        for (int i=0; i<5; i++){
            for (int j=0; j<5; j++){
                expResult.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        instance.addHoja(expResult);
        Hoja result = instance.removeHoja(index);
        assertEquals("Error en removeHoja(): ", result.compare(expResult), false);
        // TODO review the generated test code and remove the default call to fail.
        // fail("The test case is a prototype.");
    }

    /*
     * Test of indexHoja method, of class Libro.
     * El IndexHoja ya está hecho
     */
//    @Test
//    public void testIndexHoja() throws Exception {
//        System.out.println("indexHoja");
//        int index = 0;
//        Libro instance = new Libro();
//        Hoja expResult = null;
//        Hoja result = instance.indexHoja(index);
//        assertEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
//    }

    /**
     * Test of load method, of class Libro.
     */
    @Test
    public void testLoad_0args() throws ExcelAPIException, IOException{
        System.out.println("load");
        String file = "test.xlsx";
        Libro instance = new Libro();
        instance.load();
        // TODO review the generated test code and remove the default call to fail.
        // fail("The test case is a prototype.");
        
        
    }

    /*
     * Test of load method, of class Libro.
     * No hace falta este puesto que ya hay un testLoad.
     */
//    @Test
//    public void testLoad_String() {
//        System.out.println("load");
//        String fileName = "";
//        Libro instance = new Libro();
//        instance.load(fileName);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
//    }

    /**
     * Test of save method, of class Libro.
     */
    @Test
    public void testSave_0args() throws Exception {
        System.out.println("save");
        Libro instance = new Libro("test.xlsx");
        Hoja h1 = new Hoja("Hoja 1", 6, 6);
        Hoja h2 = new Hoja("Hoja 2", 10, 10);
        for (int i=0; i<6; i++){
            for (int j=0; j<6; j++){
                h1.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        for (int i=0; i<10; i++){
            for (int j=0; j<10; j++){
                h2.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        instance.addHoja(h1);
        instance.addHoja(h2);
        instance.save();
        // TODO review the generated test code and remove the default call to fail.
        // fail("The test case is a prototype.");
    } 
}
