/*
 
 */
package com.iesvdc.acceso.excelapi.excelapi;

/**
 * Esta clase almacena información del texto de una hoja de cálculo.
 * 
 * @author AndresAFB
 * @version 1.0
 */
public class Hoja {

    static double getdato(int i, int j) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
    private String[][] datos;
    private String nombre;
    private int nFilas;
    private int nColumnas;
    
    /**
     * Crea una hoja de cálculo nueva.
     */
    public Hoja() {
        this.datos = new String[5][5];
        this.nFilas = 5;
        this.nColumnas = 5;
        this.nombre = "";
    }

    /**
     * Crea una hoja nueva de tamaño nFilas por nColumnas
     * @param nFilas el número de filas
     * @param nColumnas el número de celdas que tiene cada fila
     */
    public Hoja(int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        this.nombre = "";
        this.nFilas = nFilas;
        this.nColumnas = nColumnas;
    }
    
    public Hoja(String nombre, int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        this.nombre = nombre;
        this.nFilas = nFilas;
        this.nColumnas = nColumnas;
    }
    
    public String getDato(int fila, int columna) {
        //TO-DO excepción si accedemos a una posición no válida.
        return datos [fila][columna];
    }
    
    public void setDato(String dato, int fila, int columna) {
        //TO-DO excepción si accedemos a una posición no válida.
        this.datos [fila][columna] = dato;
    }
    
    public String getNombre(){
        return this.nombre;
    }
    
    public void setNombre(String nombre){
        this.nombre = nombre;
    }

    public int getFilas() {
        return nFilas;
    }

    public int getColumnas() {
        return nColumnas;
    }
    
    public boolean compare(Hoja hoja){
        boolean iguales = true;
        
        if (this.nColumnas==hoja.getColumnas()
                && this.nFilas==hoja.getFilas()
                && this.nombre.equals(hoja.getNombre())){
                    
                for (int i=0; i<this.nFilas; i++){
                    for (int j=0; j<this.nColumnas; j++){
                        if (this.datos[i][j].equals(hoja.getDato(i, j))){
                            iguales = false;
                            break;
                        }
                    }
                    if (!iguales) break;
                }
        } else {
            iguales = false;
        }
        
        return iguales;
    }
}
