/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package VeriJSTLib;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author alhadi.rahman
 */
public class VeriImplements implements VeriInterface {
    private List<String> Data;
    private String delimiter;
    private String path;
    private String filename;
    
    private int innerMaxQTY;
    private List<String> innerData = new ArrayList<String>();
    private int innerGetSeriAwal = 0;
    private int innerGetSeriAkhir;
    private int innerKelipatan;
    public String inner_filename = "Inner";
    private String[] arrHeaderColInner;
    
    private int masterMaxQTY;
    private List<String> masterData = new ArrayList<String>();
    private int masterGetSeriAwal = 0;
    private int masterGetSeriAkhir;
    private int masterKelipatan;
    public String master_filename = "Master";
    private String[] arrHeaderColMaster;
    
    public String list_filename = "List";
    
    public VeriImplements(String path,String filename,String delimiter,int qtyInner,int qtyMaster,List<String> Data)
    {
        this.innerMaxQTY = qtyInner;
        this.masterMaxQTY = qtyMaster;
        this.Data = Data;
        this.delimiter = delimiter;
        this.filename = filename;
        this.path = path;
        this.innerGetSeriAkhir = this.innerMaxQTY - 1;
        this.innerKelipatan = this.innerMaxQTY;
        this.masterGetSeriAkhir = this.masterMaxQTY - 1;
        this.masterKelipatan = this.masterMaxQTY;
    }
    
    
    public void headerInner(String[] header) {
        this.arrHeaderColInner = header;
    }

    @Override
    public void exportInner() {
        if (this.arrHeaderColInner.length > 0) {
            int total_data = Data.size();
            int startTake_row = 0;
            int innersisa = total_data % innerMaxQTY;
            int innerbox_of = total_data / innerMaxQTY;
            if (innersisa > 0) {
                innerbox_of++;
            }
            innerGetSeriAwal = innerGetSeriAwal + startTake_row;
            innerGetSeriAkhir = innerGetSeriAkhir + startTake_row;
            String innertemp = "";
            int innerbox_ke = 1;
                for (int i = 0; i < total_data; i++) {
                    String DataTemp = Data.get(startTake_row);
                    String[] arrayTemp = DataTemp.split(delimiter);
                    if (innerGetSeriAwal == startTake_row) {
                        innertemp = arrayTemp[0];
                        innerGetSeriAwal = innerGetSeriAwal + innerKelipatan;
                    }
                        if (i == (total_data -1)) {
                            innertemp = innertemp+delimiter+arrayTemp[0];
                            int boxQTY = innerMaxQTY;
                            if (innersisa > 0 ) {
                               boxQTY = total_data - (innerMaxQTY * (innerbox_of - 1));
                            }
                            innertemp = innertemp + delimiter + filename+delimiter+String.valueOf(boxQTY);
                            innertemp = innertemp+delimiter+String.valueOf(innerbox_ke) +delimiter+"/"+delimiter+ String.valueOf(innerbox_of);
                            innerData.add(innertemp);
                            innertemp = "";
                        }
                        else
                        {
                            if (innerGetSeriAkhir == startTake_row) {
                                innertemp = innertemp+delimiter+arrayTemp[0];
                                innertemp = innertemp + delimiter+filename;
                                int boxQTY = innerMaxQTY;
                                innertemp = innertemp + delimiter + String.valueOf(boxQTY);
                                innertemp = innertemp+delimiter+String.valueOf(innerbox_ke) +delimiter+"/"+delimiter+ String.valueOf(innerbox_of);
                                innerData.add(innertemp);
                                innertemp = "";
                                //increment
                                innerGetSeriAkhir = innerGetSeriAkhir + innerKelipatan;
                                innerbox_ke++;
                            }
                        }
                        startTake_row++;
                }// exit loop
                
                XSSFWorkbook xlsxWorkbook=new XSSFWorkbook();
                XSSFSheet sheetxlsxWorkbook=xlsxWorkbook.createSheet("Result");
                XSSFRow rowData =   sheetxlsxWorkbook.createRow((short)0);
                
                // buat header
                for (int col = 0; col < this.arrHeaderColInner.length; col++) {
                    rowData.createCell((short) col).setCellValue(this.arrHeaderColInner[col]); 
                }
                
                for (int i = 0; i < innerData.size(); i++) {
                    String[] temp;
                    temp = innerData.get(i).split(delimiter);
                    String boxke = temp[4];
                    int lenghtBoxke = 3;
                    if (boxke.length() < lenghtBoxke) {
                        for (int j = 0; j < (lenghtBoxke-boxke.length()) + 1; j++) {
                            boxke = "0"+boxke;
                        }
                    }

                    rowData =   sheetxlsxWorkbook.createRow((short)(i+1));
                    rowData.createCell((short) 0).setCellValue(temp[0]);
                    rowData.createCell((short) 1).setCellValue(temp[1]);
                    rowData.createCell((short) 2).setCellValue(temp[2]);
                    rowData.createCell((short) 3).setCellValue(temp[3]);
                    rowData.createCell((short) 4).setCellValue(boxke);
                    rowData.createCell((short) 5).setCellValue(temp[5]);
                    rowData.createCell((short) 6).setCellValue(temp[6]);
                }

                FileOutputStream fos;
                try {
                    boolean CreatePathFolder = (new File(path + "\\"+"Verifikasi")).mkdirs();
                    String directory = path + "\\"+"Verifikasi"+"\\";
                    fos = new FileOutputStream(new File(directory+filename+"_"+inner_filename+".xlsx"));
                    xlsxWorkbook.write(fos);
                    fos.close();
                } catch (FileNotFoundException ex) {
                    Logger.getLogger(VeriImplements.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    Logger.getLogger(VeriImplements.class.getName()).log(Level.SEVERE, null, ex);
                }
        }
        else
        {
            System.out.println("Header File Inner belum terisi, Mohon dicek");
        }
    }

    public void headerMaster(String[] header) {
        this.arrHeaderColMaster = header;
    }

    @Override
    public void exportMaster() {
        if (this.arrHeaderColMaster.length > 0) {
            int total_data = Data.size();
            int startTake_row = 0;
            int mastersisa = total_data % masterMaxQTY;
            int masterbox_of = total_data / masterMaxQTY;
                if (mastersisa > 0) {
                    masterbox_of++;
                }
            masterGetSeriAwal = masterGetSeriAwal + startTake_row;
            masterGetSeriAkhir = masterGetSeriAkhir + startTake_row;
            String mastertemp = "";
            String mastertemp2 = "";
            int masterbox_ke = 1;
            for (int i = 0; i < total_data; i++) {
                String DataTemp = Data.get(startTake_row);
                String[] arrayTemp = DataTemp.split(delimiter);
                    if (masterGetSeriAwal == startTake_row) {
                                //mastertemp = DataDecrypt.get(startTake_row);
                                //mastertemp = filename;
                        mastertemp = arrayTemp[0];
                        masterGetSeriAwal = masterGetSeriAwal + masterKelipatan;
                    }
                    
                    if (i == (total_data -1)) {
                            /*if (masterGetSeriAwal > total_data -1) { // berfungsi untuk sisa
                                //Cell xxtemp = sheetfilexls.getRow((startTake_row - mastersisa)).getCell(0);
                                //mastertemp = DataDecrypt.get(startTake_row - innersisa + 1);
                                mastertemp = filename;
                                //System.out.println("SISA : " + mastersisa);
                            }*/
                            mastertemp = mastertemp+delimiter+arrayTemp[0];
                            int boxQTY = masterMaxQTY;
                            if (mastersisa > 0 ) {
                               boxQTY = total_data - (masterMaxQTY * (masterbox_of - 1));
                            }
                            mastertemp = mastertemp + delimiter + filename+delimiter+String.valueOf(boxQTY);
                            mastertemp = mastertemp+delimiter+String.valueOf(masterbox_ke) +delimiter+"/"+delimiter+ String.valueOf(masterbox_of);
                            masterData.add(mastertemp);
                            mastertemp = "";
                        }
                        else
                        {
                            if (masterGetSeriAkhir == startTake_row) {
                                mastertemp = mastertemp+delimiter+arrayTemp[0];
                                mastertemp = mastertemp + delimiter+filename;
                                int boxQTY = masterMaxQTY;
                                mastertemp = mastertemp + delimiter + String.valueOf(boxQTY);
                                mastertemp = mastertemp+delimiter+String.valueOf(masterbox_ke) +delimiter+"/"+delimiter+ String.valueOf(masterbox_of);
                                masterData.add(mastertemp);
                                mastertemp = "";
                                //increment
                                masterGetSeriAkhir = masterGetSeriAkhir + masterKelipatan;
                                masterbox_ke++;
                            }
                        }
                    startTake_row++;
            } // exit loop for
            
                XSSFWorkbook xlsxWorkbook=new XSSFWorkbook();
                XSSFSheet sheetxlsxWorkbook=xlsxWorkbook.createSheet("Result");
                XSSFRow rowData =   sheetxlsxWorkbook.createRow((short)0);
                
                // buat header
                for (int col = 0; col < this.arrHeaderColMaster.length; col++) {
                    rowData.createCell((short) col).setCellValue(this.arrHeaderColMaster[col]); 
                }
                
                for (int i = 0; i < masterData.size(); i++) {
                    String[] temp;
                    temp = masterData.get(i).split(delimiter);
                    String boxke = temp[4];
                    int lenghtBoxke = 3;
                    if (boxke.length() < lenghtBoxke) {
                        for (int j = 0; j < (lenghtBoxke-boxke.length()) + 1; j++) {
                            boxke = "0"+boxke;
                        }
                    }

                    rowData =   sheetxlsxWorkbook.createRow((short)(i+1));
                    rowData.createCell((short) 0).setCellValue(temp[0]);
                    rowData.createCell((short) 1).setCellValue(temp[1]);
                    rowData.createCell((short) 2).setCellValue(temp[2]);
                    rowData.createCell((short) 3).setCellValue(temp[3]);
                    rowData.createCell((short) 4).setCellValue(boxke);
                    rowData.createCell((short) 5).setCellValue(temp[5]);
                    rowData.createCell((short) 6).setCellValue(temp[6]);
                }

                FileOutputStream fos;
                try {
                    boolean CreatePathFolder = (new File(path + "\\"+"Verifikasi")).mkdirs();
                    String directory = path + "\\"+"Verifikasi"+"\\";
                    fos = new FileOutputStream(new File(directory+filename+"_"+master_filename+".xlsx"));
                    xlsxWorkbook.write(fos);
                    fos.close();
                } catch (FileNotFoundException ex) {
                    Logger.getLogger(VeriImplements.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    Logger.getLogger(VeriImplements.class.getName()).log(Level.SEVERE, null, ex);
                }
        }
        else
        {
            System.out.println("Header File Master belum terisi, Mohon dicek");
        }
    }
    
    public String addZero(String data,int total)
    {
        for (int i = 0; i < total; i++) {
            data = String.valueOf(i) + data;
        }
        return data;
    }

    @Override
    public void list_Data() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
    
}
