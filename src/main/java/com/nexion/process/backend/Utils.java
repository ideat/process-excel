package com.nexion.process.backend;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.nexion.process.backend.entity.ConfigSheet;
import com.nexion.process.backend.entity.GtzDestiny;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.supercsv.cellprocessor.Optional;
import org.supercsv.cellprocessor.ift.CellProcessor;
import org.supercsv.io.CsvBeanWriter;
import org.supercsv.io.ICsvBeanWriter;
import org.supercsv.prefs.CsvPreference;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Utils {

    private  XSSFWorkbook myWorkBook;
    private String archivo;

    public XSSFCell getCell(String cellName){
        Pattern r = Pattern.compile("^([A-Z]+)([0-9]+)$");
        Matcher m = r.matcher(cellName);
        XSSFWorkbook wb = new XSSFWorkbook();
        if(m.matches()) {
            String columnName = m.group(1);
            int rowNumber = Integer.parseInt(m.group(2));
            if(rowNumber > 0) {
                return wb.getSheetAt(0).getRow(rowNumber-1).getCell(CellReference.convertColStringToIndex(columnName));
            }
        }
        return null;
    }

    private int getRow(String cell){
        CellReference c = new CellReference(cell);
        return c.getRow();
    }

    private int getCol(String cell){
        CellReference c = new CellReference(cell);
        return c.getCol();
    }

    @SneakyThrows
    public List<XSSFSheet> getSheets(String filename, String sheets){
        String[] arrSheets = sheets.split(",");
        List<String> listSheets = Arrays.asList(arrSheets);
        File myFile = new File(filename);
        FileInputStream fis = new FileInputStream(myFile);
        List<XSSFSheet> sheetList = new ArrayList<>();
        myWorkBook = new XSSFWorkbook (fis);
        for(String s: listSheets){
            Integer index = myWorkBook.getSheetIndex(s);
            XSSFSheet sheet = myWorkBook.getSheetAt(index);
            sheetList.add(sheet);
        }
        return sheetList;
    }

    private void createWorkbook(String filename)  {
        File myFile = new File(filename);
        FileInputStream fis = null;
        String[] a = filename.split("\\\\");
        int aa = a.length;
        archivo = a[aa-1];

        try {
            fis = new FileInputStream(myFile);
            myWorkBook = new XSSFWorkbook (fis);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private List<ConfigSheet> getConfigSheets(String configJson) throws JsonProcessingException {
        List<ConfigSheet> configSheetList = new ArrayList<>();
        ObjectMapper mapper = new ObjectMapper();
        configSheetList = mapper.readValue(configJson, new TypeReference<List<ConfigSheet>>() {});
        return configSheetList;
    }

    public List<GtzDestiny> processFile(String filename, String sheets, String configJson) throws JsonProcessingException {
//        List<XSSFSheet> listSheets = getSheets(filename,sheets);
        createWorkbook(filename);
        List<ConfigSheet> configSheetList = getConfigSheets(configJson);
        List<GtzDestiny> gtzDestinyList = new ArrayList<>();
        for(ConfigSheet c: configSheetList){
            XSSFSheet sheet = myWorkBook.getSheet(c.getNameSheet());
            List<GtzDestiny> list = processSheet(c,sheet);
            gtzDestinyList.addAll(list);
        }

        return gtzDestinyList;
    }


    public List<GtzDestiny> processSheet(ConfigSheet conf, XSSFSheet sheet){
//        String sample = sheet.getRow(getRow(conf.getCellSampleName())).getCell(getCol(conf.getCellSampleName())).getStringCellValue();
        String category ="";
        if(conf.getStaticCategory().equals("false")) {
            category = sheet.getRow(getRow(conf.getCellCategory())).getCell(getCol(conf.getCellCategory())).getStringCellValue();
        }else{
            category = conf.getCellCategory();
        }
        List<GtzDestiny> gtzDestinyList = new ArrayList<>();

        Integer rd = getRow(conf.getCellStartDate());
        Integer cd = getCol(conf.getCellStartDate());
        String typeDate = conf.getTypeDate();
        DataFormatter dataFormatter = new DataFormatter();
        String value = dataFormatter.formatCellValue(sheet.getRow(rd).getCell(cd));
        String dateP = processDate(value,typeDate,conf.getStaticDate(),conf.getConditionProcess(),conf.getYear());

        Integer rss = getRow(conf.getCellStartSampleName());
        Integer css = getCol(conf.getCellStartSampleName());
        String valueSampleName = dataFormatter.formatCellValue(sheet.getRow(rss).getCell(css));
        String sampleName = processSample(valueSampleName,conf.getTypeSample(),conf.getStaticSampleName());

        Integer rps = 0;
        Integer cps =0;
        String pointSample = "";

        int controlDate = 0;
        int controlSampleName = 0;
        int posDate = 1;
        int posSample = 1;
        for(int f=0;f<conf.getTotalColDate();f++) {

           if(conf.getColSpaceDate()==0) {
               if (controlDate == conf.getColSpaceDate()) {
                   String datex = dataFormatter.formatCellValue(sheet.getRow(rd).getCell(cd + f));
                   dateP = processDate(datex, typeDate, conf.getStaticDate(), conf.getConditionProcess(), conf.getYear());
                   controlDate = 0;
                   posDate++;
               } else {
                   controlDate++;
                   posDate++;
               }
           }else{
               if (controlDate == conf.getColSpaceDate()) {
                   String datex = dataFormatter.formatCellValue(sheet.getRow(rd).getCell(cd + f));
                   dateP = processDate(datex, typeDate, conf.getStaticDate(), conf.getConditionProcess(), conf.getYear());
                   controlDate = conf.getColSpaceDate() - (conf.getColSpaceDate() - 1);
                   posDate++;
               } else {
                   controlDate++;
                   posDate++;
               }
           }
           //Nombre muestra
           if (controlSampleName == conf.getColSpaceSampleName()) {
               valueSampleName = processCellValue(sheet.getRow(rss).getCell(css+f)); //dataFormatter.formatCellValue(sheet.getRow(rss).getCell(css+f));
               sampleName = processSample(valueSampleName, conf.getTypeSample(),conf.getStaticSampleName());
               controlSampleName = 0;
               posSample++;
           } else {
               controlSampleName++;
               posSample++;
           }

           //Punto de muestra
            if(conf.getStaticPointSample().equals("split0")){
                rps = getRow(conf.getCellPointSample());
                cps = getCol(conf.getCellPointSample());
                String pointS = processCellValue(sheet.getRow(rps).getCell(cps+f));
                pointSample = pointS.split("\\(")[0];
            }else
            if(conf.getStaticPointSample().equals("false")) {
                rps = getRow(conf.getCellPointSample());
                cps = getCol(conf.getCellPointSample());
                String pointS = processCellValue(sheet.getRow(rps).getCell(cps+f));// dataFormatter.formatCellValue(sheet.getRow(rps).getCell(cps+f));
                pointSample = pointS.replace("Ã±","ñ");
            }else {
                pointSample = conf.getCellPointSample();
            }

           //Lugar de muestreo
           String samplingLocation = "";
           if(conf.getStaticSamplingLocation().equals("true")){
               samplingLocation = conf.getCellSamplingLocation();
           }else{
               Integer rsl = getRow(conf.getCellSamplingLocation());
               Integer csl = getCol(conf.getCellSamplingLocation());
               samplingLocation = processCellValue( sheet.getRow(rsl).getCell(csl)); // sheet.getRow(rsl).getCell(csl).getStringCellValue();
           }

            //Cicle for row parameters
            GtzDestiny gtzDestiny = new GtzDestiny();
            for (int i = 0; i < conf.getNumberRowsParameter(); i++) {
                Integer rp = getRow(conf.getCellParameter());
                Integer cp = getCol(conf.getCellParameter());

                String nameParameter = sheet.getRow(rp + i).getCell(cp).getStringCellValue(); //obtiene nombre parametro ej. ph

                Integer rvs = getRow(conf.getCellValueSample());
                Integer cvs = getCol(conf.getCellValueSample());


                //Unidades
                Integer rum = getRow(conf.getCellUnitMeasure());
                Integer cum = getCol(conf.getCellUnitMeasure());
                String valueUnitMeasure =  dataFormatter.formatCellValue(sheet.getRow(rum+i).getCell(cum));



                //CONTROL MINIMO
                if(category.toUpperCase().contains("MINIMO")) {
                    if (nameParameter.toUpperCase().contains("PH")) {
                        String valuePh =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCm_ph(valuePh);
                        gtzDestiny.setUm_ph(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("CONDUCTIVIDAD")) {
                        String valueCond =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCm_conductividad(valueCond);
                        gtzDestiny.setUm_conductividad(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("TURVIEDAD") || nameParameter.toUpperCase().contains("TURBIEDAD")) {
                        String valueTurb =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCm_turbiedad(valueTurb);
                        gtzDestiny.setUm_turbiedad(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("CLORO")) {
                        String valueCloro =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCm_cloro_residual(valueCloro);
                        gtzDestiny.setUm_cloro_residual(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("COLIFORMES") || nameParameter.toUpperCase().contains("FECALES")) {
                        String valueColi =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCm_coliforme_termotolerante_u1(valueColi);
                        gtzDestiny.setUm_coliforme_termotolerante_u1(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("TEMPERATURA")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCm_temperatura(valueTemp);
                        gtzDestiny.setUm_temperatura(valueUnitMeasure);
                    }else if (nameParameter.toUpperCase().contains("ECHERICHIA") || nameParameter.toUpperCase().contains("ESCHERICHIA")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); // dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCm_escherichia_coli_u1(valueTemp);
                        gtzDestiny.setUm_escherichia_coli_u1(valueUnitMeasure);
                    }
                }
                //CONTROL BASICO
                if(category.toUpperCase().contains("BASICO") || category.toUpperCase().contains("BÁSICO")) {
                    if (nameParameter.toUpperCase().contains("COLOR")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));// dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_fis_color(valueTemp);
                        gtzDestiny.setUm_fis_color(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("SABOR")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));// dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_fis_sabor_olor_aceptables(valueTemp);
                        gtzDestiny.setUm_fis_sabor_olor_aceptables(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("SOLIDOS") || nameParameter.toUpperCase().contains("SOLIDO") || nameParameter.toUpperCase().contains("SÓLIDOS")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_quim_solidos_disueltos_totales(valueTemp);
                        gtzDestiny.setUm_quim_solidos_disueltos_totales(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("ALCALINIDAD")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_quim_alcalinidad_total(valueTemp);
                        gtzDestiny.setUm_quim_alcalinidad_total(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("CALCIO")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_quim_calcio(valueTemp);
                        gtzDestiny.setUm_quim_calcio(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("CLORUROS")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));// dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_quim_cloruros(valueTemp);
                        gtzDestiny.setUm_quim_cloruros(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("DUREZA")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_quim_dureza_total(valueTemp);
                        gtzDestiny.setUm_quim_dureza_total(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("HIERRO")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_quim_hierro_total(valueTemp);
                        gtzDestiny.setUm_quim_hierro_total(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("MAGNESIO")) {
                        String valueTemp = processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));//dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_quim_magnesio(valueTemp);
                        gtzDestiny.setUm_quim_magnesio(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("MANGANESO")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_quim_manganeso(valueTemp);
                        gtzDestiny.setUm_quim_manganeso(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("SODIO")) { //TODO: SE CAMBIO AL CONTROL COMPLEMENTARIO
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));// dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCc_quim_inor_sodio(valueTemp);
                        gtzDestiny.setUm_quim_inor_sodio(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("SULFATOS")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));// dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_quim_sulfatos(valueTemp);
                        gtzDestiny.setUm_quim_sulfatos(valueUnitMeasure);
                    }else if (nameParameter.toUpperCase().contains("HETEROTROFICAS") || nameParameter.toUpperCase().contains("HETEROTROFICA")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));// dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCb_micro_heterotroficas(valueTemp);
                        gtzDestiny.setUm_micro_heterotroficas(valueUnitMeasure);
                    }
                }
                //COMPLEMENTARIO
                if(category.toUpperCase().contains("COMPLEMENTARIO") || category.toUpperCase().contains("COMPLEMENTARIOS")) {
                    if (nameParameter.toUpperCase().contains("ALUMINIO")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCc_quim_inor_aluminio(valueTemp);
                        gtzDestiny.setUm_quim_inor_aluminio(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("AMONIACO")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCc_quim_inor_amonio(valueTemp);
                        gtzDestiny.setUm_quim_inor_amonio(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("ARSENICO")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCc_quim_inor_arsenico(valueTemp);
                        gtzDestiny.setUm_quim_inor_arsenico(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("BORO")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCc_quim_inor_boro(valueTemp);
                        gtzDestiny.setUm_quim_inor_boro(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("COBRE")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));// dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCc_quim_inor_cobre(valueTemp);
                        gtzDestiny.setUm_quim_inor_cobre(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("FLORURO") || nameParameter.toUpperCase().contains("FLUORURO")
                            || nameParameter.toUpperCase().contains("FLUORUROS")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));  // dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCc_quim_inor_fluoruro(valueTemp);
                        gtzDestiny.setUm_quim_inor_fluoruro(valueUnitMeasure);
                    } else if (nameParameter.toUpperCase().contains("LANGELIER") || nameParameter.toUpperCase().contains("FLUORURO")) {
                        String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                        gtzDestiny.setCc_quim_inor_indice_langelier(valueTemp);
                        gtzDestiny.setUm_quim_inor_indice_langelier(valueUnitMeasure);
                    } else //TODO: Nitritos esta en control basico de excel
                        if (nameParameter.toUpperCase().contains("NITRITO") || nameParameter.toUpperCase().contains("NITRITOS")) {
                            String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                            gtzDestiny.setCb_quim_nitritos(valueTemp);
                            gtzDestiny.setUm_quim_nitritos(valueUnitMeasure);
                        } else//TODO: Nitratos esta en control basico de excel
                            if (nameParameter.toUpperCase().contains("NITRATO") || nameParameter.toUpperCase().contains("NITRATOS")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCb_quim_nitratos(valueTemp);
                                gtzDestiny.setUm_quim_nitritos(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("PLOMO")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_quim_inor_plomo(valueTemp);
                                gtzDestiny.setUm_quim_inor_plomo(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("ZINC")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_quim_inor_zinc(valueTemp);
                                gtzDestiny.setUm_quim_inor_zinc(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("COLFORMES TOTALES") || nameParameter.toUpperCase().contains("COLIFORMES TOTALES")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_biobac_coliformes_totales(valueTemp);
                                gtzDestiny.setUm_biobac_coliformes_totales(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("SODIO")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_quim_inor_sodio(valueTemp);
                                gtzDestiny.setUm_quim_inor_sodio(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("COLIFORMES TERMO TOLERANTE") || nameParameter.toUpperCase().contains("COLIFORME TERMO TOLERANTE")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_quim_inor_sodio(valueTemp);
                                gtzDestiny.setUm_quim_inor_sodio(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("ESCHERICHIA") || nameParameter.toUpperCase().contains("ECHERICHIA")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));// dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_biobac_escherichia_coli(valueTemp);
                                gtzDestiny.setUm_biobac_escherichia_coli(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("PSEDONOMAS") || nameParameter.toUpperCase().contains("AEROGINOSA")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_biobac_pseudomonas_aeruginosa(valueTemp);
                                gtzDestiny.setUm_biobac_pseudomonas_aeruginosa(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("CLOSTRIDIUM") || nameParameter.toUpperCase().contains("PERFRINGENS")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_biobac_clostridium_perfringens(valueTemp);
                                gtzDestiny.setUm_biobac_clostridium_perfringens(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("GIARDIA") || nameParameter.toUpperCase().contains("GIARDIAS")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_micropar_giarda_ausencia(valueTemp);
                                gtzDestiny.setUm_micropar_giarda_ausencia(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("CRYPTOSPORIDIUM") || nameParameter.toUpperCase().contains("CRYPTOSPOR")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_micropar_cryptosporidium_ausencia(valueTemp);
                                gtzDestiny.setUm_micropar_cryptosporidium_ausencia(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("AMEBAS") || nameParameter.toUpperCase().contains("AMEBA")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_micropar_amebas_ausencia(valueTemp);
                                gtzDestiny.setUm_micropar_amebas_ausencia(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("CLOROFORMO") || nameParameter.toUpperCase().contains("CLOROFOR")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_thm_cloroformo(valueTemp);
                                gtzDestiny.setUm_thm_cloroformo(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("BROMOFORMO") || nameParameter.toUpperCase().contains("BROMOFOR")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f));// dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_thm_bromoformo(valueTemp);
                                gtzDestiny.setUm_thm_bromoformo(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("DICLOROMETANO") || nameParameter.toUpperCase().contains("DICLORO")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_thm_bromo_diclorometano(valueTemp);
                                gtzDestiny.setUm_thm_bromo_diclorometano(valueUnitMeasure);
                            } else if (nameParameter.toUpperCase().contains("DIBROMO") || nameParameter.toUpperCase().contains("DIBRO")) {
                                String valueTemp =  processCellValue(sheet.getRow(rvs + i).getCell(cvs + f)); //dataFormatter.formatCellValue(sheet.getRow(rvs + i).getCell(cvs + f));
                                gtzDestiny.setCc_thm_dibromo_clorometano(valueTemp);
                                gtzDestiny.setUm_thm_dibromo_clorometano(valueUnitMeasure);
                            }
                }

                gtzDestiny.setFecha_muestra(dateP);
                gtzDestiny.setId_tipo_punto_muestra(sampleName);
                gtzDestiny.setLugar_muestreo(samplingLocation);
                gtzDestiny.setSector_unidad_analisis(pointSample);
                gtzDestiny.setNombre_archivo_excel(archivo+"-"+conf.getNameSheet());

            }
            gtzDestinyList.add(gtzDestiny);
        }
        return gtzDestinyList;
    }

    private String processSample(String sampleName, String typeSample, String staticSampleName){
        if(typeSample.equals("simple")){
            return sampleName.replace("\n","");
        }else if(typeSample.equals("process")){
            return "Nombre no procesado";
        }else if(typeSample.equals("number")){
            return "Muestra-"+sampleName;
        }else if(typeSample.equals("static")){
            return staticSampleName;
        }

        return staticSampleName;
    }

    private String processCellValue(XSSFCell cell){
        FormulaEvaluator evaluator = myWorkBook.getCreationHelper().createFormulaEvaluator();
        CellValue cellValue = evaluator.evaluate(cell);
        String result = "";
        if(cell!=null && cell.getCellType() == CellType.FORMULA){
            switch (cell.getCachedFormulaResultType()) {
                case BOOLEAN:
//                    System.out.println(cell.getBooleanCellValue());
                    result = String.valueOf(cell.getBooleanCellValue());
                    break;
                case NUMERIC:
//                    System.out.println(cell.getNumericCellValue());
                    Double v = cell.getNumericCellValue();
                    result = String.format("%1$,.2f",v);
                    break;
                case STRING:
//                    System.out.println(cell.getRichStringCellValue());
                    result = cell.getStringCellValue();
                    break;
            }
        }else{
            DataFormatter dataFormatter = new DataFormatter();
            result = dataFormatter.formatCellValue(cell);
        }
        return result;
    }

    private String processDate(String date, String typeDate, String staticDate, String conditionProcess, String year){
        if(typeDate.equals("static")){
            return staticDate;
        }else if(typeDate.equals("process")){
            //implent with regex to remove characters from string

            return "pendiente process";
        }else if(typeDate.equals("semester")){
            if(date.contains("1") || date.contains("I") || date.toUpperCase().contains("PRIMER")){
                return "15/06/"+year;
            }else{
                return "15/12/"+year;
            }
        }else if(typeDate.equals("year")){
            return "15/12/"+year;
        }else if(typeDate.equals("month-day")){
            return date +"/" + year;
        }else if(typeDate.equals("month")){
            return "15/" +date +"/" + year;
        }else if(typeDate.equals("date")){
            return date;
        }else if(typeDate.equals("split1")){
            String v="";
            String dateStr="";
            try {
                String[] ar = conditionProcess.split(",");
                v = date.split(ar[0])[1].replace(ar[1], "");
                DateTimeFormatter dateTimeFormatterInput = DateTimeFormatter.ofPattern("dd/MM/yy");
                DateTimeFormatter dateTimeFormatterOutput = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                LocalDate localDate = LocalDate.parse(v, dateTimeFormatterInput);
                dateStr = localDate.format(dateTimeFormatterOutput);
            }catch(Exception e){
                return v;
            }
            return  dateStr;
        }else if(typeDate.equals("regex")){
            Pattern pattern = Pattern.compile("'^.*\\((.*)\\)$'");
            Matcher matcher = pattern.matcher(date);
            if(matcher.find()){
                return matcher.group(1);
            }
        }
        return "no identificado";
    }

    public void exportCVS(List<GtzDestiny> gtzDestinyList, String csvFile) throws IOException {
        ICsvBeanWriter beanWriter = null;
        try {
            CsvPreference PIPE_DELIMITED = new CsvPreference.Builder('"', '|', "\n").build();
            beanWriter = new CsvBeanWriter(new FileWriter(csvFile, StandardCharsets.ISO_8859_1),
                    PIPE_DELIMITED);
            final String[] header = new String[]{
                    "id_unico",
                    "id_epsa",
                    "nombre_archivo_excel",
                    "periodo_reporte",
                    "fecha_reporte",
                    "id_lectura",
                    "id_reg_muestra",
                    "unidad_analisis",
                    "sector_unidad_analisis",
                    "codigo_punto",
                    "lugar_muestreo",
                    "fecha_muestra",
                    "latitud_muestra",
                    "longitud_muestra",
                    "altura_muestra",
                    "id_tipo_punto_muestra",
                    "metodo_utilizado",
                    "unidad_muestra",
                    "valores_max_nb512",
                    "cm_ph",
                    "um_ph",
                    "cm_temperatura",
                    "um_temperatura",
                    "cm_conductividad",
                    "um_conductividad",
                    "cm_turbiedad",
                    "um_turbiedad",
                    "cm_cloro_residual",
                    "um_cloro_residual",
                    "cm_coliforme_termotolerante_u1",
                    "um_coliforme_termotolerante_u1",
                    "cm_escherichia_coli_u1",
                    "um_escherichia_coli_u1",
                    "cb_fis_color",
                    "um_fis_color",
                    "cb_fis_sabor_olor_aceptables",
                    "um_fis_sabor_olor_aceptables",
                    "cb_quim_solidos_disueltos_totales",
                    "um_quim_solidos_disueltos_totales",
                    "cb_quim_alcalinidad_total",
                    "um_quim_alcalinidad_total",
                    "cb_quim_calcio",
                    "um_quim_calcio",
                    "cb_quim_cloruros",
                    "um_quim_cloruros",
                    "cb_quim_dureza_total",
                    "um_quim_dureza_total",
                    "cb_quim_hierro_total",
                    "um_quim_hierro_total",
                    "cb_quim_magnesio",
                    "um_quim_magnesio",
                    "cb_quim_manganeso",
                    "um_quim_manganeso",
                    "cb_quim_nitritos",
                    "um_quim_nitritos",
                    "cb_quim_nitratos",
                    "um_quim_nitratos",
                    "cb_quim_sulfatos",
                    "um_quim_sulfatos",
                    "cb_micro_heterotroficas",
                    "um_micro_heterotroficas",
                    "cc_quim_inor_aluminio",
                    "um_quim_inor_aluminio",
                    "cc_quim_inor_amonio",
                    "um_quim_inor_amonio",
                    "cc_quim_inor_arsenico",
                    "um_quim_inor_arsenico",
                    "cc_quim_inor_boro",
                    "um_quim_inor_boro",
                    "cc_quim_inor_cadmio",
                    "um_quim_inor_cadmio",
                    "cc_quim_inor_cobre",
                    "um_quim_inor_cobre",
                    "cc_quim_inor_fluoruro",
                    "um_quim_inor_fluoruro",
                    "cc_quim_inor_indice_langelier",
                    "um_quim_inor_indice_langelier",
                    "cc_quim_inor_plomo",
                    "um_quim_inor_plomo",
                    "cc_quim_inor_sodio",
                    "um_quim_inor_sodio",
                    "cc_quim_inor_zinc",
                    "um_quim_inor_zinc",
                    "cc_biobac_coliformes_totales",
                    "um_biobac_coliformes_totales",
                    "cc_biobac",
                    "um_biobac",
                    "cc_biobac_coliformes_termotolerantes",
                    "um_biobac_coliformes_termotolerantes",
                    "cc_biobac_escherichia_coli",
                    "um_biobac_escherichia_coli",
                    "cc_biobac_pseudomonas_aeruginosa",
                    "um_biobac_pseudomonas_aeruginosa",
                    "cc_biobac_clostridium_perfringens",
                    "um_biobac_clostridium_perfringens",
                    "cc_micropar_giarda_ausencia",
                    "um_micropar_giarda_ausencia",
                    "cc_micropar_cryptosporidium_ausencia",
                    "um_micropar_cryptosporidium_ausencia",
                    "cc_micropar_amebas_ausencia",
                    "um_micropar_amebas_ausencia",
                    "cc_thm_cloroformo",
                    "um_thm_cloroformo",
                    "cc_thm_bromoformo",
                    "um_thm_bromoformo",
                    "cc_thm_bromo_diclorometano",
                    "um_thm_bromo_diclorometano",
                    "cc_thm_dibromo_clorometano",
                    "um_thm_dibromo_clorometano"

            };
            final CellProcessor[] processors = getProcessors();

            // write the header
            beanWriter.writeHeader(header);
            for (final GtzDestiny gtzDestiny : gtzDestinyList) {
                beanWriter.write(gtzDestiny, header, processors);
            }
        }finally{
            if( beanWriter != null ) {
                beanWriter.close();
            }
        }

    }

    private CellProcessor[] getProcessors(){
        final CellProcessor[] processors = new CellProcessor[] {
                new Optional(), // id_unico
                new Optional(), // id_epsa
                new Optional(), // nombre_archivo_excel
                new Optional(), // periodo_reporte
                new Optional(), // fecha_reporte
                new Optional(), // id_lectura
                new Optional(), // id_reg_muestra
                new Optional(), // unidad_analisis
                new Optional(), // sector_unidad_analisis
                new Optional(), // codigo_punto
                new Optional(), // lugar_muestreo
                new Optional(), // fecha_muestra
                new Optional(), // latitud_muestra
                new Optional(), // longitud_muestra
                new Optional(), // altura_muestra
                new Optional(), // id_tipo_punto_muestra
                new Optional(), // metodo_utilizado;
                new Optional(), // unidad_muestra;
                new Optional(), // valores_max_nb512;

                new Optional(), // cm_ph;
                new Optional(), // um_ph;
                new Optional(), // cm_temperatura;
                new Optional(), // um_temperatura;
                new Optional(), // cm_conductividad;
                new Optional(), // um_conductividad;
                new Optional(), // cm_turbiedad;
                new Optional(), // um_turbiedad;
                new Optional(), // cm_cloro_residual;
                new Optional(), // um_cloro_residual;
                new Optional(), // cm_coliforme_termotolerante;
                new Optional(), // um_coliforme_termotolerante;
                new Optional(), // cm_escherichia_coli_u1;
                new Optional(), // um_escherichia_coli_u1;

                new Optional(), // cb_fis_color;
                new Optional(), // um_fis_color;
                new Optional(), // cb_fis_sabor_olor_aceptables;
                new Optional(), // um_fis_sabor_olor_aceptables;
                new Optional(), // cb_quim_solidos_disueltos_totales;
                new Optional(), // um_quim_solidos_disueltos_totales;
                new Optional(), // cb_quim_alcalinidad_total;
                new Optional(), // um_quim_alcalinidad_total;
                new Optional(), // cb_quim_calcio;
                new Optional(), // um_quim_calcio;
                new Optional(), // cb_quim_cloruros;
                new Optional(), // um_quim_cloruros;
                new Optional(), // cb_quim_dureza_total;
                new Optional(), // um_quim_dureza_total;
                new Optional(), // cb_quim_hierro_total;
                new Optional(), // um_quim_hierro_total;
                new Optional(), // cb_quim_magnesio;
                new Optional(), // um_quim_magnesio;
                new Optional(), // cb_quim_manganeso;
                new Optional(), // um_quim_manganeso;
                new Optional(), // cb_quim_nitritos;
                new Optional(), // um_quim_nitritos;
                new Optional(), // cb_quim_nitratos;
                new Optional(), // um_quim_nitratos;
                new Optional(), // cb_quim_sulfatos;
                new Optional(), // um_quim_sulfatos;
                new Optional(), // cb_micro_heterotroficas;
                new Optional(), // um_micro_heterotroficas;

                new Optional(), // cc_quim_inor_aluminio;
                new Optional(), // um_quim_inor_aluminio;
                new Optional(), // cc_quim_inor_amonio;
                new Optional(), // um_quim_inor_amonio;
                new Optional(), // cc_quim_inor_arsenico;
                new Optional(), // um_quim_inor_arsenico;
                new Optional(), // cc_quim_inor_boro;
                new Optional(), // um_quim_inor_boro;
                new Optional(), // cc_quim_inor_cadmio;
                new Optional(), // um_quim_inor_cadmio;
                new Optional(), // cc_quim_inor_cobre;
                new Optional(), // um_quim_inor_cobre
                new Optional(), // cc_quim_inor_fluoruro;
                new Optional(), // um_quim_inor_fluoruro;
                new Optional(), // cc_quim_inor_indice_langelier;
                new Optional(), // um_quim_inor_indice_langelier;
                new Optional(), // cc_quim_inor_plomo;
                new Optional(), // um_quim_inor_plomo;
                new Optional(), // cc_quim_inor_sodio;
                new Optional(), // um_quim_inor_sodio;
                new Optional(), // cc_quim_inor_zinc;
                new Optional(), // um_quim_inor_zinc;
                new Optional(), // cc_biobac_coliformes_totales;
                new Optional(), // um_biobac_coliformes_totales;
                new Optional(), // cc_biobac;
                new Optional(), // um_biobac;
                new Optional(), // cc_biobac_coliformes_termotolerantes;
                new Optional(), // um_biobac_coliformes_termotolerantes;
                new Optional(), // cc_biobac_escherichia_coli;
                new Optional(), // um_biobac_escherichia_coli;
                new Optional(), // cc_biobac_pseudomonas_aeruginosa;
                new Optional(), // um_biobac_pseudomonas_aeruginosa;
                new Optional(), // cc_biobac_clostridium_perfringens;
                new Optional(), // um_biobac_clostridium_perfringens;
                new Optional(), // cc_micropar_giarda_ausencia;
                new Optional(), // um_micropar_giarda_ausencia;
                new Optional(), // cc_micropar_cryptosporidium_ausencia;
                new Optional(), // um_micropar_cryptosporidium_ausencia;
                new Optional(), // cc_micropar_amebas_ausencia;
                new Optional(), // um_micropar_amebas_ausencia;
                new Optional(), // cc_thm_cloroformo;
                new Optional(), // um_thm_cloroformo;
                new Optional(), // cc_thm_bromoformo;
                new Optional(), // um_thm_bromoformo;
                new Optional(), // cc_thm_bromo_diclorometano;
                new Optional(), // um_thm_bromo_diclorometano;
                new Optional(), // cc_thm_dibromo_clorometano;
                new Optional() // um_thm_dibromo_clorometano;


        };

        return processors;
    }

}
