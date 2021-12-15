package com.nexion.process.views.processexcelfile;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.nexion.process.backend.Utils;
import com.nexion.process.backend.entity.GtzDestiny;
import com.nexion.process.views.MainLayout;
import com.vaadin.flow.component.button.Button;
import com.vaadin.flow.component.notification.Notification;
import com.vaadin.flow.component.orderedlayout.HorizontalLayout;
import com.vaadin.flow.component.orderedlayout.VerticalLayout;
import com.vaadin.flow.component.textfield.TextArea;
import com.vaadin.flow.component.textfield.TextField;
import com.vaadin.flow.router.PageTitle;
import com.vaadin.flow.router.Route;
import com.vaadin.flow.router.RouteAlias;
import com.wontlost.sweetalert2.Config;
import com.wontlost.sweetalert2.SweetAlert2Vaadin;

import java.io.IOException;
import java.util.List;

@PageTitle("Process Excel File")
@Route(value = "process", layout = MainLayout.class)
@RouteAlias(value = "", layout = MainLayout.class)
public class ProcessExcelFileView extends VerticalLayout {

    private TextField fileName;
    private TextField sheets;
    private TextField csvFile;
    private Button process;
    private TextArea textConfiguration;

    public ProcessExcelFileView() {
        fileName = new TextField("Archivo");
        fileName.setWidthFull();
        sheets = new TextField("Hojas");
        sheets.setWidthFull();
        csvFile = new TextField("Archivo destino");
        csvFile.setWidthFull();
        textConfiguration = new TextArea("Configuracion");
        textConfiguration.setWidthFull();
        textConfiguration.setHeight("500px");

        VerticalLayout layout = new VerticalLayout();
        layout.setHeight("550px");
        layout.add(textConfiguration);

        process = new Button("Procesar");
        process.addClickListener(e -> {
            Utils util = new Utils();
            Config config = new Config();
            try {
                List<GtzDestiny> result = util.processFile(fileName.getValue(),sheets.getValue(),textConfiguration.getValue());
                util.exportCVS(result,csvFile.getValue());

                config.setTitle("Información");
                config.setText("Archivo generado");
                config.setIcon("info");
                new SweetAlert2Vaadin(config).open();

            } catch (JsonProcessingException ex) {
                ex.printStackTrace();
                config.setTitle("Error");
                config.setText("Error al procesar el archivo");
                config.setIcon("error");
                new SweetAlert2Vaadin(config).open();
            } catch (IOException ioException) {
                config.setTitle("Error");
                config.setText("Error al crear/leer el archivo");
                config.setIcon("error");
                ioException.printStackTrace();
            }

        });


        add(fileName, sheets, csvFile, layout, process);
    }

}
