package com.sipcommb.convertir.Controller;

import com.sipcommb.convertir.Model.Datos;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@Controller
public class FileController {


    /**
     * Platilla de inicio
     * @return
     */
    @GetMapping("/")
    public String Inicio() {
        return "index";
    }

    /**
     * Redirigir al usuario al inicio de la aplicación
     *
     * @return
     */
    @RequestMapping("*")
    public String manejarSolicitudNoEncontrada() {
        // Redirigir al usuario al inicio de la aplicación
        return "redirect:/";
    }


    /**
     * Lee el archivo y agrupa los datos
     *
     * @param neogate:String nombre del negate
     * @param archivo:File   el archivo Columna 1: Apartamento, Columna 2: Telefono
     * @throws IOException
     * @return: Archivo txt
     */

    @RequestMapping(value = "/archivo", method = RequestMethod.POST)
    public ResponseEntity<?> procesarArchivo(@RequestParam("neogate") String neogate, @RequestParam("file") MultipartFile archivo) throws IOException {
        List<Datos> datosList = new ArrayList<>();
        try (InputStream is = archivo.getInputStream()) {

            Workbook workbook = new XSSFWorkbook(is);
            Sheet sheet = workbook.getSheetAt(0);
            //Recorre las filas
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                // Obtener la fila por su índice
                Row row = sheet.getRow(rowIndex);
                Datos datos = new Datos();
                boolean bandera = false;//si cell es diferente a null y
                for (int columnIndex = 0; columnIndex <= 1; columnIndex++) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null) {
                        if (!Objects.equals(cell.toString(), "")) {
                            bandera = true;
                            switch (columnIndex) {
                                case 0:
                                    String apartamento=null;
                                    if (Objects.requireNonNull(cell.getCellType()) == CellType.NUMERIC) {
                                        Long numero = (long) cell.getNumericCellValue();
                                        apartamento = String.valueOf(numero);
                                    } else {
                                        apartamento = cell.toString();
                                    }
                                    datos.setApartamento(apartamento);
                                    break;
                                case 1:
                                    Long telefono = (long) cell.getNumericCellValue();
                                    datos.setTelefono(telefono);
                                    break;
                            }
                        }

                    }
                }
                if (bandera)
                    datosList.add(datos);
            }
        } catch (IOException e) {

            e.printStackTrace();
        }

        return exportarTxt(datosList, neogate);
    }

    /**
     * Crea el documento txt y agrega los datos necesarios
     *
     * @param datos:datos    grupados
     * @param neogate:nombre del nepgete
     * @return
     * @throws IOException
     */
    public ResponseEntity<?> exportarTxt(List<Datos> datos, String neogate) throws IOException {

        File file = File.createTempFile("exten", ".txt");
        try (FileOutputStream fos = new FileOutputStream(file)) {


            StringBuilder sb = new StringBuilder();
            for (Datos empleado : datos) {
                sb.append("exten =>_")
                        .append(empleado.getApartamento())
                        .append(",1,ResetCDR()\n")
                        .append("exten =>_")
                        .append(empleado.getApartamento())
                        .append(",2,SetAMAFlags(billing)\n")
                        .append("exten =>_")
                        .append(empleado.getApartamento())
                        .append(",3,NoOp;SetVar(c_number=${primero}${EXTEN})\n")
                        .append("exten =>_")
                        .append(empleado.getApartamento())
                        .append(",4,Dial(SIP/Troncal_" + neogate + "/")
                        .append(empleado.getTelefono())
                        .append("${EXTEN:4})\n\n");
            }
            fos.close();
            byte[] contenido = sb.toString().getBytes(StandardCharsets.UTF_8);
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.TEXT_PLAIN);
            headers.setContentDispositionFormData("attachment", "exten.txt");
            headers.setContentLength(contenido.length);
            return new ResponseEntity<>(contenido, headers, HttpStatus.OK);

        } catch (IOException e) {
            // Manejo de excepciones de E/S

            return new ResponseEntity<>("error", HttpStatus.BAD_REQUEST);
        }

    }


    @ExceptionHandler(Exception.class)
    @ResponseStatus(HttpStatus.INTERNAL_SERVER_ERROR)
    public ModelAndView manejarExcepcion() {
        ModelAndView mav = new ModelAndView("error");
        mav.addObject("mensaje", "Verifique el formato del archivo, si el error persiste contacte se con el administrador");
        return mav;


    }

}
