package BMPPROAR.InspUsinagem;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.fasterxml.jackson.databind.ObjectMapper;

@Controller
@RequestMapping("/")
public class HomeController {

    private static final String FILE_PATH = "\\\\server-ad\\Publico\\Engenharia_qualidade\\CEPS\\";

    @GetMapping("/")
    public String showForm(Model model) {
        List<String> operadorList = Arrays.asList("", "Maicon", "Ezequias", "Rubens", "Pedro", "Andrei");
        List<String> opcoesList = Arrays.asList("", "Biglia", "Cope", "Integrex");
        List<String> ladosList = Arrays.asList("", "1° Lado", "2° Lado", "Lado Único");

        model.addAttribute("operador", operadorList);
        model.addAttribute("opcoes", opcoesList);
        model.addAttribute("lados", ladosList);

        return "index.html";
    }

    @PostMapping("/enviar")
    public String createAndSaveFile(@RequestParam("codigo") String codigo,
            @RequestParam("op") String op,
            @RequestParam("info") String infoJson,
            @RequestParam("operador") String operador,
            @RequestParam("opcoes") String opcoes,
            @RequestParam("lado") String lado) {
    	try {
    		ObjectMapper objectMapper = new ObjectMapper();
    		Double[][] infoMatrix = objectMapper.readValue(infoJson, Double[][].class);
            String filePath = FILE_PATH + opcoes + "\\" + codigo + "_" + op + ".xlsx";
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Dados");

            int rowNum = 0;
            XSSFRow row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue("Código");
            row.createCell(1).setCellValue("OP");
            row.createCell(2).setCellValue("Operador");
            row.createCell(3).setCellValue("Lado");

            row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(codigo);
            row.createCell(1).setCellValue(op);
            row.createCell(2).setCellValue(operador);
            row.createCell(3).setCellValue(lado);

            for (int i = 0; i < infoMatrix.length; i++) {
                row = sheet.createRow(rowNum++);
                for (int j = 0; j < infoMatrix[i].length; j++) {
                    row.createCell(j).setCellValue(infoMatrix[i][j]);
                }
            }

            try (OutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }

            System.out.println("Arquivo Excel salvo com sucesso: " + filePath);
        } catch (IOException e) {
            System.out.println("Erro ao salvar o arquivo: " + e.getMessage());
        }
        return "redirect:/";
    }
}
