package school.sptech;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LeitorExcel {

    public List<Livro> extrairLivros(String nomeArquivo) {
        List<Livro> livrosExtraidos = new ArrayList<>();

        try (
              InputStream arquivo = new FileInputStream(nomeArquivo);
              Workbook workbook = new XSSFWorkbook(arquivo) // caso seja .xls troque para HSSFWorkbook
        ) {

            System.out.printf("Iniciando leitura do arquivo %s%n", nomeArquivo);

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    printarCabecalho(row);
                    continue;
                }

                // Extraindo valor das células e criando objeto Livro
                System.out.println("Lendo linha " + row.getRowNum());

                Integer id = (int) row.getCell(0).getNumericCellValue();
                String titulo = row.getCell(1).getStringCellValue();
                String autor = row.getCell(2).getStringCellValue();
                LocalDate dataLancamento = row.getCell(3).getLocalDateTimeCellValue().toLocalDate();

                Livro livro = new Livro(id, titulo, autor, dataLancamento);
                livrosExtraidos.add(livro);
            }

            printarLinhas();
            System.out.println("Leitura do arquivo finalizada");
            printarLinhas();

            return livrosExtraidos;
        } catch (IOException e) {
            return livrosExtraidos;
        }
    }

    private void printarCabecalho(Row row) {
        printarLinhas();
        System.out.println("Lendo cabeçalho");
        for (int i = 0; i < 4; i++) {
            String coluna = row.getCell(i).getStringCellValue();
            System.out.println("Coluna " + i + ": " + coluna);
        }
        printarLinhas();
    }

    private void printarLinhas() {
        System.out.println("-".repeat(20));
    }
}
