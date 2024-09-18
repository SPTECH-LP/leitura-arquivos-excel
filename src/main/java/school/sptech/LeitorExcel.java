package school.sptech;

import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LeitorExcel {

    public List<Livro> extrarLivros(InputStream arquivo) {
        try {
            System.out.println("\nIniciando leitura do arquivo\n");

            // Criando um objeto Workbook a partir do arquivo recebido
            Workbook workbook = new XSSFWorkbook(arquivo);
            Sheet sheet = workbook.getSheetAt(0);

            List<Livro> livrosExtraidos = new ArrayList<>();

            // Iterando sobre as linhas do arquivo
            for (Row row : sheet) {

                // A primeira linha representa o cabeçalho!
                if (row.getRowNum() == 0) {
                    System.out.println("Recebendo cabeçalho");

                    for (int i = 0; i < 4; i++) {
                        System.out.print(row.getCell(i).getStringCellValue() + " | ");
                    }
                    System.out.println();
                    continue;
                }

                // Extraindo valor das células e criando objeto Livro
                System.out.println("Lendo linha " + row.getRowNum());
                Livro livro = new Livro();
                livro.setId((int) row.getCell(0).getNumericCellValue());
                livro.setTitulo(row.getCell(1).getStringCellValue());
                livro.setAutor(row.getCell(2).getStringCellValue());
                livro.setDataLancamento(converterDate(row.getCell(3).getDateCellValue()));

                livrosExtraidos.add(livro);
            }

            // Fechando o arquivo para liberar recursos
            workbook.close();
            System.out.println("\nLeitura do arquivo finalizada\n");

            return livrosExtraidos;
        } catch (IOException e) {
            // Caso ocorra algum erro durante a leitura do arquivo uma exceção será lançada
            throw new RuntimeException(e);
        }
    }

    private LocalDate converterDate(Date data) {
        return data.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }
}
