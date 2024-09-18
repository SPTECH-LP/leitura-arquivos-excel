package school.sptech;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

public class Main {

    public static void main(String[] args) throws IOException {
        // Carregando o arquivo de excel
        Path caminho = Path.of("melhores-livros.xlsx");
        InputStream arquivo = Files.newInputStream(caminho);

        // Extraindo os livros do arquivo
        LeitorExcel leitorExcel = new LeitorExcel();
        List<Livro> livrosExtraidos = leitorExcel.extrarLivros(arquivo);

        // Fechando o arquivo para liberar recursos
        arquivo.close();

        // Exibindo os livros extraídos
        System.out.println("Livros extraídos:");
        for (Livro livro : livrosExtraidos) {
            System.out.println(livro);
        }
    }
}