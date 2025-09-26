package school.sptech;

import java.util.List;

public class Main {

    public static void main(String[] args) {
        String nomeArquivo = "melhores-livros.xlsx";

        // Extraindo os livros do arquivo
        LeitorExcel leitorExcel = new LeitorExcel();
        List<Livro> livrosExtraidos = leitorExcel.extrairLivros(nomeArquivo);

        System.out.println("Livros extraídos:");
        for (Livro livro : livrosExtraidos) {
            System.out.println(livro);
        }
    }
}