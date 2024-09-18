package school.sptech;

import java.time.LocalDate;

public class Livro {

    private Integer id;
    private String titulo;
    private String autor;
    private LocalDate dataLancamento;

    public Livro() {
    }

    public Livro(Integer id, String nome, String autor, LocalDate dataLancamento) {
        this.id = id;
        this.titulo = nome;
        this.autor = autor;
        this.dataLancamento = dataLancamento;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getTitulo() {
        return titulo;
    }

    public void setTitulo(String titulo) {
        this.titulo = titulo;
    }

    public String getAutor() {
        return autor;
    }

    public void setAutor(String autor) {
        this.autor = autor;
    }

    public LocalDate getDataLancamento() {
        return dataLancamento;
    }

    public void setDataLancamento(LocalDate dataLancamento) {
        this.dataLancamento = dataLancamento;
    }

    @Override
    public String toString() {
        return "Livro{" +
              "id=" + id +
              ", titulo='" + titulo + '\'' +
              ", autor='" + autor + '\'' +
              ", dataLancamento=" + dataLancamento +
              '}';
    }
}
