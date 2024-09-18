# Exemplo de leitura de arquivos excel utilizando a biblioteca Apache POI

## ⚙️ Configuração de Ambiente:

### Passo a passo da execução:

1. Clonar Repositório
2. Abrir projeto no IntelliJ ou sua IDE de preferencia
3. Executar projeto na IDE

## 📚 O que é o Apache POI?

Apache POI é uma biblioteca Java desenvolvida pela Apache Software Foundation que fornece suporte
para formatos de arquivos do Microsoft Office, como .doc, .docx, .xls, .xlsx, etc.

Neste exemplo, vamos utilizar a biblioteca Apache POI para ler arquivos Excel (.xlsx e .xls).

---

## 💡Lendo arquivos Excel com Apache POI

### 1. Adicione as dependências Maven necessárias ao seu arquivo `pom.xml`:

```xml

<dependencies>
  <dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>5.3.0</version>
  </dependency>
  <dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.3.0</version>
  </dependency>
</dependencies>
```

### 2. Instancie um objeto Workbook

Para ler nosso arquivo excel, precisamos instanciar um
objeto `XSSFWorkbook`. Ele representará o arquivo Excel que queremos ler no Java.

Devemos passar o arquivo que estamos lendo no construtor da classe. Esse arquivo é representado como
objeto `InputStream`.

```java
Path caminho = Path.of("melhores-livros.xlsx");
InputStream arquivo = Files.newInputStream(caminho);

Workbook workbook = new XSSFWorkbook(arquivo);
```

### 3. Manipule os dados do Workbook

Um arquivo excel comúm, possui planilhas, linhas e células. Assim também é com o objeto Workbook.

Para acessar uma planilha específica, utilizamos o método `getSheetAt()` passando o índice da
planilha.

**OBS:** O índice começa em 0!

```java
Sheet sheet = workbook.getSheetAt(0);
```

Para acessar as linhas e colunas, utilizamos o método `getRow()` e `getCell()`, passando também um
índice.

```java
Row row = sheet.getRow(0);
Cell cell = row.getCell(0);
```

Para pegar o valor de uma célula, podemos utilizar vários métodos. Temos métodos diferentes para
tipos de dados diferentes.

Por exemplo, se quisermos pegar o valor de uma célula como String, podemos utilizar o
método `getStringCellValue()`.

```java
String valor = cell.getStringCellValue();
```

### 4. Feche o Workbook

Quando terminar de ler o arquivo, é importante fechar o Workbook para liberar recursos.

```java
workbook.close();
```

<span style="color:#ffb02e;font-size:1.5em;">⚠️ IMPORTANTE!</span>

O Apache POI **NÃO** consegue ler arquivos .csv; É necessário
convertê-lo para .xlsx ou .xls.