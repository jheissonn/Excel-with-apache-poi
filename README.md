# Usando apache poi

Este aplicativo Spring Boot tem dois endpoints para demonstrar o uso do Apache POI para gerar um arquivo Excel. Alguns logs
é fornecido para obter uma sensação de desempenho e escalabilidade.

## Exportação de Excel sem Streaming

Esse tipo de exportação retém a pasta de trabalho na memória até que ela seja gravada. Funciona bem para exportações menores.

**Nom-streaming**

- [http://localhost:8080/excel/non-streaming-excel?columns=10&rows=50000](http://localhost:8080/excel/non-streaming-excel)

## Streaming Excel Export

Esse tipo de exportação grava dados no disco e libera memória depois que um número X de linhas foi adicionado à pasta de trabalho.
O número de linhas antes da gravação do disco pode ser especificado. O padrão é 100. Os arquivos temporários são mesclados em um único Excel
Arquivo. Depois que todas as linhas foram exportadas, chamar "dispose" na pasta de trabalho remove todos os arquivos temporários e renderiza o
instância de pasta de trabalho inutilizável. Uma nova instância de pasta de trabalho precisaria ser criada para anexar.

Este tipo de exportação não parece suportar células rich text. No entanto, um único estilo pode ser aplicado a todo o conteúdo em
a célula muito bem. Eu adoraria ser provado errado sobre isso. Uma possível solução alternativa seria reabrir a pasta de trabalho
após o processamento inicial como uma pasta de trabalho sem streaming e adicione as células de rich text.

Este tipo de exportação é bom para grandes exportações.

**Streaming-EXCEL**

- [http://localhost:8080/excel/streaming-excel?columns=10&rows=50000](http://localhost:8080/excel/streaming-excel)


## Exemplos utilizando Células, Estilos e validação estilo dropdown