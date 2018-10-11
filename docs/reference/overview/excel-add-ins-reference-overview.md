# <a name="excel-javascript-api-overview"></a>Visão geral da API JavaScript do Excel

Você pode usar a API JavaScript do Excel para criar suplementos para o Excel 2016. A lista a seguir mostra os objetos de alto nível do Excel que estão disponíveis na API. Os links de página dos objetos contêm uma descrição das respectivas propriedades, relações e métodos disponíveis. Acesse os links no menu para saber mais.

Alguns dos principais objetos do Excel são listados abaixo por conveniência: 

- [Workbook](/javascript/api/excel/excel.workbook): o objeto de nível superior que inclui os objetos da pasta de trabalho relacionada, como planilhas, tabelas, intervalos, etc. Você pode usá-lo também para enumerar as referências relacionadas.

- [Worksheet](/javascript/api/excel/excel.worksheet): representa uma planilha em uma pasta de trabalho. 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): uma coleção de objetos **Worksheet** em uma pasta de trabalho.

- [Range](/javascript/api/excel/excel.range): representa uma célula, uma linha, uma coluna ou uma seleção de células contendo um ou mais blocos contíguos de células.

- [Table](/javascript/api/excel/excel.table): representa uma coleção de células organizadas, projetada para facilitar o gerenciamento dos dados.
    - [TableCollection](/javascript/api/excel/excel.tablecollection): uma coleção de tabelas em uma pasta de trabalho ou planilha.
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): uma coleção de todas as colunas em uma tabela.
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection): uma coleção de todas as linhas em uma tabela.

- [Chart](/javascript/api/excel/excel.chart): representa um objeto chart em uma planilha, que é uma representação visual dos dados subjacentes.
    - [ChartCollection](/javascript/api/excel/excel.chartcollection): uma coleção de gráficos em uma planilha.

- [TableSort](/javascript/api/excel/excel.tablesort): representa um objeto que gerencia as operações de classificação em objetos **Table**.

- [RangeSort](/javascript/api/excel/excel.rangesort): representa um objeto que gerencia as operações de classificação em objetos **Range**.

- [Filter](/javascript/api/excel/excel.filter): representa um objeto que gerencia a filtragem da coluna de uma tabela.

- [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): representa a proteção de um objeto **Worksheet**.

- [NamedItem](/javascript/api/excel/excel.nameditem): representa um nome definido de um intervalo de células ou de um valor. 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): uma coleção dos objetos **NamedItem** em uma pasta de trabalho.

- [Binding](/javascript/api/excel/excel.binding): uma classe abstrata que representa uma associação a uma seção da pasta de trabalho.
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection): uma coleção dos objetos **Binding** em uma pasta de trabalho.

## <a name="excel-javascript-api-open-specifications"></a>Especificações abertas da API JavaScript do Excel

À medida que projetamos e desenvolvemos novas APIs para suplementos do Excel,  elas ficam disponíveis para seus comentários na nossa página [Especificação da API aberta](../openspec.md) . Descubra quais são os novos recursos no pipeline das APIs JavaScript do Excel e forneça sua opinião sobre nossas especificações de design.

## <a name="excel-javascript-api-reference"></a>Referência da API JavaScript do Excel

Para obter informações detalhadas sobre a API JavaScript do Excel, consulte a [documentação de referência de API do JavaScript de Excel](/javascript/api/excel).

## <a name="see-also"></a>Confira também

- [Visão geral dos suplementos do Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Visão geral da plataforma de suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Exemplos de suplementos do Excel no GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
