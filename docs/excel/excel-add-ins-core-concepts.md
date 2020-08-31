---
title: Conceitos fundamentais de programação com a API JavaScript do Excel
description: Use a API JavaScript do Excel para criar suplementos para o Excel.
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: dde7dc66e0746fc4d9cf91ed3df824fab05c109d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292584"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a>Conceitos fundamentais de programação com a API JavaScript do Excel

Este artigo descreve como usar a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) para desenvolver suplementos para o Excel 2016 ou versões posteriores. Ele apresenta os conceitos básicos que são fundamentais para usar a API e fornece orientações para executar tarefas específicas, como leitura ou gravação em um intervalo grande, atualização de todas as células do intervalo e muito mais.

> [!IMPORTANT]
> Confira [Usar o modelo da API específica do aplicativo](../develop/application-specific-api-model.md) para saber mais sobre a natureza assíncrona das APIs do Excel e como elas funcionam com a pasta de trabalho.  

## <a name="officejs-apis-for-excel"></a>APIs Office.js para Excel

Um suplemento do Excel interage com objetos no Excel usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:

* **API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.

* **APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.

Enquanto você provavelmente use a API JavaScript do Excel para desenvolver a maioria das funcionalidades em suplementos que visam o Excel 2016, você também usará objetos na API comum. Por exemplo:

* [Contexto](/javascript/api/office/office.context): o objeto `Context` representa o ambiente de tempo de execução do suplemento e oferece acesso aos principais objetos da API. Ele consiste em detalhes da configuração da pasta de trabalho, como `contentLanguage` e `officeTheme`, além de fornecer informações sobre o ambiente de tempo de execução do suplemento, como `host` e `platform`. Além disso, ele fornece o método `requirements.isSetSupported()`, que você pode usar para verificar se o conjunto de requisitos especificado é suportado pelo aplicativo Excel onde o suplemento está sendo executado.
* [Documento](/javascript/api/office/office.document): o objeto `Document` fornece o método `getFileAsync()`, que você pode usar para baixar o arquivo do Excel em que o suplemento está sendo executado.

A imagem a seguir ilustra quando você pode usar a API JavaScript do Excel ou as APIs comuns.

![Imagem das diferentes entre a API JS do Excel e as APIs comuns](../images/excel-js-api-common-api.png)

## <a name="object-model"></a>Modelo de objetos

Para entender as APIs do Excel, você deve entender como os componentes de uma pasta de trabalho estão relacionados entre si.

* Uma **Pasta de trabalho** contém uma ou mais **Planilhas**.
* Uma **Planilha** concede acesso a células por meio de objetos de **Intervalo**.
* Um **Intervalo** representa um grupo de células contíguas.
* Os **Intervalos** são usados para criar e colocar **Tabelas**, **Gráficos**, **Formas** e outras visualizações de dados ou objetos da organização.
* Uma **Planilha** contém coleções desses objetos de dados que estão presentes na planilha individual.
* As **Pastas de trabalho** contêm coleções de alguns desses objetos de dados (por exemplo, **Tabelas**) para toda a **Pasta de trabalho**.

### <a name="ranges"></a>Intervalos

Um intervalo é um grupo de células contíguas na pasta de trabalho. Os suplementos costumam usar uma notação estilo A1 (por ex.: **B3** para a única célula na coluna **B** e linha **3** ou **C2:F4** para as células das colunas **C** a **F** e linhas **2** a **4**) para definir intervalos.

Os intervalos têm três propriedades principais: `values`, `formulas` e `format`. Essas propriedades recebem ou definem os valores da célula, as fórmulas a serem avaliadas e a formatação visual das células.

#### <a name="range-sample"></a>Exemplo de intervalo

O exemplo a seguir mostra como criar registros de vendas. Essa função usa objetos `Range` para definir os valores, fórmulas e formatos.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

Esse exemplo cria os seguintes dados na planilha atual:

![Um registro de vendas mostrando as linhas de valores, uma coluna de fórmulas e cabeçalhos formatados.](../images/excel-overview-range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Gráficos, tabelas e outros objetos de dados

As APIs JavaScript do Excel podem criar e manipular estruturas de dados e visualizações no Excel. As tabelas e gráficos são dois dos objetos mais usados, mas as APIs oferecem suporte a tabelas dinâmicas, formas, imagens e muito mais.

#### <a name="creating-a-table"></a>Criar uma tabela

Criar tabelas usando intervalos de dados preenchidos. Controles de formatação e tabela (por exemplo, filtros) são aplicados automaticamente ao intervalo.

O exemplo a seguir cria uma tabela usando os intervalos do exemplo anterior.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

Usar esse código de exemplo na planilha com os dados anteriores cria a tabela a seguir:

![Uma tabela criada a partir do registro de vendas anterior.](../images/excel-overview-table-sample.png)

#### <a name="creating-a-chart"></a>Criar um gráfico

Crie gráficos para visualizar os dados em um intervalo. As APIs suportam inúmeras variedades de gráficos que podem ser personalizadas de acordo com suas necessidades.

O exemplo a seguir cria um gráfico de colunas simples para três itens e o coloca 100 pixels abaixo da parte superior da planilha.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

Executar esse exemplo na planilha com a tabela anterior cria o seguinte gráfico:

![Um gráfico de colunas mostrando as quantidades de três itens do registro de vendas anterior.](../images/excel-overview-chart-sample.png)

## <a name="run-options"></a>Executar opções

`Excel.run` tem uma sobrecarga que recebe um objeto [RunOptions](/javascript/api/excel/excel.runoptions). Este contém um conjunto de propriedades que afetam o comportamento de plataforma quando a função é executada. A propriedade a seguir tem suporte no momento:

* `delayForCellEdit`: Determina se o Excel atrasa solicitação em lote até que o usuário sai do modo de edição de célula. Quando **verdadeira**, a solicitação em lote é atrasada e executada quando o usuário sai do modo de edição de célula. Quando **falsa**, a solicitação em lote falha automaticamente se o usuário está no modo de edição de célula (causando um erro para alcançar o usuário). O comportamento padrão sem nenhuma propriedade `delayForCellEdit` especificada é equivalente a quando é **falsa**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="null-or-blank-property-values"></a>valores de propriedade nulos ou em branco

`null` e as cadeias de caracteres esvaziadas têm implicações especiais nas APIs JavaScript do Excel. Elas são usadas para representar células vazias, sem formatação ou valores padrão. Essa seção detalha o uso da `null` e de uma cadeia de caracteres vazia ao obter e definir as propriedades.

### <a name="null-input-in-2-d-array"></a>entrada nula em uma matriz 2D

No Excel, um intervalo é representado por uma matriz 2D, onde a primeira dimensão é linhas e a segunda dimensão é colunas. Para definir valores, o formato do número ou a fórmula apenas para células específicas em um intervalo, especifique os valores, o formato do número ou a fórmula para essas células na matriz 2D, bem como `null` para todas as outras células na matriz 2D.

Por exemplo, para atualizar o formato do número apenas para uma célula em um intervalo e manter o formato de número existente para todas as outras células no intervalo, especifique o novo formato de número para a célula a ser atualizada e `null` para todas as outras células. O trecho de código a seguir define um novo formato de número para a quarta célula no intervalo e não altera o formato de número para as primeiras três células no intervalo.

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a>entrada nula para uma propriedade

`null` não é uma entrada válida para uma propriedade única. Por exemplo, o trecho de código a seguir não é válido, pois a propriedade `values` do intervalo não pode ser definida como `null`.

```js
range.values = null;
```

Da mesma forma, o seguinte snippet de código não é válido, pois `null` não é um valor válido para a propriedade `color`.

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a>Valores da propriedade nula na resposta

A formatação de propriedades como `size` e `color` conterá valores `null` na resposta quando valores diferentes existirem no intervalo especificado. Por exemplo, se você recuperar um intervalo e carregar sua propriedade `format.font.color`:

* Se todas as células no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especificará essa cor.
* Se houver várias cores de fonte dentro do intervalo, `range.format.font.color` será `null`.

### <a name="blank-input-for-a-property"></a>Entrada em branco para uma propriedade

Quando você especificar um valor em branco para uma propriedade (isto é, duas aspas sem espaço entre elas `''`), ele será interpretado como uma instrução para limpar ou redefinir a propriedade. Por exemplo:

* Se você especificar um valor em branco para a propriedade `values` de um intervalo, o conteúdo do intervalo será apagado.
* Se você especificar um valor em branco para a propriedade `numberFormat`, o formato de número será redefinido para `General`.
* Se você especificar um valor em branco para a propriedade `formula` e a propriedade `formulaLocale`, os valores de fórmula serão apagados.

### <a name="blank-property-values-in-the-response"></a>Valores da propriedade em branco na resposta

Para operações de leitura, um valor de propriedade em branco na resposta (isto é, duas aspas sem espaço entre elas `''`) indica que a célula não contém dados nem valor. No primeiro exemplo abaixo, a primeira e a última célula no intervalo não contêm dados. No segundo exemplo, as primeiras duas células no intervalo não contêm uma fórmula.

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="requirement-sets"></a>Conjuntos de requisitos

Os conjuntos de requisitos são grupos nomeados de membros da API. Um Suplemento do Office pode executar uma verificação de tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um aplicativo do Office dá suporte às APIs necessárias ao suplemento. Para identificar os conjuntos de requisitos específicos que estão disponíveis em cada plataforma suportada, confira [Conjuntos de requisitos da API JavaScript do Excel](../reference/requirement-sets/excel-api-requirement-sets.md).

### <a name="checking-for-requirement-set-support-at-runtime"></a>Verificando o suporte ao conjunto de requisitos no tempo de execução

O exemplo de código a seguir mostra como determinar se o aplicativo do Office, onde o suplemento está em execução, dá suporte ao conjunto de requisitos da API especificado.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Definindo o suporte ao conjunto de requisitos no manifesto

Você pode usar o [elemento Requirements](../reference/manifest/requirements.md) no manifesto do suplemento para especificar os conjuntos de requisitos mínimos e/ou os métodos de API exigidos pelo suplemento para ser ativado. Se a plataforma ou o aplicativo do Office não der suporte aos conjuntos de requisitos ou aos métodos de API que são especificados no `Requirements`elemento do manifesto, o suplemento não será executado nesse aplicativo ou plataforma e não será exibido na lista de suplementos que são mostrados em **Meus Suplementos**.

O exemplo de código a seguir mostra o elemento `Requirements` em um manifesto de suplemento que especifica se o suplemento deve ser carregado em todos os aplicativos cliente do Office que dão suporte ao conjunto de requisitos ExcelApi, versão 1.3 ou superior.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Para disponibilizar seu suplemento em todas as plataformas de um aplicativo do Office, como Excel Online, Windows e iPad, é recomendável verificar o suporte a requisitos no tempo de execução, em vez de definir o suporte ao conjunto de requisitos no manifesto.

### <a name="requirement-sets-for-the-officejs-common-api"></a>Conjuntos de requisitos para a API comum Office.js

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](../reference/requirement-sets/office-add-in-requirement-sets.md).

## <a name="handle-errors"></a>Lidar com erros

Quando ocorre um erro de API, a API retorna um objeto `error` que contém um código e uma mensagem. Para saber mais sobre o tratamento de erros, incluindo uma lista de erros da API, confira [Tratamento de erro](excel-add-ins-error-handling.md).

## <a name="see-also"></a>Confira também

* [Crie seu primeiro suplemento do Excel](../quickstarts/excel-quickstart-jquery.md)
* [Exemplos de código de suplementos do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Otimização de desempenho da API JavaScript do Excel](../excel/performance.md)
* [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
* [Problemas comuns de codificação e comportamentos inesperados da plataforma](../develop/common-coding-issues.md).
