---
title: Visão geral da API JavaScript do Excel
description: ''
ms.date: 06/10/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: aa9574a93252c0011b211c39e37cc013beb64432
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910144"
---
# <a name="excel-javascript-api-overview"></a>Visão geral da API JavaScript do Excel

Você pode usar a API JavaScript do Excel para criar suplementos para o Excel 2016 ou posterior. A lista a seguir mostra os objetos de alto nível do Excel que estão disponíveis na API. Os links de página dos objetos contêm uma descrição dos respectivos eventos, propriedades e métodos que estão disponíveis no objeto. Acesse os links no menu para saber mais.

Alguns dos principais objetos do Excel são listados abaixo para conveniência:

- [Workbook](/javascript/api/excel/excel.workbook): o objeto de nível superior que inclui os objetos da pasta de trabalho relacionada, como planilhas, tabelas, intervalos, etc. Você pode usá-lo também para enumerar as referências relacionadas.

- [Worksheet](/javascript/api/excel/excel.worksheet): representa uma planilha em uma pasta de trabalho.
  - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): uma coleção de objetos **Worksheet** em uma pasta de trabalho.
  - [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): representa a proteção de um objeto **Worksheet**.

- [Range](/javascript/api/excel/excel.range): representa uma célula, uma linha, uma coluna ou uma seleção de células contendo um ou mais blocos contíguos de células.
  - [ConditionalFormat](/javascript/api/excel/excel.conditionalformat): um objeto que define uma regra e um formato aplicado ao intervalo quando a condição da regra for atendida.
  - [DataValidation](/javascript/api/excel/excel.datavalidation): um objeto que restringe a entrada do usuário a um intervalo baseado em diferentes critérios.
  - [RangeSort](/javascript/api/excel/excel.rangesort): representa um objeto que gerencia as operações de classificação em um intervalo.

- [Table](/javascript/api/excel/excel.table): representa uma coleção de células organizadas, projetada para facilitar o gerenciamento dos dados.
  - [TableCollection](/javascript/api/excel/excel.tablecollection): uma coleção de tabelas em uma pasta de trabalho ou planilha.
  - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): uma coleção de todas as colunas em uma tabela.
  - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection): uma coleção de todas as linhas de uma tabela.
  - [TableSort](/javascript/api/excel/excel.tablesort): representa um objeto que gerencia as operações de classificação em uma tabela.

- [Chart](/javascript/api/excel/excel.chart): representa um objeto de gráfico em uma planilha, que é uma representação visual dos dados subjacentes.
  - [ChartCollection](/javascript/api/excel/excel.chartcollection): uma coleção de gráficos em uma planilha.

- [PivotTable](/javascript/api/excel/excel.pivottable): representa uma Tabela Dinâmica do Excel, que é um agrupamento hierárquico e apresentação de dados.
  - [PivotTableCollection](/javascript/api/excel/excel.pivottablecollection): uma coleção de Tabelas Dinâmicas em uma planilha.

- [Filter](/javascript/api/excel/excel.filter): representa um objeto que gerencia a filtragem da coluna de uma tabela.

- [NamedItem](/javascript/api/excel/excel.nameditem): representa um nome definido de um intervalo de células ou de um valor.
  - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): uma coleção dos objetos **NamedItem** em uma pasta de trabalho.

- [Binding](/javascript/api/excel/excel.binding): Uma classe abstrata que representa uma associação a uma seção da pasta de trabalho.
  - [BindingCollection](/javascript/api/excel/excel.bindingcollection): uma coleção dos objetos **Binding** em uma pasta de trabalho.

## <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Excel

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office oferece suporte para as APIs necessárias para um suplemento. Para saber mais sobre conjuntos de requisitos da API JavaScript do Excel, consulte o artigo [Conjuntos de requisitos da API JavaScript do Excel](../requirement-sets/excel-api-requirement-sets.md).

## <a name="excel-javascript-api-reference"></a>Referência da API JavaScript do Excel

Para saber mais sobre a API JavaScript do Excel, consulte a [Documentação de referência da API JavaScript do Excel](/javascript/api/excel).

## <a name="see-also"></a>Confira também

- [Visão geral dos suplementos do Excel](/office/dev/add-ins/excel/excel-add-ins-overview)
- [Visão geral da plataforma Suplementos do Office](/office/dev/add-ins/overview/office-add-ins)
- [Exemplos de suplementos do Excel no GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
- [Especificações abertas da API](../openspec/openspec.md)
