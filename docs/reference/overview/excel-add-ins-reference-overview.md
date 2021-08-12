---
title: Visão geral da API JavaScript do Excel
description: Saiba mais sobre as APIs JavaScript do Excel
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 1177cd4837e5cbebfc6cef533932486612d273c2a489fcd8f26b37b90726ce56
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085029"
---
# <a name="excel-javascript-api-overview"></a>Visão geral da API JavaScript do Excel

Um suplemento do Excel interage com objetos no Excel usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:

* **API JavaScript do Excel:** São as [APIs específicas do aplicativo](../../develop/application-specific-api-model.md) para Excel. Introduzida com o Office 2016, a [API JavaScript do Excel](/javascript/api/excel) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.

* **APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.

Esta seção da documentação se concentra na API JavaScript do Excel, que você usará para desenvolver a maioria das funcionalidades em suplementos destinados ao Excel na Web ou Excel 2016 ou posterior. Para obter informações sobre a API Comum, consulte [Modelo de objeto comum de API JavaScript](../../develop/office-javascript-api-object-model.md).

## <a name="learn-object-model-concepts"></a>Aprender os conceitos do modelo de objeto

Confira o [Modelo de objeto JavaScript do Excel em suplementos do Office](../../excel/excel-add-ins-core-concepts.md) para obter informações sobre conceitos importantes do modelo de objeto.

Para ter a experiência prática com o uso da API de JavaScript do Excel para acessar objetos no Excel, conclua o [Tutorial do suplemento do Excel](../../tutorials/excel-tutorial.md).

## <a name="learn-api-capabilities"></a>Conheça os recursos da API

Cada recurso principal da API do Excel possui um artigo ou conjunto de artigos explorando o que esse recurso pode fazer e o modelo de objeto relevante.

* [Gráficos](../../excel/excel-add-ins-charts.md)
* [Comentário](../../excel/excel-add-ins-comments.md)
* [Formatação condicional](../../excel/excel-add-ins-conditional-formatting.md)
* [Funções personalizadas](../../excel/custom-functions-overview.md)
* [Validação de dados](../../excel/excel-add-ins-data-validation.md)
* [Eventos](../../excel/excel-add-ins-events.md)
* [Tabelas Dinâmicas](../../excel/excel-add-ins-pivottables.md)
* [Faixas](../../excel/excel-add-ins-ranges-get.md) e [Células](../../excel/excel-add-ins-cells.md)
* [RangeAreas (vários intervalos)](../../excel/excel-add-ins-multiple-ranges.md)
* [Formas](../../excel/excel-add-ins-shapes.md)
* [Tabelas](../../excel/excel-add-ins-tables.md)
* [Pastas de trabalho e APIs no Nível do Aplicativo](../../excel/excel-add-ins-workbooks.md)
* [Planilhas](../../excel/excel-add-ins-worksheets.md)

Para saber mais sobre o modelo de objeto API JavaScript do Excel, consulte a [Documentação de referência da API JavaScript do Excel](/javascript/api/excel).

## <a name="try-out-code-samples-in-script-lab"></a>Experimente amostras de código no Script Lab

Use o [Script Lab](../../overview/explore-with-script-lab.md) para começar a trabalhar rapidamente com um conjunto de exemplos internos que mostram como concluir tarefas com a API. Você pode executar as amostras no Script Lab para ver instantaneamente o resultado no painel de tarefas ou planilha, examinar os exemplos para saber como a API funciona e até mesmo usar amostras para criar um protótipo do seu próprio suplemento.

## <a name="see-also"></a>Confira também

* [Documentação de Suplementos do Excel](../../excel/index.yml)
* [Visão geral dos suplementos do Excel](../../excel/excel-add-ins-overview.md)
* [Referência da API JavaScript do Excel](/javascript/api/excel)
* [Disponibilidade de aplicativos e plataformas de cliente Office para Suplementos do Office](../../overview/office-add-in-availability.md)
* [Usando o modelo de API específica do aplicativo](../../develop/application-specific-api-model.md)
