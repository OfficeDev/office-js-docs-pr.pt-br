---
title: Visão geral da API JavaScript do Excel
description: Saiba mais sobre as APIs JavaScript do Excel
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 80340b4990b56b2ba4d51f2a028480af3e267828
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650804"
---
# <a name="excel-javascript-api-overview"></a>Visão geral da API JavaScript do Excel

Um suplemento do Excel interage com objetos no Excel usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:

* **API de JavaScript do Excel para**: estas são as [APIs específicas do aplicativo](../../develop/application-specific-api-model.md) para o Excel. Introduzida com o Office 2016, a [API de JavaScript do Excel](/javascript/api/excel) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.

* **APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.

Esta seção da documentação concentra-se na API JavaScript do Excel, que você usará para desenvolver a maior parte da funcionalidade em suplementos direcionados para o Excel na Web ou para o Excel 2016 ou posterior. Para obter mais informações do API comum, consulte [Modelo do objeto do JavaScript API comum](../../develop/office-javascript-api-object-model.md).

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
