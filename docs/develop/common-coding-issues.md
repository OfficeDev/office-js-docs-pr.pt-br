---
title: Problemas de codificação comuns e comportamentos de plataforma inesperados
description: Uma lista de problemas da plataforma de API JavaScript do Office frequentemente encontrada pelos desenvolvedores.
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 8cea95e3214585ba8e0b77535916f9c564dde9df
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902120"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a>Problemas de codificação comuns e comportamentos de plataforma inesperados

Este artigo realça aspectos da API JavaScript do Office que podem resultar em comportamento inesperado ou exigir padrões de codificação específicos para obter o resultado desejado. Se você encontrar um problema que pertença à lista, informe-nos usando o formulário de comentários na parte inferior do artigo.

## <a name="some-properties-must-be-set-with-json-structs"></a>Algumas propriedades devem ser definidas com as estruturas JSON

> [!NOTE]
> Esta seção só se aplica às APIs específicas do host para Excel e Word.

Algumas propriedades devem ser definidas como estruturas JSON, em vez de definir suas subpropriedades individuais. Um exemplo disso é encontrado no [PageLayout](/javascript/api/excel/excel.pagelayout). A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui:

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

No exemplo anterior, você ***não*** poderá atribuir `zoom` um valor diretamente: `sheet.pageLayout.zoom.scale = 200;`. Essa instrução gera um erro porque `zoom` não está carregada. Mesmo que `zoom` fosse carregado, o conjunto de escala não terá efeito. Todas as operações de contexto `zoom`acontecem em, atualizando o objeto de proxy no suplemento e substituindo os valores definidos localmente.

Esse comportamento difere das [Propriedades de navegação](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , como [Range. Format](/javascript/api/excel/excel.range#format). As propriedades `format` de podem ser definidas usando a navegação de objeto, conforme mostrado aqui:

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Você pode identificar uma propriedade que deve ter suas subpropriedades definidas com uma estrutura JSON verificando seu modificador somente leitura. Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente. Propriedades graváveis como `PageLayout.zoom` devem ser definidas com uma estrutura JSON. Em Resumo:

- Propriedade somente leitura: as subpropriedades podem ser definidas por meio de navegação.
- Propriedade writable: as subpropriedades devem ser definidas com uma estrutura JSON (e não podem ser definidas por meio de navegação).

## <a name="setting-read-only-properties"></a>Configuração de propriedades somente leitura

As [definições do TypeScript](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura. Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro. O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a>Confira também

- [OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): o local para relatar e exibir problemas com a plataforma de suplementos do Office e APIs JavaScript.
- [Estouro de pilha](https://stackoverflow.com/questions/tagged/office-js): o local para solicitar e exibir perguntas de programação sobre as APIs JavaScript do Office. Certifique-se de aplicar a marca "Office-js" à sua pergunta ao postar no estouro de pilha.
- [UserVoice](https://officespdev.uservoice.com/): o local para sugerir novos recursos para a plataforma de suplementos do Office e APIs JavaScript do Office.
