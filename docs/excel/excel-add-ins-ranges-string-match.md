---
title: Encontre uma cadeia de caracteres usando a EXCEL JavaScript
description: Saiba como encontrar uma cadeia de caracteres em um intervalo usando Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: efd2671781a8ce8d3e8aeda88f87abb3ad5058a35878f28f47f50305cff1b038
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087393"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a>Encontre uma cadeia de caracteres dentro de um intervalo usando a EXCEL JavaScript

Este artigo fornece um exemplo de código que localiza uma cadeia de caracteres dentro de um intervalo usando Excel API JavaScript. Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a>Corresponder a uma cadeia de caracteres dentro de um intervalo

O objeto `Range` tem um método `find` para pesquisar uma cadeia especificada dentro do intervalo. Ele retorna o intervalo da primeira célula com o texto correspondente.

O exemplo de código a seguir localiza a primeira célula com um valor igual à cadeia de caracteres **Alimentos** e registra o seu endereço no console. Observe que `find` exibe um erro `ItemNotFound` se a cadeia de caracteres especificada não existir no intervalo. Se você acha que a cadeia de caracteres especificada pode não estar no intervalo, use o método [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) para que seu código manipule normalmente esse cenário.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

Quando o método `find` é chamado em um intervalo que representa uma única célula, a planilha inteira é pesquisada. A pesquisa começa na célula e segue na direção especificada pelo `SearchCriteria.searchDirection`, envolvendo as extremidades da planilha, se necessário.

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Encontre células especiais em um intervalo usando a API JavaScript Excel JavaScript](excel-add-ins-ranges-special-cells.md)
