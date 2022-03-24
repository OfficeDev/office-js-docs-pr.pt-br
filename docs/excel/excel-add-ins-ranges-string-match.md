---
title: Encontre uma cadeia de caracteres usando Excel API JavaScript
description: Saiba como encontrar uma cadeia de caracteres dentro de um intervalo usando Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 68b4e206c2592a4a96f5eaeb4c49aef6c0a42aad
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745484"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a>Encontre uma cadeia de caracteres dentro de um intervalo usando a EXCEL JavaScript

Este artigo fornece um exemplo de código que localiza uma cadeia de caracteres dentro de um intervalo usando Excel API JavaScript. Para ver a lista completa de propriedades e métodos compatíveis `Range` com o objeto, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a>Corresponder a uma cadeia de caracteres dentro de um intervalo

O objeto `Range` tem um método `find` para pesquisar uma cadeia especificada dentro do intervalo. Ele retorna o intervalo da primeira célula com o texto correspondente.

O exemplo de código a seguir localiza a primeira célula com um valor igual à cadeia de caracteres **Alimentos** e registra o seu endereço no console. Observe que `find` exibe um erro `ItemNotFound` se a cadeia de caracteres especificada não existir no intervalo. Se você acha que a cadeia de caracteres especificada pode não estar no intervalo, use o método [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) para que seu código manipule normalmente esse cenário.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let table = sheet.tables.getItem("ExpensesTable");
    let searchRange = table.getRange();
    let foundRange = searchRange.find("Food", {
        completeMatch: true, // Match the whole cell value.
        matchCase: false, // Don't match case.
        searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address");
    await context.sync();

    console.log(foundRange.address);
});
```

Quando o método `find` é chamado em um intervalo que representa uma única célula, a planilha inteira é pesquisada. A pesquisa começa na célula e segue na direção especificada pelo `SearchCriteria.searchDirection`, envolvendo as extremidades da planilha, se necessário.

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Encontre células especiais em um intervalo usando a api javascript Excel javascript](excel-add-ins-ranges-special-cells.md)
