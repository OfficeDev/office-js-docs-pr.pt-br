---
title: Conceitos fundamentais de programação com a API JavaScript do Word
description: Use as APIs JavaScript do Word para criar suplementos para o Word.
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: 1e7a90d4be378ed9b2c1f30ebebd4a0beec45a11
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293090"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a>Conceitos fundamentais de programação com a API JavaScript do Word

Este artigo descreve conceitos fundamentais para o uso da [API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md) para criar suplementos para o Word 2016 ou posterior.

## <a name="referencing-officejs"></a>Referenciando Office.js

Você pode obter referência do Office.js nos seguintes locais:

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`: use esse recurso para os suplementos de produção.

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use esse recurso para experimentar recursos de visualização.

## <a name="word-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Word

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office oferece suporte para as APIs necessárias para um suplemento. Para saber mais sobre conjuntos de requisitos da API JavaScript do Word, consulte conjuntos de requisitos da [API JavaScript do Word](../reference/requirement-sets/word-api-requirement-sets.md).

## <a name="running-word-add-ins"></a>Execução de suplementos do Word

Para executar seu suplemento, use um manipulador de eventos `Office.initialize`. Confira [Entendendo a API](../develop/understanding-the-javascript-api-for-office.md) para saber mais sobre a inicialização do suplemento.

Os suplementos que visam o Word 2016 ou posterior podem usar as APIs específicas do Word. Eles passam a lógica de interação do Word como uma função no método `Word.run()`. Confira [Usando o modelo de API específico do aplicativo](../develop/application-specific-api-model.md) para saber mais sobre como interagir com o documento do Word neste modelo de programação.

O exemplo a seguir mostra como inicializar e executar um suplemento do Word usando o método `Word.run()`.

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

## <a name="see-also"></a>Confira também

- [Visão geral da API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md)
- [Criar seu primeiro suplemento do Word](../quickstarts/word-quickstart.md)
- [Tutorial de suplemento do Word](../tutorials/word-tutorial.md)
- [Referências da API JavaScript do Word](/javascript/api/word)
