---
title: Conceitos fundamentais de programação com a API JavaScript do Word
description: Use as APIs JavaScript do Word para criar suplementos para o Word.
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: 00a7405d4d89279049d2724dda4fa1384a88dca4
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/05/2019
ms.locfileid: "35576725"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a>Conceitos fundamentais de programação com a API JavaScript do Word

Este artigo descreve conceitos fundamentais para o uso da [API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md) para criar suplementos para o Word 2016 ou posterior.

## <a name="referencing-officejs"></a>Referenciando Office.js

Você pode obter referência do Office.js nos seguintes locais:

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`: use esse recurso para os suplementos de produção.

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use esse recurso para experimentar recursos de visualização.

## <a name="word-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Word

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office oferece suporte para as APIs necessárias para um suplemento. Para saber mais sobre conjuntos de requisitos da API JavaScript do Word, consulte conjuntos de requisitos da [API JavaScript do Word](../reference/requirement-sets/word-api-requirement-sets.md).

## <a name="running-word-add-ins"></a>Execução de suplementos do Word

Para executar o suplemento, use um manipulador de eventos **Office.initialize**. Confira [Entendendo a API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) para saber mais sobre a inicialização do suplemento.

Os suplementos direcionados ao Word 2016 ou posterior são executados passando uma função para o método **Word.run()**. A função passada para o método **run** deve ter um argumento de contexto. Esse [objeto de contexto](/javascript/api/word/word.requestcontext) é diferente do objeto de contexto obtido do objeto do Office, mas ele é usado para interagir com o ambiente de tempo de execução do Word. O objeto de contexto fornece acesso ao modelo de objeto da API JavaScript do Word. O exemplo a seguir mostra como iniciar e executar um suplemento do Word usando o método **Word.run()**.

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

### <a name="asynchronous-nature-of-word-apis"></a>Natureza assíncrona das APIs do Word

A API JavaScript do Word é carregada pelo Office.js. Ela muda a maneira de interagir com objetos, como documentos e parágrafos. Em vez de fornecer APIs assíncronas individuais para recuperar e atualizar cada um desses objetos, a API JavaScript do Word fornece objetos JavaScript “proxy” que correspondem aos objetos reais em execução no Word. Você pode interagir com esses objetos proxy ao ler e gravar, simultaneamente, suas propriedades e chamar, de forma simultânea, métodos para executar operações neles. Essas interações com objetos proxy não são percebidas imediatamente no script em execução. O método **context.sync** sincroniza o estado entre o JavaScript em execução e os objetos reais do Office, executando instruções na fila e recuperando propriedades de objetos carregados do Word para uso no seu script.

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>Sincronizar documentos do Word com objetos proxy da API JavaScript do Word

O modelo de objeto da API JavaScript do Word é combinado livremente com os objetos no Word. Os objetos da API JavaScript do Word são proxies de objetos em um documento do Word. As ações executadas em objetos proxy não são percebidas no Word até que o estado do documento seja sincronizado. Por outro lado, o estado do documento do Word não é percebido em objetos proxy, até que o estado do documento seja sincronizado. Para sincronizar o estado do documento, execute o método **context.sync()**. O exemplo a seguir mostra a criação de um objeto proxy do corpo e um comando na fila para carregar a propriedade de texto nesse objeto e usa o método **context.sync()** para sincronizar o corpo do documento do Word com o objeto proxy do corpo.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    body.load("text");

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>Execução de um lote de comandos

Os objetos proxy do Word têm métodos para acessar e atualizar o modelo de objeto. Esses métodos são executados sequencialmente na ordem em que foram enfileirados no lote. Todos os comandos na fila do lote são executados quando o método **context.sync()** é chamado.

O exemplo a seguir mostra como a fila de comandos funciona. Quando o método **context.sync()** é chamado, o comando para carregar o corpo de texto é executado no Word. Em seguida, ocorre o comando para inserir o texto no corpo do Word. Na sequência, os resultados são retornados ao objeto proxy do corpo. O valor da propriedade **body.text**, na API JavaScript do Word, é o valor do corpo do documento do Word, <u>antes</u> da inserção do texto no documento do Word.

```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    body.load("text");

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="see-also"></a>Confira também

- [Visão geral da API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md)
- [Criar seu primeiro suplemento do Word](../quickstarts/word-quickstart.md)
- [Tutorial de suplemento do Word](../tutorials/word-tutorial.md)
- [Referências da API JavaScript do Word](/javascript/api/word) 


