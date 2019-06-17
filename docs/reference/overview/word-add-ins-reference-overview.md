---
title: Visão geral da API JavaScript do Word
description: ''
ms.date: 06/10/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 92b66b98776c1ad6b2d824af8bf13b01f2807384
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910200"
---
# <a name="word-javascript-api-overview"></a>Visão geral da API JavaScript do Word

O Word fornece um conjunto sofisticado de APIs que você pode usar para criar suplementos que interajam com metadados e com o conteúdo do documento. Use essas APIs para criar experiências envolventes que integrem e estendam o Word. Você pode importar e exportar conteúdo, montar novos documentos de diferentes fontes de dados e se integrar com fluxos de trabalho do documento para criar soluções de documento personalizadas.

Você pode usar duas APIs JavaScript para interagir com metadados e objetos em um documento do Word:

- API JavaScript do Word – introduzida no Office 2016.
- [API JavaScript para Office](../javascript-api-for-office.md) (Office.js) – introduzida no Office 2013.

## <a name="word-javascript-api"></a>API JavaScript do Word

A API JavaScript do Word é carregada pelo Office.js. Ela muda a maneira de interagir com objetos, como documentos e parágrafos. Em vez de fornecer APIs assíncronas individuais para recuperar e atualizar cada um desses objetos, essa API fornece objetos JavaScript "proxy" que correspondem aos objetos reais em execução no Word. Você pode interagir com esses objetos proxy quando ler e gravar de forma síncrona as respectivas propriedades e quando chamar, também de forma síncrona, métodos para executar operações neles. Essas interações com objetos proxy não são percebidas imediatamente no script em execução. O método **context.sync** sincroniza o estado entre o JavaScript em execução e os objetos reais do Office, executando instruções na fila e recuperando propriedades de objetos carregados do Word para uso no seu script.

## <a name="javascript-api-for-office"></a>JavaScript API for Office

Você pode obter referência do Office.js nos seguintes locais:

- https://appsforoffice.microsoft.com/lib/1/hosted/office.js: use esse recurso para os suplementos de produção.
- https://appsforoffice.microsoft.com/lib/beta/hosted/office.js: use esse recurso quando estiver experimentando recursos de visualização.

Se estiver usando o [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), você poderá baixar as [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) para obter modelos de projeto que incluam o Office.js.  Você pode usar o [nuget para obter o Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).

Se você usar TypeScript e se tiver npm, poderá obter as definições de TypeScript ao digitar isto na interface da linha de comando: `typings install office-js --ambient`.

## <a name="running-word-add-ins"></a>Execução de suplementos do Word

Para executar o suplemento, use um manipulador de eventos Office.initialize. Consulte [Compreender a API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) para saber mais sobre a inicialização de suplementos.

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

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>Sincronizar documentos do Word com objetos proxy da API JavaScript do Word

O modelo de objeto da API JavaScript do Word é combinado livremente com os objetos no Word. Os objetos da API JavaScript do Word são proxies de objetos em um documento do Word. As ações executadas em objetos proxy não são percebidas no Word até que o estado do documento seja sincronizado. Por outro lado, o estado do documento do Word não é percebido em objetos proxy, até que o estado do documento seja sincronizado. Para sincronizar o estado do documento, execute o método **context.sync()**. O exemplo a seguir mostra a criação de um objeto proxy do corpo e um comando na fila para carregar a propriedade de texto nesse objeto e usa o método **context.sync()** para sincronizar o corpo do documento do Word com o objeto proxy do corpo.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>Execução de um lote de comandos

Os objetos proxy do Word dispõem de métodos para acessar e atualizar o modelo de objeto. Esses métodos são executados sequencialmente na ordem em que foram colocados na fila do lote. Todos os comandos na fila do lote são executados quando o método context.sync() é chamado.

O exemplo a seguir mostra como funciona a fila de comandos. Quando o método **context.sync()** é chamado, o comando para carregar o corpo de texto é executado no Word. Em seguida, ocorre o comando para inserir o texto no corpo do Word. Os resultados são retornados ao objeto proxy do corpo. O valor da propriedade **body.text**, na API JavaScript do Word, é o valor do corpo do documento do Word, <u>antes</u> da inserção do texto no documento do Word.

```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

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

## <a name="word-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Word

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office oferece suporte para as APIs necessárias para um suplemento. Para saber mais sobre conjuntos de requisitos da API JavaScript do Word, consulte o artigo [Conjuntos de requisitos da API JavaScript do Word](../requirement-sets/word-api-requirement-sets.md).

## <a name="word-javascript-api-reference"></a>Referências da API JavaScript do Word

Para saber mais sobre a API JavaScript do Word, consulte a [Documentação de referência da API JavaScript do Word](/javascript/api/word).

## <a name="see-also"></a>Confira também

- [Visão geral dos suplementos do Word](/office/dev/add-ins/word/word-add-ins-programming-overview)
- [Visão geral da plataforma Suplementos do Office](/office/dev/add-ins/overview/office-add-ins)
- [Exemplos de suplementos do Word no GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
- [Especificações abertas da API](../openspec/openspec.md)
