---
title: Inicialize seu suplemento do Office
description: Saiba como inicializar seu Suplemento do Office.
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: a809a353a54fbb7bd10f0d1d5920d8a6881d2a6f
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958668"
---
# <a name="initialize-your-office-add-in"></a>Inicialize seu suplemento do Office

Os Suplementos do Office têm sempre uma lógica de inicialização para fazer coisas como:

- Verifique se a versão do usuário do Office dá suporte a todas as APIs do Office que seu código chama.

- Verifique a existência de determinados artefatos, como uma planilha com um nome específico.

- Solicite que o usuário selecione algumas células no Excel e insira um gráfico inicializado com esses valores selecionados.

- Estabeleça associações.

- Use a API de Caixa de Diálogo do Office para solicitar ao usuário valores de configurações de suplemento padrão.

No entanto, um Suplemento do Office não pode chamar nenhuma APIs JavaScript do Office com êxito até que a biblioteca seja carregada. Este artigo descreve as duas maneiras pelas quais seu código pode garantir que a biblioteca tenha sido carregada.

- Inicializar com `Office.onReady()`.
- Inicializar com `Office.initialize`.

> [!TIP]
> Recomendamos que use `Office.onReady()`em vez de`Office.initialize`. Embora `Office.initialize` ainda tenha suporte, fornece `Office.onReady()` mais flexibilidade. Você pode atribuir apenas um manipulador e `Office.initialize` ele é chamado apenas uma vez pela infraestrutura do Office. Você pode chamar `Office.onReady()` em locais diferentes em seu código e usar retornos de chamada diferentes.
> 
> Para saber mais sobre as diferenças entre essas técnicas, veja [Principais diferenças entre Office.initialize e Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).

Para saber mais sobre a sequência de eventos na inicialização do suplemento, confira [Carregar o ambiente de tempo de execução e o DOM](loading-the-dom-and-runtime-environment.md).

## <a name="initialize-with-officeonready"></a>Inicializar com o Office.onReady()

`Office.onReady()` é uma função assíncrona que retorna um objeto [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) enquanto verifica se a biblioteca Office.js está carregada. Quando a biblioteca é carregada, ela resolve o Promise como um objeto que especifica o aplicativo cliente do Office `Office.HostType` com um valor de enumeração (`Excel`, `Word`etc.) `Office.PlatformType` e a plataforma com um valor de enumeração (`PC`, , `Mac`, `OfficeOnline`etc.). O Promise será resolvido imediatamente quando a biblioteca estiver carregada ao`Office.onReady()` ser chamada.

Uma maneira de chamar `Office.onReady()` é passar uma função de retorno de chamada. Veja um exemplo.

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

Como alternativa, é possível encadear um método `then()` à chamada de `Office.onReady()`, em vez de passar um retorno de chamada. Por exemplo, o código a seguir verifica se a versão do Excel do usuário é compatível com todas as APIs que o suplemento pode chamar.

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

Aqui está o mesmo exemplo usando as palavras-chave `async` e `await` o typescript.

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

Se estiver usando estruturas JavaScript adicionais que incluam testes e manipuladores próprios de inicialização, *geralmente* eles devem ser colocados dentro da resposta para `Office.onReady()`. Por exemplo, [o método JQuery](https://jquery.com) `$(document).ready()` seria referenciado da seguinte maneira:

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

No entanto, há exceções a essa prática. Por exemplo, suponha que você queira abrir seu suplemento em um navegador (em vez de fazer sideload dele em um aplicativo do Office) para depurar sua interface do usuário com ferramentas de navegador. Nesse cenário, uma vez Office.js determina que ele está em execução fora de um aplicativo host do Office, `null` ele chamará o retorno de chamada e resolverá a promessa com o host e a plataforma.

Outra exceção seria se você quiser que um indicador de progresso apareça no painel de tarefas enquanto o suplemento está sendo carregado. Nesse cenário, seu código deve chamar o jQuery `ready` e usar seu retorno de chamada para renderizar o indicador de progresso. Em seguida, `Office.onReady` o retorno de chamada pode substituir o indicador de progresso pela interface do usuário final.

## <a name="initialize-with-officeinitialize"></a>Inicializar com Office.initialize

Um evento de inicialização é disparado quando a biblioteca do Office.js está carregada e pronta para a interação com o usuário. É possível atribuir um manipulador ao `Office.initialize` que implementa a lógica de inicialização. Veja a seguir um exemplo que verifica se a versão do Excel do usuário é compatível com todas as APIs que o suplemento pode chamar.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Se você estiver usando estruturas JavaScript adicionais que incluem seu próprio manipulador de inicialização ou testes, `Office.initialize` elas geralmente devem ser colocadas dentro do evento (as exceções descritas na seção **Initialize com Office.onReady()** anteriormente também se aplicam nesse caso). Por exemplo, [o método JQuery](https://jquery.com) `$(document).ready()` seria referenciado da seguinte maneira:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

Para suplementos de conteúdo e painel de tarefas, `Office.initialize` fornece um parâmetro _reason_ adicional. Esse parâmetro especifica como um suplemento foi adicionado ao documento atual. Você pode usar isso para fornecer uma lógica diferente para quando um suplemento é inserido pela primeira vez, em comparação com quando já existia dentro do documento.

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

Para saber mais, veja [Evento Office.initialize](/javascript/api/office) e [Enumeração da InitializationReason](/javascript/api/office/office.initializationreason).

## <a name="major-differences-between-officeinitialize-and-officeonready"></a>Principais diferenças entre Office.initialize e Office.onReady

- É possível atribuir apenas um manipulador a `Office.initialize`, e ela é chamada apenas uma vez pela infraestrutura do Office, mas você pode chamar `Office.onReady()` em diferentes locais no código, e usar diferentes retornos de chamadas. Por exemplo, o código pode chamar `Office.onReady()`, logo que o script personalizado é carregado com um retorno de chamada que executa uma lógica de inicialização. Além disso, o código pode ter um botão no painel de tarefas, cujo script chama `Office.onReady()` com um retorno de chamada diferente. Quando isso ocorre, o segundo retorno de chamada é executado quando o botão é clicado.

- O evento `Office.initialize` é disparado no final do processo interno, e que o Office.js é inicializado automaticamente. Ele também é disparado *imediatamente* após o término do processo interno. Se o código no qual você atribui um manipulador ao evento for executado muito tempo após o evento ser disparado, então o manipulador não será executado. Por exemplo, se estiver usando o gerenciador de tarefas WebPack, ele poderá configurar a home page do suplemento para carregar arquivos de polyfill, após carregar o Office.js, mas antes de carregar o JavaScript personalizado. Quando o script carrega e atribui o manipulador, o evento de inicialização já ocorreu. Mas nunca é "tarde demais" para chamar `Office.onReady()`. Caso o evento de inicialização já tenha ocorrido, o retorno de chamada é executado imediatamente.

> [!NOTE]
> Mesmo que não tenha uma lógica de inicialização, você deve atribuir ou chamar `Office.onReady()` uma função vazia para `Office.initialize` quando o JavaScript do suplemento for carregado. Algumas combinações de aplicativos e plataformas do Office não carregarão o painel de tarefas até que uma delas ocorra. Os exemplos a seguir mostram essas duas abordagens.
>
>```js
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="debug-initialization"></a>Inicialização de depuração

Para obter informações sobre como depurar `Office.initialize` as `Office.onReady()` funções e as funções, consulte [Depurar as funções initialize e onReady](../testing/debug-initialize-onready.md).

## <a name="see-also"></a>Confira também

- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [Carregando o DOM e o ambiente de tempo de execução](loading-the-dom-and-runtime-environment.md)