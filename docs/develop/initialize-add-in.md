---
title: Inicialize seu suplemento do Office
description: Saiba como inicializar o suplemento do Office.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 5dc9d0143ac9eaab18625e280891bd601fa9f899
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293321"
---
# <a name="initialize-your-office-add-in"></a>Inicialize seu suplemento do Office

Os Suplementos do Office têm sempre uma lógica de inicialização para fazer coisas como:

- Verifique se a versão do Office do usuário é compatível com todas as APIs do Office que seu código chama.

- Certifique-se de que a existência de determinados artefatos, como uma planilha com um nome específico.

- Solicita que o usuário selecione algumas células no Excel e, em seguida, insira um gráfico inicializado com os valores selecionados.

- Estabeleça associações.

- Use a API de caixa de diálogo do Office para solicitar ao usuário os valores padrão das configurações do suplemento.

No entanto, um suplemento do Office não pode chamar com êxito nenhuma API JavaScript do Office até que a biblioteca seja carregada. Este artigo descreve as duas maneiras pelas quais o código pode garantir que a biblioteca tenha sido carregada:

- Inicializar `Office.onReady()` .
- Inicializar `Office.initialize` .

> [!TIP]
> Recomendamos que use `Office.onReady()`em vez de`Office.initialize`. Embora `Office.initialize` ainda tenha suporte, o `Office.onReady()` oferece mais flexibilidade. Você pode atribuir apenas um manipulador ao `Office.initialize` e ele é chamado apenas uma vez pela infraestrutura do Office. Você pode chamar `Office.onReady()` diferentes locais no seu código e usar retornos de chamada diferentes.
> 
> Para saber mais sobre as diferenças entre essas técnicas, veja [Principais diferenças entre Office.initialize e Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).

Para saber mais sobre a sequência de eventos na inicialização do suplemento, confira [Carregar o ambiente de tempo de execução e o DOM](loading-the-dom-and-runtime-environment.md).

## <a name="initialize-with-officeonready"></a>Inicializar com o Office.onReady()

`Office.onReady()` é um método assíncrono que retorna um objeto [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) enquanto verifica se a biblioteca de Office.js foi carregada. Quando a biblioteca é carregada, ela resolve a promessa como um objeto que especifica o aplicativo cliente do Office com um `Office.HostType` valor de enumeração ( `Excel` , `Word` etc.) e a plataforma com um `Office.PlatformType` valor de enumeração ( `PC` , `Mac` , `OfficeOnline` etc.). O Promise será resolvido imediatamente quando a biblioteca estiver carregada ao`Office.onReady()` ser chamada.

Uma maneira de chamar `Office.onReady()` é passá-la por um método de retorno de chamada. Exemplo:

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

Este é o mesmo exemplo que usa as palavras-chave `async` e `await` em TypeScript:

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

Se estiver usando estruturas JavaScript adicionais que incluam testes e manipuladores próprios de inicialização, *geralmente* eles devem ser colocados dentro da resposta para `Office.onReady()`. Por exemplo, a função [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

No entanto, há exceções a essa prática. Por exemplo, suponha que você deseja abrir seu suplemento em um navegador (em vez de Sideload-lo em um aplicativo do Office) para depurar sua interface do usuário com ferramentas de navegador. Já que o Office.js não será carregado no navegador, `onReady` não será executado e o `$(document).ready` não será executado quando chamado dentro de `onReady` no Office. 

Outra exceção seria se você quiser que um indicador de progresso seja exibido no painel de tarefas enquanto o suplemento estiver sendo carregado. Neste cenário, o código deve chamar o jQuery `ready` e usar seu retorno de chamada para renderizar o indicador de progresso. Em seguida, a chamada de retorno do Office `onReady` pode substituir o indicador de progresso com a interface do usuário final. 

## <a name="initialize-with-officeinitialize"></a>Inicializar com Office.initialize

Um evento de inicialização é disparado quando a biblioteca do Office.js está carregada e pronta para a interação com o usuário. É possível atribuir um manipulador ao `Office.initialize` que implementa a lógica de inicialização. Veja a seguir um exemplo que verifica se a versão do Excel do usuário é compatível com todas as APIs que o suplemento pode chamar.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Se você estiver usando estruturas JavaScript adicionais que incluam seu próprio manipulador de inicialização ou testes, elas *deverão ser* colocadas no `Office.initialize` evento (as exceções descritas na seção **inicializar com Office. onReady ()** anteriormente também serão aplicadas neste caso). Por exemplo, a função [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:

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
> Mesmo que não tenha uma lógica de inicialização, você deve atribuir ou chamar `Office.onReady()` uma função vazia para `Office.initialize` quando o JavaScript do suplemento for carregado. Algumas combinações de aplicativos e plataformas do Office não carregarão o painel de tarefas até que uma dessas situações aconteça. Os exemplos a seguir mostram essas duas abordagens.
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a>Confira também

- [Entendendo a API JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [Carregando o DOM e o ambiente de tempo de execução](loading-the-dom-and-runtime-environment.md)