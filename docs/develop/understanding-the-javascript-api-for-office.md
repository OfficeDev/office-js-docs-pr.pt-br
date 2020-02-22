---
title: Noções básicas da API JavaScript para Office
description: ''
ms.date: 06/21/2019
localization_priority: Priority
ms.openlocfilehash: a82437fc82d9c9a31e75d448579f37d440784aa2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163511"
---
# <a name="understanding-the-javascript-api-for-office"></a>Noções básicas da API JavaScript para Office

Este artigo fornece informações sobre a API JavaScript para Office e como usá-la. Para referenciar as informações, consulte [API JavaScript para Office](/office/dev/add-ins/reference/javascript-api-for-office). Para obter informações sobre como atualizar os arquivos de projeto do Visual Studio para a versão mais recente da API JavaScript para Office, consulte [Atualizar a versão da API JavaScript para Office e arquivos de esquema do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Fazer referência à biblioteca da API JavaScript para Office no suplemento

A biblioteca da [API JavaScript para Office](/office/dev/add-ins/reference/javascript-api-for-office) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo de host associado, como Excel-15.js e Outlook-15.js. O método mais simples de fazer referência à API é usando nossa CDN e adicionando o seguinte `<script>` à marca `<head>` da sua página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

Isso baixará e colocará os arquivos da API JavaScript para Office em cache quando o suplemento for carregado pela primeira vez a fim de garantir que o suplemento esteja usando a implementação mais recente do Office.js e de seus arquivos associados na versão especificada.

Para saber mais sobre a CDN do Office.js, inclusive como é feito o controle de versão e como lidar com a compatibilidade com versões anteriores, confira [Fazendo referência à biblioteca da API JavaScript para Office na CDN (rede de distribuição de conteúdo)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Inicialização do suplemento

**Aplica-se a:** todos os tipos de suplementos

Os Suplementos do Office têm sempre uma lógica de inicialização para fazer coisas como:

- Verificar se a versão do Office do usuário será compatível com todas as APIs do Office chamadas pelo código.

- Garantir a existência de determinados artefatos, como uma planilha de nome específico.

- Solicitar ao usuário selecionar algumas células no Excel e inserir um gráfico inicializado com esses valores selecionados.

- Estabeleça associações.

- Usar a API de caixa de diálogo do Office para solicitar ao usuário definir valores de configurações padrão para o suplemento.

O código de inicialização só deverá chamar as APIs do Office.js quando a biblioteca estiver carregada. Há duas maneiras pelas quais o código pode garantir que a biblioteca seja carregada. Elas estão descritas nas seções a seguir: 

- [Inicializar com Office.onReady()](#initialize-with-officeonready)
- [Inicializar com Office.initialize](#initialize-with-officeinitialize)

> [!TIP]
> Recomendamos que use `Office.onReady()`em vez de`Office.initialize`. Embora `Office.initialize` ainda tenha suporte, usar`Office.onReady()` oferece mais flexibilidade. É possível atribuir apenas um manipulador a `Office.initialize`, e ela é chamada apenas uma vez pela infraestrutura do Office. Mas você pode chamar `Office.onReady()` em diferentes locais no código, e usar diferentes retornos de chamadas.
> 
> Para saber mais sobre as diferenças entre essas técnicas, veja [Principais diferenças entre Office.initialize e Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).

Para saber mais sobre a sequência de eventos na inicialização do suplemento, confira [Carregar o ambiente de tempo de execução e o DOM](loading-the-dom-and-runtime-environment.md).

### <a name="initialize-with-officeonready"></a>Inicializar com o Office.onReady()

`Office.onReady()` é um método assíncrono que retorna um objeto Promise enquanto verifica se a biblioteca do Office.js está carregada. Somente quando a biblioteca é carregada, ela resolve o Promise como um objeto que especifica o aplicativo host do Office com um valor de enumeração `Office.HostType` (`Excel`, `Word` etc.), e a plataforma com um valor de enumeração `Office.PlatformType` (`PC`, `Mac`, `OfficeOnline` etc.). O Promise será resolvido imediatamente quando a biblioteca estiver carregada ao`Office.onReady()` ser chamada.

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

No entanto, há exceções a essa prática. Por exemplo, digamos que você queira abrir o suplemento em um navegador (em vez de fazer sideload em um host do Office) para depurar a interface do usuário com ferramentas de navegador. Já que o Office.js não será carregado no navegador, `onReady` não será executado e o `$(document).ready` não será executado quando chamado dentro de `onReady` no Office. Outra exceção: você deseja que um indicador de progresso seja exibido no painel de tarefas enquanto o suplemento está sendo carregado. Nesse cenário, o código deve chamar `ready` da jQuery e usa a respectiva chamada de retorno para renderizar o indicador de progresso. Em seguida, a chamada de retorno do Office `onReady` pode substituir o indicador de progresso com a interface do usuário final. 

### <a name="initialize-with-officeinitialize"></a>Inicializar com Office.initialize

Um evento de inicialização é disparado quando a biblioteca do Office.js está carregada e pronta para a interação com o usuário. É possível atribuir um manipulador ao `Office.initialize` que implementa a lógica de inicialização. Veja a seguir um exemplo que verifica se a versão do Excel do usuário é compatível com todas as APIs que o suplemento pode chamar.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Se estiver usando estruturas JavaScript adicionais que incluam testes e manipuladores próprios de inicialização, *geralmente* eles devem ser colocados dentro do evento `Office.initialize`. No entanto, as exceções descritas anteriormente na seção **Inicializar com Office.onReady()** também se aplicam neste caso. Por exemplo, a função `$(document).ready()` do [JQuery](https://jquery.com) pode ser referenciada da seguinte maneira:

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

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Principais diferenças entre Office.initialize e Office.onReady

- É possível atribuir apenas um manipulador a `Office.initialize`, e ela é chamada apenas uma vez pela infraestrutura do Office, mas você pode chamar `Office.onReady()` em diferentes locais no código, e usar diferentes retornos de chamadas. Por exemplo, o código pode chamar `Office.onReady()`, logo que o script personalizado é carregado com um retorno de chamada que executa uma lógica de inicialização. Além disso, o código pode ter um botão no painel de tarefas, cujo script chama `Office.onReady()` com um retorno de chamada diferente. Quando isso ocorre, o segundo retorno de chamada é executado quando o botão é clicado.

- O evento `Office.initialize` é disparado no final do processo interno, e que o Office.js é inicializado automaticamente. Ele também é disparado *imediatamente* após o término do processo interno. Se o código no qual você atribui um manipulador ao evento for executado muito tempo após o evento ser disparado, então o manipulador não será executado. Por exemplo, se estiver usando o gerenciador de tarefas WebPack, ele poderá configurar a home page do suplemento para carregar arquivos de polyfill, após carregar o Office.js, mas antes de carregar o JavaScript personalizado. Quando o script carrega e atribui o manipulador, o evento de inicialização já ocorreu. Mas nunca é "tarde demais" para chamar `Office.onReady()`. Caso o evento de inicialização já tenha ocorrido, o retorno de chamada é executado imediatamente.

> [!NOTE]
> Mesmo que não tenha uma lógica de inicialização, você deve atribuir ou chamar `Office.onReady()` uma função vazia para `Office.initialize` quando o JavaScript do suplemento for carregado. Algumas combinações de host e da plataforma do Office não carregam o painel de tarefas até uma das delas aconteça. Os exemplos a seguir mostram essas duas abordagens.
>
>```js  
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Modelo de objeto de API JavaScript para Office

Depois de inicializado, o suplemento pode interagir com o host (por exemplo, o Excel ou o Outlook). A página [Modelo do objeto do JavaScript API comum](office-javascript-api-object-model.md) possui mais detalhes sobre padrões de uso específicos. Há também documentação de referência detalhada para [APIs Comuns](/office/dev/add-ins/reference/javascript-api-for-office) e hosts específicos.
