---
title: Noções básicas da API JavaScript para Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 12e7d9030ec37746f84e3fc725cddda2a5675761
ms.sourcegitcommit: 5bef9828f047da03ecf2f43c6eb5b8514eff28ce
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/31/2018
ms.locfileid: "23782791"
---
# <a name="understanding-the-javascript-api-for-office"></a>Noções básicas da API JavaScript para Office

Este artigo fornece informações sobre a API JavaScript para Office e como usá-la. Para referenciar as informações, consulte [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office). Para obter informações sobre como atualizar os arquivos de projeto do Visual Studio para a versão mais recente da API JavaScript para Office, consulte [Atualizar a versão da API JavaScript para Office e arquivos de esquema do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Fazer referência à biblioteca da API JavaScript para Office no suplemento

A biblioteca da [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo de host associado, como Excel-15.js e Outlook-15.js. O método mais simples de fazer referência à API é usando nossa CDN e adicionando o seguinte `<script>` à marca `<head>` da sua página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Isso baixará e colocará os arquivos da API JavaScript para Office em cache quando o suplemento for carregado pela primeira vez a fim de garantir que o suplemento esteja usando a implementação mais recente do Office.js e de seus arquivos associados na versão especificada.

Para saber mais sobre a CDN do Office.js, incluindo como é feito o controle de versão e como lidar com a compatibilidade com versões anteriores, veja [Fazer referência à biblioteca da API JavaScript para Office a partir da sua rede de distribuição de conteúdo (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Inicializar o seu suplemento

**Aplica-se a:** todos os tipos de suplementos

Suplementos do Office geralmente têm uma lógica de inicialização para realizar tarefas como:

- Verificar se a versão do Office do usuário tem suporte para todas as APIs do Office chamadas pelo seu código.

- Garantir a existência de determinados artefatos, como uma planilha com um nome específico.

- Solicitar que o usuário selecione algumas células no Excel e, em seguida, inserir um gráfico inicializado com esses valores selecionados.

- Estabelecer associações.

- Use a caixa de diálogo da API do Office para solicitar valores padrão de configurações de suplementos para o usuário.

Mas o seu código de inicialização não deve chamar nenhuma API do Office.js até que a biblioteca esteja totalmente carregada. Existem duas maneiras do seu código garantir que a biblioteca está carregada. Eles são descritos nas seções a seguir. Recomendamos que você use a técnica mais recente e flexível, chamando `Office.onReady()`. Ainda há suporte para a técnica antiga, que atribui um manipulador a `Office.initialize`. Consulte também [Principais diferenças entre Office.initialize e Office.onReady()](#major-differences-between-office-initialize-and-office-onready).

Para obter mais detalhes sobre a sequência de eventos na inicialização de suplementos, consulte [Carregar o DOM e o ambiente de execução](loading-the-dom-and-runtime-environment.md).

### <a name="initialize-with-officeonready"></a>A inicialização com Office.onReady()

`Office.onReady()` é um método assíncrono que retorna um objeto Promise enquanto verifica se a biblioteca do Office.js foi totalmente carregada. Somente quando a biblioteca for carregada, ele resolve o Promise como um objeto que especifica o aplicativo de host do Office com um valor de enumeração `Office.HostType` (`Excel`, `Word`, etc) e a plataforma com um valor de enumeração `Office.PlatformType` (`PC`, `Mac`, `OfficeOnline`, etc.). Se a biblioteca já estiver carregada quando `Office.onReady()` é chamado, o objeto Promise é resolvido imediatamente.

Uma maneira de chamar `Office.onReady()` é passando um método de retorno para ele. Por exemplo:

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

Outra opção é encadear um método `then()` para a chamada de `Office.onReady()`, em vez de passar um retorno. Por exemplo, o código a seguir verifica se a versão do Excel do usuário oferece suporte a todas as APIs que o suplemento pode chamar.

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

Este é o mesmo exemplo, usando as palavras-chave `async` e `await` em TypeScript:

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

Se você estiver usando estruturas adicionais do JavaScript que incluem seu próprio manipulador de inicialização ou testes, elas *normalmente* devem ser colocadas dentro da resposta para `Office.onReady()`. Por exemplo, a função [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

No entanto, há exceções a essa prática. Por exemplo, suponha que você deseja abrir o suplemento em um navegador (em vez de fazer o sideload em um host do Office) para depurar sua interface de usuário com as ferramentas do navegador. Como o Office.js não é carregado no navegador, `onReady` não é executado e `$(document).ready` não é executado se chamado dentro do `onReady` do Office. Outra exceção: você quer que um indicador de andamento apareça no painel de tarefas enquanto o suplemento está sendo carregado. Nesse cenário, seu código deve chamar o `ready` do jQuery e usar seu retorno para renderizar o indicador de andamento. Em seguida, o retorno do `onReady` do Office pode substituir o indicador de andamento com a interface final do usuário. 

### <a name="initialize-with-officeinitialize"></a>Inicializar com Office.initialize

Um evento initialize é acionado quando a biblioteca do Office.js está totalmente carregada e pronta para interação com o usuário. Você pode atribuir um manipulador a `Office.initialize`, que implementa a lógica de inicialização. O exemplo a seguir verifica se a versão do Excel do usuário oferece suporte a todas as APIs que o suplemento pode chamar.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Se você estiver usando estruturas adicionais do JavaScript que incluem seu próprio manipulador de inicialização ou testes, elas *normalmente* devem ser colocadas dentro do evento `Office.initialize`. (Mas as exceções descritas na seção anterior, **Inicializar com Office.onReady()**, também se aplicam a este caso.) Por exemplo, a função do [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

Para suplementos de conteúdo e painel de tarefas, `Office.initialize` fornece um parâmetro adicional _reason_. Esse parâmetro especifica como um suplemento foi adicionado ao documento atual. Você pode usar isso para fornecer uma lógica diferente para quando um suplemento for inserido pela primeira vez, em contraste com a usada quando ele já existia dentro do documento.

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

Para obter mais informações, confira [Evento Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) e [Enumeração InitializationReason](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration).

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Principais diferenças entre Office.initialize e Office.onReady()

- Você pode atribuir apenas um manipulador para `Office.initialize` e ele é chamado apenas uma vez pela infraestrutura do Office; mas você pode chamar `Office.onReady()` em diferentes locais do seu código e usar diferentes retornos de chamada. Por exemplo, seu código pode chamar `Office.onReady()` assim que seu script personalizado for carregado com um retorno de chamada que executa a lógica de inicialização; seu código também pode ter um botão no painel de tarefas cujo script chama `Office.onReady()` com um retorno de chamada diferente. Nesse caso, o segundo retorno de chamada é executado quando o botão é clicado.

- O evento `Office.initialize` é acionado ao final do processo interno no qual Office.js inicializa a si mesmo. E ele é acionado *imediatamente* após o término do processo interno. Se o código em que você atribui um manipulador para o evento é executado muito tempo depois do evento ser acionado, o seu manipulador não é executado. Por exemplo, se você estiver usando o gerenciador de tarefas WebPack, ele pode configurar a página inicial do suplemento para carregar arquivos polyfill após carregar o Office.js e antes de carregar seu JavaScript personalizado. Quando o seu script carregar e atribuir o manipulador, o evento initialize já terá acontecido. Mas nunca é "tarde demais" para chamar `Office.onReady()`. Se o evento initialize já tiver acontecido, o retorno de chamada é executado imediatamente.

> [!NOTE]
> Mesmo se você não tiver uma lógica de inicialização, é uma boa prática chamar `Office.onReady()` ou atribuir uma função vazia a `Office.initialize` quando o JavaScript do seu suplemento for carregado, pois algumas combinações de plataforma e host do Office não carregam o painel de tarefas até que um desses aconteça. As linhas a seguir mostram as duas maneiras de que isso pode ser feito:
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Modelo de objeto da API JavaScript do Office

Uma vez inicializado, o suplemento pode interagir com o host (por exemplo, Excel, Outlook). A página [Modelo de objeto da API JavaScript do Office](office-javascript-api-object-model.md) tem mais detalhes sobre padrões de uso específicos. Há também documentação de referência detalhada para as [APIs compartilhadas](https://dev.office.com/reference/add-ins/javascript-api-for-office) e hosts específicos.

## <a name="api-support-matrix"></a>Matriz de suporte da API

Esta tabela resume a API e os recursos compatíveis com os tipos de suplemento (conteúdo, painel de tarefas e Outlook) e os aplicativos do Office que podem hospedá-los quando o usuário especifica os aplicativos hospedados pelo Office compatíveis com o suplemento usando o [esquema 1.1 do manifesto de suplementos e recursos compatíveis com a v1.1 da API JavaScript para Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Nome do host**|Banco de dados|Pasta de trabalho|Caixa de correio|Apresentação|Documento|Projeto|
||**Aplicativos host** **compatíveis**|Aplicativos Web do Access|Excel,<br/>Excel Online|Outlook,<br/>Aplicativo Web do Outlook,<br/>OWA para dispositivos|PowerPoint,<br/>PowerPoint Online|Word|Projeto|
|**Tipos de suplemento com suporte**|Conteúdo|S|S||S|||
||Painel de tarefas||S||S|S|S|
||Outlook|||S||||
|**Recursos da API compatíveis**|Ler/gravar texto||S||S|S|S<br/>(Somente leitura)|
||Ler/gravar matriz||S|||S||
||Ler/gravar tabela||S|||S||
||Ler/gravar HTML|||||S||
||Leitura/gravação<br/>Open XML do Office|||||S||
||Ler propriedades de tarefa, recurso, modo de exibição e campo||||||S|
||Eventos alterados pela seleção||S|||S||
||Obter documento inteiro||||S|S||
||Associações e eventos de associação|S<br/>(Somente vinculações de tabela totais e parciais)|S|||S||
||Ler/gravar partes XML personalizadas|||||S||
||Persistir dados de estado de suplemento (configurações)|S<br/>(Por suplemento do host)|S<br/>(Por documento)|S<br/>(Por caixa de correio)|S<br/>(Por documento)|S<br/>(Por documento)||
||Eventos alterados pelas configurações|S|S||S|S||
||Obter o modo de exibição ativo<br/>e visualizar eventos alterados||||S|||
||Navegar para locais<br/>no documento||S||S|S||
||Ativar contextualmente<br/>usando regras e RegEx|||S||||
||Ler propriedades do item|||S||||
||Ler perfil de usuário|||S||||
||Obter anexos|||S||||
||Obter o token de identidade do usuário|||S||||
||Chamar os serviços Web do Exchange|||S||||
