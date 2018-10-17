---
title: Noções básicas da API JavaScript para Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e9d9efdda5e237ab076d22d50b1f7ded5e075845
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505942"
---
# <a name="understanding-the-javascript-api-for-office"></a>Noções básicas da API JavaScript para Office

Este artigo fornece informações sobre a API JavaScript para Office e como usá-la. Para referenciar as informações, consulte [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js). Para obter informações sobre como atualizar os arquivos de projeto do Visual Studio para a versão mais recente da API JavaScript para Office, consulte [Atualizar a versão da API JavaScript para Office e arquivos de esquema do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) seu suplemento no AppSource e disponibilizá-lo na experiência do Office, verifique se está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deverá funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Fazer referência à biblioteca da API JavaScript para Office no suplemento

A biblioteca da [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo host associado, como Excel-15.js e Outlook-15.js. O método mais simples de fazer referência à API é usando nossa CDN e adicionando o seguinte `<script>` à marca `<head>` da sua página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Isso baixará e colocará os arquivos da API JavaScript para Office em cache quando o suplemento for carregado pela primeira vez a fim de garantir que o suplemento esteja usando a implementação mais recente do Office.js e de seus arquivos associados para a versão especificada.

Para saber mais sobre a CDN do Office.js, incluindo como é feito o controle de versão e como lidar com a compatibilidade com versões anteriores, veja [Fazer referência à biblioteca da API JavaScript para Office a partir da sua rede de distribuição de conteúdo (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Inicializar o seu suplemento

**Aplica-se a:** todos os tipos de suplementos

Suplementos do Office geralmente têm uma lógica de inicialização para realizar tarefas como:

- Verificar se a versão do Office do usuário tem suporte para todas as APIs do Office chamadas pelo seu código.

- Garantir a existência de determinados artefatos, como uma planilha com um nome específico.

- Solicitar que o usuário selecione algumas células no Excel e, em seguida, inserir um gráfico inicializado com esses valores selecionados.

- Estabelecer associações.

- Use a caixa de diálogo da API do Office para solicitar ao usuário valores padrão de configurações de suplementos.

Mas seu código de inicialização não deve chamar nenhuma APIs Office.js até a biblioteca ser totalmente carregada. Existem duas maneiras de seu código garantir que a biblioteca esteja carregada. Elas são descritas nas seções a seguir. Recomendamos que você use a técnica mais recente e flexível, chamada `Office.onReady()`. A técnica mais antiga, atribuindo um manipulador de `Office.initialize`, ainda é suportada. Consulte também [Principais diferenças entre Office.initialize e Office.onReady()](#major-differences-between-office-initialize-and-office-onready).

Para obter mais detalhes sobre a sequência de eventos na inicialização de suplementos, consulte [Carregamento do DOM e o ambiente de tempo de execução](loading-the-dom-and-runtime-environment.md).

### <a name="initialize-with-officeonready"></a>A inicialização com Office.onReady()

`Office.onReady()` é um método assíncrono que retorna um objeto Promise enquanto verifica se a biblioteca do Office. js está totalmente carregada. Quando e somente quando a biblioteca for carregada, ele resolve o Promise como um objeto que especifica o aplicativo host do Office com um `Office.HostType` valor enum (`Excel`, `Word`, etc) e a plataforma com um `Office.PlatformType` valor enum (`PC`, `Mac`, `OfficeOnline`, etc.). Se a biblioteca já estiver carregada quando `Office.onReady()` for chamado, o Promise resolve imediatamente.

Uma maneira de chamar `Office.onReady()` é passar a ele um método de retorno de chamada. Aqui está um exemplo:

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

Como alternativa, você pode encadear um método `then()` para a chamada de `Office.onReady()`, em vez de passar um retorno de chamada. Por exemplo, o código a seguir verifica se a versão do usuário do Excel oferece suporte a todas as APIs que o suplemento pode chamar.

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

Se você estiver usando estruturas JavaScript adicionais que incluem seus próprios manipuladores de inicialização ou testes, esses devem ser *geralmente* colocados dentro da resposta a  `Office.onReady()` . Por exemplo, a função [JQuery](https://jquery.com) `$(document).ready()`seria referenciada da seguinte maneira:

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

No entanto, há exceções a essa prática. Por exemplo, suponha que você deseja abrir o suplemento em um navegador (em vez de fazer o sideload em um host do Office) para depurar sua interface de usuário com as ferramentas do navegador. Como o Office. js não será carregado no navegador, `onReady` não será executado e o `$(document).ready` não será executado se for chamado dentro do Office `onReady`. Outra exceção: você deseja que um indicador de progresso apareça no painel de tarefas, enquanto o suplemento está carregando. Neste cenário, seu código deve chamar o `ready` do jQuery e usar sua chamada de retorno para renderizar o indicador de progresso. Em seguida, a chamada de retorno do `onReady`do Office pode substituir o indicador de progresso com a interface do usuário final. 

### <a name="initialize-with-officeinitialize"></a>Inicializar com Office.initialize

Um evento initialize é acionado quando a biblioteca Office.js está totalmente carregada e pronta para a interação do usuário. Você pode atribuir um manipulador a `Office.initialize` que implementa a lógica de inicialização. O exemplo a seguir verifica se a versão do usuário do Excel oferece suporte a todas as APIs que o suplemento pode chamar.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Se você estiver usando as estruturas adicionais do JavaScript que incluem seu próprio manipulador de inicialização ou testes, elas *geralmente* devem ser colocadas dentro do evento `Office.initialize` . (Mas as exceções descritas na seção **Inicializar com Office.onReady()** anteriormente, se aplicam neste caso também.) Por exemplo, a função do [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

Para suplementos de conteúdo e de painel de tarefas, `Office.initialize` fornece um parâmetro _reason_ adicional. Esse parâmetro pode ser usado para determinar como um suplemento foi adicionado ao documento atual. Você pode usar isso para fornecer lógica diferente quando um suplemento é inserido pela primeira vez em comparação a quando já existia dentro do documento.

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

Para obter mais informações, confira [Evento Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) e [Enumeração InitializationReason](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js).

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Principais diferenças entre Office.initialize e Office.onReady()

- Você pode atribuir apenas um manipulador para `Office.initialize` e é chamado, apenas uma vez, pela infra-estrutura do Office; mas você pode chamar `Office.onReady()` em diferentes locais do seu código e usar diferentes retornos de chamada. Por exemplo, seu código poderia chamar `Office.onReady()` assim que seu script personalizado for carregado com um retorno de chamada que executa a lógica de inicialização; e seu código também pode ter um botão no painel de tarefas, cujo script chama `Office.onReady()` com um retorno de chamada diferente. Nesse caso, o segundo retorno de chamada é executado quando o botão é clicado.

- O evento `Office.initialize` é acionado ao final do processo interno no qual o Office. js se inicializa. E ele aciona *imediatamente* após o término do processo interno. Se o código em que você atribui um manipulador de evento for executado muito tempo depois do evento ser acionado, não execute o seu manipulador. Por exemplo, se você estiver usando o Gerenciador de tarefas WebPack, ele pode configurar a página inicial do suplemento para carregar os arquivos polifyll depois de carregar o Office. js, mas antes de carregar seu JavaScript personalizado. No momento em seu script carregar e atribuir o manipulador, o evento initialize já aconteceu. Mas nunca é "muito tarde" para chamar `Office.onReady()`. Se o evento initialize já tiver acontecido, o retorno de chamada executará imediatamente.

> [!NOTE]
> Mesmo que você não tenha uma lógica de inicialização, é uma boa prática chamar `Office.onReady()` ou atribuir uma função vazia a `Office.initialize` quando o JavaScript do seu suplemento for carregado, pois algumas combinações de plataforma e host do Office não carregarão o painel de tarefas até que uma dessas situações aconteça. As linhas a seguir mostram as duas maneiras de fazer isso:
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Modelo de objeto da API JavaScript para Office

Uma vez inicializado, o suplemento pode interagir com o host (ex. Excel, Outlook). A página do [Modelo de Objeto da API Javascript do Office](office-javascript-api-object-model.md) tem mais detalhes sobre modelos de uso específicos. Também há documentação de referência detalhada para [APIs compartilhadas](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) e hosts específicos.

## <a name="api-support-matrix"></a>Matriz de suporte da API

Esta tabela resume a API e os recursos compatíveis com os tipos de suplemento (conteúdo, painel de tarefas e Outlook) e os aplicativos do Office que podem hospedá-los quando o usuário especifica os aplicativos host do Office compatíveis com o suplemento usando o [esquema de manifesto de suplementos 1.1 e recursos compatíveis com a API JavaScript para Office v1.1](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Nome do host**|Banco de dados|Pasta de trabalho|Caixa de correio|Apresentação|Documento|Projeto|
||**Aplicativos host** **compatíveis**|Aplicativos Web do Access|Excel,<br/>Excel Online|Outlook,<br/>Aplicativo Web do Outlook,<br/>OWA para dispositivos|PowerPoint,<br/>PowerPoint Online|Word|Project|
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
||Associações e eventos de associação|S<br/>(Somente associações de tabela totais e parciais)|S|||S||
||Ler/gravar partes XML personalizadas|||||S||
||Persistir dados de estado de suplemento (configurações)|S<br/>(Por suplemento do host)|S<br/>(Por documento)|S<br/>(Por caixa de correio)|S<br/>(Por documento)|S<br/>(Por documento)||
||Eventos alterados pelas configurações|S|S||S|S||
||Obter o modo de exibição ativo<br/>e visualizar eventos alterados||||S|||
||Navegar para locais<br/>no documento||S||S|S||
||Ativar contextualmente<br/>usando regras e RegEx|||S||||
||Ler propriedades do item|||S||||
||Ler perfil do usuário|||S||||
||Obter anexos|||S||||
||Obter o token de identidade do usuário|||S||||
||Chamar os serviços Web do Exchange|||S||||
