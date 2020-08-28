---
title: Carregar o ambiente de tempo de execução e DOM
description: Carregar o ambiente de tempo de execução de suplementos do Office e DOM
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 02f950ca23d52b333f704c7d8aed431cb426a6f0
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293272"
---
# <a name="loading-the-dom-and-runtime-environment"></a>Carregar o ambiente de tempo de execução e DOM

Um suplemento deve garantir que o DOM e o ambiente de tempo de execução de Suplementos do Office sejam carregados antes de executar sua própria lógica personalizada.

## <a name="startup-of-a-content-or-task-pane-add-in"></a>Inicialização de um suplemento de conteúdo ou de painel de tarefas

A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento de conteúdo ou de painel de tarefas no Excel, no PowerPoint, no Project ou no Word.

![Fluxo de eventos ao iniciar um suplemento de conteúdo ou de painel de tarefas](../images/office15-app-sdk-loading-dom-agave-runtime.png)

Os eventos a seguir ocorrem quando um suplemento de conteúdo ou de painel de tarefas é iniciado:

1. O usuário abre um documento que já contém um suplemento ou insere um suplemento no documento.

2. O aplicativo cliente do Office lê o manifesto XML do suplemento no AppSource, em um catálogo de aplicativos no SharePoint ou no catálogo de pasta compartilhada do qual ele se originou.

3. O aplicativo cliente do Office abre a página HTML do suplemento em um controle de navegador.

    As próximas duas etapas, as etapas 4 e 5, ocorrem de forma assíncrona e em paralelo. Por esse motivo, o código do suplemento deve garantir que o DOM e o ambiente do tempo de execução do suplemento tenham terminado de carregar antes de prosseguir.

4. O controle do navegador carrega o corpo do HTML e DOM e chama o manipulador de eventos para o `window.onload` evento.

5. O aplicativo cliente do Office carrega o ambiente de tempo de execução, que baixa e armazena em cache os arquivos da biblioteca da API JavaScript do Office do servidor de rede de distribuição de conteúdo (CDN) e, em seguida, chama o manipulador de eventos do suplemento para o evento [Initialize](/javascript/api/office#office-initialize-reason-) do objeto [Office](/javascript/api/office) , se um manipulador tiver sido atribuído a ele. Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador. Para obter mais informações sobre a distinção entre o `Office.initialize` e o `Office.onReady` , consulte [Initialize Your Add-in](initialize-add-in.md).

6. Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.


## <a name="startup-of-an-outlook-add-in"></a>Inicialização de um suplemento do Outlook

A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento do Outlook em execução no desktop, tablet ou smartphone.

![Fluxo de eventos ao inicializar um suplemento do Outlook](../images/outlook15-loading-dom-agave-runtime.png)

Os eventos a seguir ocorrem quando um suplemento Outlook é iniciado:

1. Quando é iniciado, o Outlook lê os manifestos XML para suplementos do Outlook que foram instalados na conta de email do usuário.

2. O usuário seleciona um item no Outlook.

3. Se o item selecionado satisfizer as condições de ativação de um suplemento do Outlook, o Outlook ativará o suplemento e tornará seu botão visíveis na interface de usuário.

4. Se o usuário clicar no botão para iniciar o suplemento do Outlook, o Outlook abrirá a página HTML em um controle de navegador. As próximas duas etapas, as etapas 5 e 6, ocorrerem em paralelo.

5. O controle do navegador carrega o corpo do HTML e DOM e chama o manipulador de eventos para o `onload` evento.

6. O Outlook carrega o ambiente de tempo de execução, que baixa e armazena em cache a API do JavaScript para arquivos da biblioteca a partir do servidor da rede de distribuição de conteúdo (CDN) e chama manipulador de eventos para o evento [initialize](/javascript/api/office#office-initialize-reason-) do objeto do suplemento do [Office](/javascript/api/office). Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador. Para obter mais informações sobre a distinção entre o `Office.initialize` e o `Office.onReady` , consulte [Initialize Your Add-in](initialize-add-in.md).

7. Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.


## <a name="checking-the-load-status"></a>Verificar o status de carregamento

Uma maneira de verificar se o ambiente de tempo de execução e o DOM concluíram o carregamento é usar a função [.ready()](https://api.jquery.com/ready/) do jQuery: `$(document).ready()`. Por exemplo, o manipulador de eventos a seguir `onReady` garante que o dom seja carregado primeiro antes que o código específico para inicializar o suplemento seja executado. Subsequentemente, o `onReady` manipulador continua a usar a propriedade [Mailbox. Item](/javascript/api/outlook/office.mailbox#item) para obter o item atualmente selecionado no Outlook e chama a função principal do suplemento, `initDialer` .

```js
Office.onReady()
    .then(
        // Checks for the DOM to load.
        $(document).ready(function () {
            // After the DOM is loaded, add-in-specific code can run.
            var mailbox = Office.context.mailbox;
            _Item = mailbox.item;
            initDialer();
        });
);
```

Como alternativa, você pode usar o mesmo código em um `initialize` manipulador de eventos, conforme mostrado no exemplo a seguir.

```js
Office.initialize = function () {
    // Checks for the DOM to load.
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        var mailbox = Office.context.mailbox;
        _Item = mailbox.item;
        initDialer();
    });
}
```

Essa mesma técnica pode ser usada nos `onReady` `initialize` manipuladores ou de qualquer suplemento do Office.

O suplemento do Outlook de amostra de discagem telefônica mostra uma abordagem ligeiramente diferente usando somente o JavaScript para verificar essas mesmas condições.

> [!IMPORTANT]
> Mesmo que seu suplemento não tenha tarefas de inicialização para executar, você deve incluir pelo menos uma chamada `Office.onReady` ou atribuir uma função de `Office.initialize` manipulador de eventos mínima, conforme mostrado nos exemplos a seguir.
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> Se você não chamar `Office.onReady` ou atribuir um `Office.initialize` manipulador de eventos, seu suplemento poderá gerar um erro quando for iniciado. Além disso, se um usuário tentar usar o suplemento com um cliente virtual do Office Online, como Excel, PowerPoint ou Outlook, ele não funcionará.
>
> Se o suplemento incluir mais de uma página, sempre que carregar uma nova página, a página deverá chamar `Office.onReady` ou atribuir um manipulador de `Office.initialize` eventos.

## <a name="see-also"></a>Confira também

- [Entendendo a API JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [Inicialize seu suplemento do Office](initialize-add-in.md)
