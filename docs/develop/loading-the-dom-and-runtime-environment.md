---
title: Carregar o ambiente de tempo de execução e DOM
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: b1f63d9fe012ed8c8a5cf4a0f7de862ddabcd4d3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449833"
---
# <a name="loading-the-dom-and-runtime-environment"></a>Carregar o ambiente de tempo de execução e DOM

Um suplemento deve garantir que o DOM e o ambiente de tempo de execução de Suplementos do Office sejam carregados antes de executar sua própria lógica personalizada. 

## <a name="startup-of-a-content-or-task-pane-add-in"></a>Inicialização de um suplemento de conteúdo ou de painel de tarefas

A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento de conteúdo ou de painel de tarefas no Excel, no PowerPoint, no Project, no Word ou no Access.

![Fluxo de eventos ao iniciar um suplemento de conteúdo ou de painel de tarefas](../images/office15-app-sdk-loading-dom-agave-runtime.png)

Os eventos a seguir ocorrem quando um suplemento de conteúdo ou de painel de tarefas é iniciado:

1. O usuário abre um documento que já contém um suplemento ou insere um suplemento no documento.

2. O aplicativo host do Office lê o manifesto XML do suplemento a partir do AppSource, catálogo de suplementos no SharePoint ou catálogo de pastas compartilhada do qual ele se originou.

3. O aplicativo host do Office abre a página de HTML do suplemento em um controle de navegador.

    As próximas duas etapas, as etapas 4 e 5, ocorrem de forma assíncrona e em paralelo. Por esse motivo, o código do suplemento deve garantir que o DOM e o ambiente do tempo de execução do suplemento tenham terminado de carregar antes de prosseguir.

4. O controle do navegador carrega o corpo do HTML e DOM e chama o manipulador de eventos para o evento **window.onload**.

5. O aplicativo host do Office carrega o ambiente de tempo de execução, que baixa e armazena em cache a API do JavaScript para arquivos da biblioteca a partir do servidor da rede de distribuição de conteúdo (CDN) e chama manipulador de eventos do suplemento para o evento [initialize](/javascript/api/office#office-initialize) do objeto [Office](/javascript/api/office), se um identificador for atribuído a ele. Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador. Para saber mais sobre a distinção entre `Office.initialize` e `Office.onReady`, confira [Inicializar o suplemento](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).

6. Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.


## <a name="startup-of-an-outlook-add-in"></a>Inicialização de um suplemento do Outlook

A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento do Outlook em execução no desktop, tablet ou smartphone.

![Fluxo de eventos ao inicializar um suplemento do Outlook](../images/outlook15-loading-dom-agave-runtime.png)

Os eventos a seguir ocorrem quando um suplemento Outlook é iniciado:

1. Quando é iniciado, o Outlook lê os manifestos XML para suplementos do Outlook que foram instalados na conta de email do usuário.

2. O usuário seleciona um item no Outlook.

3. Se o item selecionado satisfizer as condições de ativação de um suplemento do Outlook, o Outlook ativará o suplemento e tornará seu botão visíveis na interface de usuário.

4. Se o usuário clicar no botão para iniciar o suplemento do Outlook, o Outlook abrirá a página HTML em um controle de navegador. As próximas duas etapas, as etapas 5 e 6, ocorrerem em paralelo.

5. O controle do navegador carrega o corpo do HTML e DOM e chama o manipulador de eventos para o evento **onload**.

6. O Outlook carrega o ambiente de tempo de execução, que baixa e armazena em cache a API do JavaScript para arquivos da biblioteca a partir do servidor da rede de distribuição de conteúdo (CDN) e chama manipulador de eventos para o evento [initialize](/javascript/api/office#office-initialize) do objeto do suplemento do [Office](/javascript/api/office). Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador. Para saber mais sobre a distinção entre `Office.initialize` e `Office.onReady`, confira [Inicializar o suplemento](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).

7. Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.


## <a name="checking-the-load-status"></a>Verificar o status de carregamento

Uma maneira de verificar se o ambiente de tempo de execução e o DOM concluíram o carregamento é usar a função [.ready()](https://api.jquery.com/ready/) do jQuery: `$(document).ready()`. Por exemplo, o manipulador de eventos **onReady** garante que o DOM seja carregado antes do código específico para inicializar as execuções de suplementos. Subsequentemente, o manipulador de eventos **onReady** prossegue e usa a propriedade [mailbox.item](/javascript/api/outlook/office.mailbox) para obter o item selecionado atual no Outlook, e chama a função principal do suplemento, `initDialer`.

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

Como alternativa, você pode usar o mesmo código em um manipulador de eventos **initialize** manipulador de eventos, conforme mostrado no exemplo a seguir.

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

Essa mesma técnica pode ser usada nos manipuladores **onReady** ou **initialize** de qualquer Suplemento do Office.

O suplemento do Outlook de amostra de discagem telefônica mostra uma abordagem ligeiramente diferente usando somente o JavaScript para verificar essas mesmas condições. 

> [!IMPORTANT]
> Mesmo que o suplemento não tenha tarefas de inicialização para executar, você deve incluir pelo menos uma chamada do manipulador de eventos **Office.onReady** ou atribuir a função mínima do manipulador de eventos **Office.initialize**, como mostra o exemplo a seguir.
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> Se você não chamar **Office.onReady** ou atribuir um manipulador de eventos**Office.initialize**, o suplemento pode gerar um erro quando inicia. Além disso, se um usuário tentar usar o suplemento com um cliente Web do Office Online, como o Excel Online, o PowerPoint Online ou o Outlook Web App, ele não funcionará.
>
> Se o suplemento incluir mais de uma página, essa página deverá chamar o manipulador de eventos **Office.onReady** atribuir um manipulador de eventos **Office.initialize** sempre que uma nova página for carregada.

## <a name="see-also"></a>Confira também

- [Noções básicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md)
