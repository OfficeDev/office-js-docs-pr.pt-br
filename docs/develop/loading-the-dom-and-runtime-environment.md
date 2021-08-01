---
title: Carregar o ambiente de tempo de execução e DOM
description: Carregue o dom e Office ambiente de tempo de execução de complementos.
ms.date: 04/20/2021
localization_priority: Normal
ms.openlocfilehash: 76322de3d830747a825b0934993f8801daf8a9ae
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671524"
---
# <a name="loading-the-dom-and-runtime-environment"></a>Carregar o ambiente de tempo de execução e DOM

Um suplemento deve garantir que o DOM e o ambiente de tempo de execução de Suplementos do Office sejam carregados antes de executar sua própria lógica personalizada.

## <a name="startup-of-a-content-or-task-pane-add-in"></a>Inicialização de um suplemento de conteúdo ou de painel de tarefas

A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento de conteúdo ou de painel de tarefas no Excel, no PowerPoint, no Project ou no Word.

![Flow de eventos ao iniciar um conteúdo ou um complemento do painel de tarefas.](../images/office15-app-sdk-loading-dom-agave-runtime.png)

Os eventos a seguir ocorrem quando um conteúdo ou um complemento do painel de tarefas é iniciado.

1. O usuário abre um documento que já contém um suplemento ou insere um suplemento no documento.

2. O Office de cliente lê o manifesto XML do add-in do AppSource, um catálogo de aplicativos no SharePoint ou o catálogo de pastas compartilhadas de onde ele se origina.

3. O Office cliente abre a página HTML do complemento em um controle do navegador.

    As próximas duas etapas, as etapas 4 e 5, ocorrem de forma assíncrona e em paralelo. Por esse motivo, o código do suplemento deve garantir que o DOM e o ambiente do tempo de execução do suplemento tenham terminado de carregar antes de prosseguir.

4. O controle do navegador carrega o corpo DOM e HTML e chama o manipulador de eventos para o `window.onload` evento.

5. O aplicativo cliente Office carrega o ambiente de tempo de execução Office, que baixa e armazena em cache os arquivos da biblioteca da API JavaScript do servidor de rede de distribuição de conteúdo (CDN) e chama o manipulador de eventos do complemento para o evento [de inicialização](/javascript/api/office#Office_initialize_reason_) do objeto [Office,](/javascript/api/office) se um manipulador tiver sido atribuído a ele. Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador. Para obter mais informações sobre a distinção `Office.initialize` entre e , consulte `Office.onReady` [Initialize your add-in](initialize-add-in.md).

6. Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.


## <a name="startup-of-an-outlook-add-in"></a>Inicialização de um suplemento do Outlook

A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento do Outlook em execução no desktop, tablet ou smartphone.

![Flow de eventos ao iniciar Outlook de um complemento.](../images/outlook15-loading-dom-agave-runtime.png)

Os eventos a seguir ocorrem quando um Outlook de usuário é iniciado.

1. Quando é iniciado, o Outlook lê os manifestos XML para suplementos do Outlook que foram instalados na conta de email do usuário.

2. O usuário seleciona um item no Outlook.

3. Se o item selecionado satisfizer as condições de ativação de um suplemento do Outlook, o Outlook ativará o suplemento e tornará seu botão visíveis na interface de usuário.

4. Se o usuário clicar no botão para iniciar o suplemento do Outlook, o Outlook abrirá a página HTML em um controle de navegador. As próximas duas etapas, as etapas 5 e 6, ocorrerem em paralelo.

5. O controle do navegador carrega o corpo DOM e HTML e chama o manipulador de eventos para o `onload` evento.

6. O Outlook carrega o ambiente de tempo de execução, que baixa e armazena em cache a API do JavaScript para arquivos da biblioteca a partir do servidor da rede de distribuição de conteúdo (CDN) e chama manipulador de eventos para o evento [initialize](/javascript/api/office#Office_initialize_reason_) do objeto do suplemento do [Office](/javascript/api/office). Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador. Para obter mais informações sobre a distinção `Office.initialize` entre e , consulte `Office.onReady` [Initialize your add-in](initialize-add-in.md).

7. Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.

## <a name="see-also"></a>Conferir também

- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [Inicialize seu suplemento do Office](initialize-add-in.md)
