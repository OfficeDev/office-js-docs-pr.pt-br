---
title: Carregar o ambiente de tempo de execução e DOM
description: ''
ms.date: 01/23/2018
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
    
5. O aplicativo host do Office carrega o ambiente de tempo de execução, que baixa e armazena em cache a API do JavaScript para arquivos da biblioteca a partir do servidor da rede de distribuição de conteúdo (CDN) e chama manipulador de eventos do suplemento para o evento [initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) do objeto [Office](https://dev.office.com/reference/add-ins/shared/office).
    
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
    
6. O Outlook chama o manipulador de eventos para o evento [initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) do objeto [Office](https://dev.office.com/reference/add-ins/shared/office) do suplemento.
    
7. Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.
    

## <a name="checking-the-load-status"></a>Verificar o status de carregamento


Uma maneira de verificar se o ambiente de tempo de execução e o DOM concluíram o carregamento é usar a função [.ready()](http://api.jquery.com/ready/) do jQuery: `$(document).ready()`. Por exemplo, a seguinte função do manipulador de eventos **initialize** garante que o DOM seja carregado antes do código específico para inicializar as execuções de suplementos. Subsequentemente, o manipulador de eventos **inicializar** prossegue e usa a propriedade [mailbox.item](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) para obter o item selecionado atual no Outlook, e chama a função principal do suplemento, `initDialer`.


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

Essa mesma técnica pode ser usada no manipulador **initialize** de qualquer Suplemento do Office.

O suplemento do Outlook de amostra de discagem telefônica mostra uma abordagem ligeiramente diferente usando somente o JavaScript para verificar essas mesmas condições. 

> [!IMPORTANT]
> Mesmo que o suplemento não tenha tarefas de inicialização para executar, você deve incluir pelo menos uma função mínima do manipulador de eventos **Office.initialize**, como mostra o exemplo a seguir.

```js
Office.initialize = function () {
};
```

Se você não incluir um manipulador de eventos **Office.initialize**, o suplemento poderá gerar um erro ao ser iniciado. Além disso, se um usuário tentar usar o suplemento com um cliente virtual do Office Online, como o Excel Online, PowerPoint Online ou Outlook Web App, ele não funcionará.

Se o suplemento incluir mais de uma página, essa página deverá incluir ou chamar um manipulador de eventos **Office.initialize** sempre que uma nova página for carregada.


## <a name="see-also"></a>Veja também

- [Noções básicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md)
    
