---
title: Crie seu primeiro suplemento do painel de tarefas do OneNote
description: Saiba como criar um suplemento do painel de tarefas do OneNote simples usando a API JS do Office.
ms.date: 06/10/2022
ms.prod: onenote
ms.localizationpriority: high
ms.openlocfilehash: 9b5f4dd941ed8cc107bee04bc67a368520439948
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66090849"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a>Crie seu primeiro suplemento do painel de tarefas do OneNote

Neste artigo, você verá o processo de criação de um suplemento do painel de tarefas do OneNote.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project`
- **Escolha o tipo de script:** `Javascript`
- **Qual será o nome do suplemento?** `My Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `OneNote`

![Captura de tela apresentando os avisos e respostas do gerador Yeoman em uma interface de linha de comando.](../images/yo-office-onenote.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explore o projeto

O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico.

- O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.
- O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.
- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.
- O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo cliente do Office.

## <a name="update-the-code"></a>Atualizar o código

No seu editor de código, abra o arquivo **./src/taskpane/taskpane.js** e adicione o seguinte código na função `run`. Este código usa a API JavaScript do OneNote para definir o título da página e adicionar um contorno ao corpo da página.

```js
try {
    await OneNote.run(async (context) => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands.
        await context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## <a name="try-it-out"></a>Experimente

1. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Inicie o servidor Web local. Execute o seguinte comando no diretório raiz do seu projeto.

    ```command&nbsp;line
    npm run dev-server
    ```

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

1. No [OneNote Online](https://www.onenote.com/notebooks), abra um bloco de anotações e crie uma nova página.

1. Escolha **Inserir > Suplementos do Office** para abrir a caixa de diálogo Suplementos do Office.

    - Se você estiver conectado à sua conta de consumidor, selecione a guia **MEUS SUPLEMENTOS** e escolha  **Carregar Meu Suplemento**.

    - Se você estiver conectado com a sua conta corporativa ou de estudante, selecione a guia **MINHA ORGANIZAÇÃO** e escolha **Carregar Meu Suplemento**.

    A imagem a seguir mostra a guia **MEUS SUPLEMENTOS** para blocos de anotações do consumidor.

    ![A caixa de diálogo Suplementos do Office mostrando a guia MEUS SUPLEMENTOS.](../images/onenote-office-add-ins-dialog.png)

1. Na caixa de diálogo Carregar Suplemento, navegue até **manifest.xml** na pasta do projeto e escolha **Carregar**.

1. Na guia **Página Inicial**, na faixa de opções, escolha o botão **Mostrar Painel de Tarefas**. O painel de tarefa do suplemento abre em um iFrame ao lado da página do OneNote.

1. Na parte inferior do painel de tarefas, escolha o link **Executar** para definir o título da página e adicionar um contorno ao corpo da página.

    ![Captura de tela apresentando o suplemento criado com base nesse passo a passo: exibir o painel de opções do painel de tarefas e o painel de tarefas no OneNote.](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do painel de tarefas do OneNote! A seguir, saiba mais sobre os principais conceitos da criação de suplementos do OneNote.

> [!div class="nextstepaction"]
> [Visão geral da programação da API JavaScript do OneNote](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Desenvolver Suplementos do Office ](../develop/develop-overview.md)
- [Visão geral da programação da API JavaScript do OneNote](../onenote/onenote-add-ins-programming-overview.md)
- [Referência da API JavaScript do OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Usando o Visual Studio Code para publicar](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)