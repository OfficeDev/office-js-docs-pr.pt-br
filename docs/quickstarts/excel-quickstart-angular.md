---
title: Criar um suplemento do painel de tarefas do Excel usando o Angular
description: Aprenda a criar um suplemento do painel de tarefas simples do Excel usando a API do Office JS e o lado a lado.
ms.date: 06/10/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 372649188d8f617f65e0c2eddc4d758047b1a2cc
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091150"
---
# <a name="use-angular-to-build-an-excel-task-pane-add-in"></a>Criar um suplemento do painel de tarefas do Excel usando o Angular

Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Excel usando o Angular e a API JavaScript do Excel.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project using Angular framework`
- **Escolha o tipo de script:** `TypeScript`
- **Qual será o nome do suplemento?** `My Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `Excel`

![Captura de tela da interface de linha de comando do gerador de Suplemento do Yeoman Office, com tipo de projeto definido para a estrutura Angular.](../images/yo-office-excel-angular-2.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explore o projeto

O projeto de suplemento que você criou com o gerador Yeoman contém um código de exemplo para um suplemento básico do painel de tarefas. Se você quiser examinar os principais componentes do seu projeto de suplemento, abra o projeto no seu editor de código e revise os arquivos listados abaixo. Quando estiver pronto para experimentar o suplemento, prossiga para a próxima seção.

- O arquivo **manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento. Para saber mais sobre o arquivo **manifest.xml**, consulte [manifesto XML de suplementos do Office](../develop/add-in-manifests.md).
- O arquivo **./src/taskpane/app/app.component.html** contém a marcação HTML do painel de tarefas.
- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.
- O arquivo **./src/taskpane/app/app.component.ts** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o Excel.

## <a name="try-it-out"></a>Experimente

1. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

1. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Captura de tela do menu da página inicial do Excel, com o botão Mostrar Painel de Tarefas realçado.](../images/excel-quickstart-addin-3b.png)

1. Selecione um intervalo de células na planilha.

1. Na parte inferior do painel de tarefas, escolha o link **Executar** para definir a cor do intervalo selecionado como amarelo.

    ![Captura de tela do Excel, com o painel de tarefas do suplemento aberto e o botão Executar realçado no painel de tarefas do suplemento.](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel usando o Angular! A seguir, saiba mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo o tutorial de suplemento do Excel.

> [!div class="nextstepaction"]
> [Tutorial de suplemento do Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Desenvolver Suplementos do Office](../develop/develop-overview.md)
- [Modelo de objeto JavaScript do Excel em Suplementos do Office](../excel/excel-add-ins-core-concepts.md)
- [Exemplos de código do suplemento do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Usando Visual Studio Code para publicar](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
