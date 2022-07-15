---
title: Crie o seu primeiro suplemento do painel de tarefas do Project
description: Saiba como criar um Suplemento do Excel simples usando a API JS do Office.
ms.date: 07/13/2022
ms.prod: project
ms.localizationpriority: high
ms.openlocfilehash: c2f0e31b5a4c958cd155dfeb6d1648f7a2697c69
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797474"
---
# <a name="build-your-first-project-task-pane-add-in"></a>Crie o seu primeiro suplemento do painel de tarefas do Project

Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Project.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project 2016 ou posterior no Windows

## <a name="create-the-add-in"></a>Criar o suplemento

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project`
- **Escolha o tipo de script:** `Javascript`
- **Qual será o nome do suplemento?** `My Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `Project`

![Os prompts e respostas para o gerador Yeoman em uma interface de linha de comando.](../images/yo-office-project.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explore o projeto

O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico.

- O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.
- O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.
- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.
- O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo cliente do Office. Neste início rápido, o código define o campo `Name` e o campo `Notes` da tarefa selecionada de um projeto.

## <a name="try-it-out"></a>Experimente

1. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Inicie o servidor Web local.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    Execute o seguinte comando no diretório raiz do seu projeto. O servidor Web local é iniciado quando este comando é executado.

    ```command&nbsp;line
    npm run dev-server
    ```

1. Em Project, crie um plano de projeto simples.

1. Carregue seu suplemento no Project seguindo as instruções em [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

1. Selecione uma única tarefa dentro do projeto.

1. Na parte inferior do painel de tarefas, escolha o link **Executar** para renomear a tarefa selecionada e adicionar anotações à tarefa selecionada.

    ![O aplicativo Project com o suplemento do painel de tarefas carregado.](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do painel de tarefas do Project! Em seguida, saiba mais sobre os recursos de um suplemento do Project e explore os cenários comuns.

> [!div class="nextstepaction"]
> [Suplementos do Project](../project/project-add-ins.md)

## <a name="see-also"></a>Confira também

- [Desenvolver Suplementos do Office](../develop/develop-overview.md)
- [Principais conceitos dos Suplementos do Office](../overview/core-concepts-office-add-ins.md)
- [Usando Visual Studio Code para publicar](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
