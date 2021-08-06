---
title: Criar um suplemento do painel de tarefas do Excel usando o React
description: Aprenda a criar um suplemento do painel de tarefas simples do Excel usando a API JS do Office e o React.
ms.date: 08/04/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 74a28f3914ddbc54188d3b8baa33fc1faa7a30fe
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773389"
---
# <a name="use-react-to-build-an-excel-task-pane-add-in"></a>Criar um suplemento do painel de tarefas do Excel usando o React

Neste artigo, você passará pelo processo de criar um suplemento do painel de tarefas do Excel usando o React e a API JavaScript do Excel.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project using React framework`
- **Escolha o tipo de script:** `TypeScript`
- **Qual será o nome do suplemento?** `My Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `Excel`

![Captura de tela da interface de linha de comando do gerador do suplemento do Yeoman Office, com o tipo de projeto definido para a estrutura React.](../images/yo-office-excel-react-2.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explore o projeto

O projeto de suplemento que você criou com o gerador Yeoman contém um código de exemplo para um suplemento básico do painel de tarefas. Se você quiser examinar os principais componentes do seu projeto de suplemento, abra o projeto no seu editor de código e revise os arquivos listados abaixo. Quando estiver pronto para experimentar o suplemento, prossiga para a próxima seção.

- O arquivo **manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento. Para saber mais sobre o arquivo **manifest.xml**, consulte [manifesto XML de suplementos do Office](../develop/add-in-manifests.md).
- O arquivo **./src/taskpane/taskpane.html** define a estrutura HTML do painel de tarefas e os arquivos na pasta **./src/taskpane/components** definem as diversas partes da interface do usuário do painel de tarefas.
- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.
- O arquivo **./src/taskpane/components/App.tsx** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o Excel.

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

Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel usando o React! Em seguida, saiba mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo o tutorial de suplemento do Excel.

> [!div class="nextstepaction"]
> [Tutorial de suplemento do Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>Confira também

- [Tutorial de suplemento do Excel](../tutorials/excel-tutorial.md)
- [Modelo de objeto JavaScript do Excel em Suplementos do Office](../excel/excel-add-ins-core-concepts.md)
- [Exemplos de código do suplemento do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)