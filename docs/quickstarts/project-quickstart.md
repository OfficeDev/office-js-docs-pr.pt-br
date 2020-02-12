---
title: Crie o seu primeiro suplemento do painel de tarefas do Project
description: Saiba como criar um Suplemento do Excel simples usando a API JS do Office.
ms.date: 01/16/2020
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 821cdc9f32b0fbc2b48e2a92259f340e65a03f64
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950618"
---
# <a name="build-your-first-project-task-pane-add-in"></a>Crie o seu primeiro suplemento do painel de tarefas do Project

Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Project.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project 2016 ou posterior no Windows

## <a name="create-the-add-in"></a>Criar o suplemento

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project`
- **Escolha o tipo de script:** `Javascript`
- **Qual será o nome do suplemento?** `My Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `Project`

![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-project.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explore o projeto

O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico. 

- O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.
- O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.
- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.
- O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.

## <a name="update-the-code"></a>Atualizar o código

No seu editor de código, abra o arquivo **./src/taskpane/taskpane.js** e adicione o seguinte código dentro da função **executar**. Este código usa a API JavaScript do Office para configurar o `Name` campo e `Notes` campo da tarefa selecionada.

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## <a name="try-it-out"></a>Experimente

1. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Inicie o servidor Web local.

    > [!NOTE]
    > Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

    Execute o seguinte comando no diretório raiz do seu projeto. O servidor Web local é iniciado quando este comando é executado.

    ```command&nbsp;line
    npm start
    ```

3. Em Project, crie um plano de projeto simples.

4. Carregue seu suplemento no Project seguindo as instruções em [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

5. Selecione uma única tarefa dentro do projeto.

6. Na parte inferior do painel de tarefas, escolha o link **Executar** para renomear a tarefa selecionada e adicionar anotações à tarefa selecionada.

    ![Captura de tela do aplicativo Project com o suplemento do painel de tarefas carregado](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do painel de tarefas do Project! Em seguida, saiba mais sobre os recursos de um suplemento do Project e explore os cenários comuns.

> [!div class="nextstepaction"]
> [Suplementos do Project](../project/project-add-ins.md)

## <a name="see-also"></a>Confira também

- [Criando Suplementos do Office ](../overview/office-add-ins-fundamentals.md)
- [Principais conceitos dos Suplementos do Office](../overview/core-concepts-office-add-ins.md)
- [Desenvolver Suplementos do Office ](../develop/develop-overview.md)
