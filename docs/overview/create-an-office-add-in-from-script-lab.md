---
title: Criar um suplemento autônomo do Office com base em seu código do Script Lab
description: Saiba como mover seu trecho do Script Lab para um projeto Yo Office
ms.topic: how-to
ms.date: 04/07/2022
ms.localizationpriority: high
ms.openlocfilehash: 725ce9b44c55b46e6d0ab0c085973947fcf88201
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810145"
---
# <a name="create-a-standalone-office-add-in-from-your-script-lab-code"></a>Criar um suplemento autônomo do Office com base em seu código do Script Lab

Se você criou um trecho do Script Lab, convém transformá-lo em um suplemento autônomo. É possível copiar o código do Script Lab para um projeto gerado pelo [Gerador Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md) (também chamado de "Yo Office"). Em seguida, é possível continuar desenvolvendo o código como um suplemento que pode eventualmente implantar em outras pessoas.

As etapas neste artigo referem-se ao [Visual Studio Code](https://code.visualstudio.com/), mas é possível usar qualquer editor de código que preferir.

## <a name="create-a-new-yo-office-project"></a>Criar um novo projeto Yo Office

É necessário criar o projeto de suplemento autônomo que será o novo local de desenvolvimento para o código do trecho.

Execute o comando `yo office --projectType taskpane --ts true --host <host> --name "basic-sample"`, onde `<host>` é um dos seguintes valores.

- excel
- outlook
- powerpoint
- palavra

> [!IMPORTANT]
> O `--name` de argumento deve estar entre aspas duplas, mesmo que não tenha espaços.

O comando anterior cria uma nova pasta de projeto chamada **basic-sample**. Ele está configurado para ser executado no host especificado e usa TypeScript. O Script Lab usa TypeScript por padrão, mas a maioria dos trechos é JavaScript. É possível criar um projeto Yo Office JavaScript, se preferir, mas certifique-se de que qualquer código que você copiar seja JavaScript.

## <a name="open-the-snippet-in-script-lab"></a>Abra o trecho no Script Lab

Use um trecho existente no Script Lab para saber como copiar um trecho para um projeto gerado pelo Yo Office.

1. Abra o Office (Word, Excel, PowerPoint ou Outlook) e abra o Script Lab.
1. Selecione **Script Lab** > **Código**. Se você estiver trabalhando no Outlook, abra uma mensagem de email para ver o Script Lab na faixa de opções.
1. No painel de tarefas do Script Lab, escolha **Exemplos**. Em seguida, selecione um exemplo básico com base em qual host do Office você está trabalhando.
    - Para o Excel ou o Word, escolha o exemplo de **Chamada à API Básica (TypeScript)**.
    - Para o Outlook, escolha o exemplo **Usar configurações de suplemento**.
    - Para o PowerPoint, escolha o exemplo de **Chamada à API Básica (Office 2013)**.

## <a name="copy-snippet-code-to-visual-studio-code"></a>Copiar o código de trecho para o código do Visual Studio

Agora é possível copiar o código do trecho para o projeto Yo Office no VS Code.

- No VS Code, abra o projeto **basic-sample**.

Nas próximas etapas, você copiará o código de várias guias no Script Lab.

:::image type="content" source="../images/script-lab-script-tabs.png" alt-text="Captura de tela de guias no Script Lab.":::

### <a name="copy-task-pane-code"></a>Copiar código do painel de tarefas

1. No VS Code, abra o arquivo **/src/taskpane/taskpane.ts**. Se você estiver usando um projeto JavaScript, o nome do arquivo será **taskpane.js**.
1. No Script Lab, selecione a guia **Script**.
1. Copie todo o código na guia **Script** para a área de transferência. Substitua todo o conteúdo de **taskpane.ts** (ou **taskpane.js** para JavaScript) pelo código copiado.

### <a name="copy-task-pane-html"></a>Copiar HTML do painel de tarefas

1. No VS Code, abra o arquivo **/src/taskpane/taskpane.html**.
1. No Script Lab, selecione a guia **HTML**.
1. Copie todo o HTML na guia **HTML** para a área de transferência. Substitua todo o HTML dentro da marca `<body>` pelo HTML que você copiou.

### <a name="copy-task-pane-css"></a>Copiar CSS do painel de tarefas

1. No VS Code, abra o arquivo **/src/taskpane/taskpane.css**.
1. No Script Lab, selecione a guia **CSS**.
1. Copie todo o CSS na guia **CSS** para a área de transferência. Substitua todo o conteúdo de **taskpane.css** pelo CSS copiado.
1. Salve todas as alterações nos arquivos que você atualizou nas etapas anteriores.

## <a name="add-jquery-support"></a>Adicionar suporte ao jQuery

O Script Lab usa jQuery nos trechos. É necessário adicionar essa dependência ao projeto Yo Office para executar o código com êxito.

1. Abra o arquivo **taskpane.html** e adicione a seguinte marca de script à seção `<head>`.

    ```html
     <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-3.3.1.js"></script>
    ```

    > [!NOTE]
    > A versão específica do jQuery pode variar. É possível determinar qual versão Script Lab está sendo usada escolhendo a guia **Bibliotecas**.

1. Abra um terminal no VS Code e insira os comandos a seguir.

    ```command&nbsp;line
    npm install --save-dev jquery@3.1.1
    npm install --save-dev @types/jquery@3.3.1
    ```

Se você criou um trecho que tem dependências de biblioteca adicionais, certifique-se de adicioná-las ao projeto Yo Office. Localize uma lista de todas as dependências de biblioteca na guia **Bibliotecas** no Script Lab.

## <a name="handle-initialization"></a>Inicialização do identificador

O Script Lab lida com a inicialização `Office.onReady` automaticamente. Será necessário modificar o código para fornecer seu próprio identificador `Office.onReady`.

1. Abra o arquivo **taskpane.ts** (ou **outaskpane.js** para JavaScript).
1. Para Excel ou Word, substitua:

    ```typescript
    $("#run").click(() => tryCatch(run));
    ```

    por:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(() => tryCatch(run));
      });
    });
    ```

1. Para Outlook, substitua:

    ```typescript
    $("#get").click(get);
    $("#set").click(set);
    $("#save").click(save);
    ```

    por:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#get").click(get);
        $("#set").click(set);
        $("#save").click(save);
      });
    });
    ```

1. Para PowerPoint, substitua:

    ```typescript
    $("#run").click(run);
    ```

    por:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(run);
      });
    });
    ```

1. Salve o arquivo.

## <a name="custom-functions"></a>Funções personalizadas

Se o trecho usar funções personalizadas, será necessário usar o modelo de funções personalizadas do Yo Office. Para transformar funções personalizadas em um suplemento autônomo, siga estas etapas.

1. Execute o comando `yo office --projectType excel-functions --ts true --name "functions-sample"`.

    > [!IMPORTANT]
    > O `--name` de argumento deve estar entre aspas duplas, mesmo que não tenha espaços.

1. Abra p Excel e, em seguida, abra p Script Lab.
1. Selecione **Script Lab** > **Código**.
1. No painel de tarefas do Script Lab, escolha **Exemplos** e, em seguida, escolha o exemplo de **Função personalizada básica**.
1. Abra o arquivo **/src/functions/functions.ts**. Se você estiver usando um projeto JavaScript, o nome do arquivo será **functions.js**.
1. No Script Lab, selecione a guia **Script**.
1. Copie todo o código na guia **Script** para a área de transferência. Cole o código na parte superior do **functions.ts** (ou **functions.js** para JavaScript) com o código copiado.
1. Salve o arquivo.

## <a name="test-the-standalone-add-in"></a>Testar o suplemento autônomo

Depois que todas as etapas forem concluídas, execute e teste seu suplemento autônomo. Execute o comando a seguir para começar.

```command&nbsp;line
npm start
```

O Office será iniciado e será possível abrir o painel de tarefas do suplemento na faixa de opções. Parabéns! Agora é possível continuar criando seu suplemento como um projeto autônomo.

## <a name="console-logging"></a>Log do console

Muitos trechos no Script Lab gravam a saída em uma seção de console na parte inferior do painel de tarefas. O projeto Yo Office não tem uma seção de console. Todas as instruções `console.log*` serão gravadas no console de depuração padrão (como as ferramentas de desenvolvedor do navegador). Se você quiser que a saída vá para o painel de tarefas, será necessário atualizar o código.
