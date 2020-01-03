---
ms.date: 11/04/2019
title: 'Tutorial: compartilhar dados e eventos entre as funções personalizadas do Excel e o painel de tarefas (versão prévia)'
ms.prod: excel
description: No Excel, compartilhe dados e eventos entre as funções personalizadas e o painel de tarefas.
localization_priority: Priority
ms.openlocfilehash: 16affeb29bd5950198f81f85e44adaf812067829
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814128"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a>Tutorial: compartilhar dados e eventos entre as funções personalizadas do Excel e o painel de tarefas (versão prévia)

As funções personalizadas do Excel e o painel de tarefas compartilham dados globais e podem fazer chamadas de função entre si. Para configurar o projeto para que as funções personalizadas possam funcionar com o painel de tarefas, siga as instruções neste artigo.

> [!NOTE]
> Os recursos descritos neste artigo estão em versão prévia e sujeitos a alterações. No momento, eles não têm suporte para utilização em ambientes de produção. Os recursos de versão prévia deste artigo só estão disponíveis no Excel no Windows. Para experimentar os recursos de versão prévia, você precisará [ingressar no Office Insider](https://insider.office.com/join).  Uma boa maneira de experimentar recursos de versão prévia é usar uma assinatura do Office 365. Caso ainda não tenha uma assinatura do Office 365, obtenha uma ingressando no [Programa para Desenvolvedores do Office 365](https://developer.microsoft.com/office/dev-program).

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

Use o gerador Yeoman para criar um projeto de suplemento do Excel. Execute o comando a seguir e responda às solicitações com as seguintes respostas:

```command&nbsp;line
yo office
```

- Escolha um tipo de projeto: **Projeto de suplemento de funções personalizadas do Excel**
- Escolha um tipo de script: **JavaScript**
- Qual será o nome do seu suplemento? **Meu suplemento do Office**

![Captura de tela das solicitações de resposta do seu Office para criar o projeto de suplemento.](../images/yo-office-excel-project.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

## <a name="configure-the-manifest"></a>Configurar o manifesto

1. Inicie o código do Visual Studio e abra o projeto **Meu suplemento do Office**.
2. Abra o arquivo **manifest.xml**.
3. Altere a seção `<Requirements>` para usar o **CustomFunctionsRuntime** versão **1.2**, como mostrado no código a seguir.
    
    ```xml
    <Requirements> 
    <Sets DefaultMinVersion="1.1">
    <Set Name="CustomFunctionsRuntime" MinVersion="1.2"/>
    </Sets>
    </Requirements>
    ```
    
4. No elemento `<Host>` da pasta de trabalho, adicione a seção `<Runtimes>` a seguir. O tempo de vida precisa ser **longo** para que as funções personalizadas ainda possam funcionar, mesmo quando o painel de tarefas estiver fechado.
    
    ```xml
    <Hosts>
    <Host xsi:type="Workbook">
    <Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
    </Runtimes>
    ```
    
5. No elemento `<Page>`, altere o local de origem de **Functions.Page.Url** para **TaskPaneAndCustomFunction.Url**.

    ```xml
    <AllFormFactors>
    ...
    <Page>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Page>
    ...
    ```

6. Na seção `<DesktopFormFactor>`, altere o **FunctionFile** de **Commands.Url** para usar **TaskPaneAndCustomFunction.Url**.
    
    ```xml
    <DesktopFormFactor>
    <GetStarted>
    ...
    </GetStarted>
    <FunctionFile resid="TaskPaneAndCustomFunction.Url"/>
    ```
    
7. Na seção `<Action>`, altere o local de origem de **Taskpane.Url** para **TaskPaneAndCustomFunction.Url**.
    
    ```xml
    <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ButtonId1</TaskpaneId>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Action>
    ```
    
8. Adicione uma nova **ID de Url** para **TaskPaneAndCustomFunction.Url** que aponta para **taskpane.html**.
     
    ```xml
    <bt:Urls>
    <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
    ...
    <bt:Url id="TaskPaneAndCustomFunction.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
    ...
    ```
    
9. Salve suas alterações e recompile o projeto.
    
    ```command&nbsp;line
    npm run build
    ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a>Compartilhar o estado entre as funções personalizadas e o código do painel de tarefas 

Agora que as funções personalizadas são executadas no mesmo contexto que o código do painel de tarefas, elas podem compartilhar o estado diretamente sem usar o objeto **Armazenamento**. As instruções a seguir mostram como compartilhar uma variável global entre as funções personalizadas e o código do painel de tarefas.

### <a name="create-custom-functions-to-get-or-store-shared-state"></a>Criar funções personalizadas para obter ou armazenar o estado compartilhado

1. No código do Visual Studio, abra o arquivo **src/functions/functions.js**.
2. Na linha 1, insira o código a seguir na parte superior. Isso inicializará uma variável global chamada **sharedState**.
    
    ```js
    window.sharedState = "empty";
    ```
    
3. Adicione o código a seguir para criar uma função personalizada que armazena valores para a variável **sharedState**.
    
    ```js
    /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
    function storeValue(sharedValue) {
    window.sharedState = sharedValue;
    return "value stored";
    }
    ```
    
4. Adicione o código a seguir para criar uma função personalizada que obtém o valor atual da variável **sharedState**.

    ```js
    /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
    function getValue() {
    return window.sharedState;
    }
    ```
    
5. Salve o arquivo.

### <a name="create-task-pane-controls-to-work-with-global-data"></a>Criar controles do painel de tarefas para trabalhar com dados globais 

1. Abra o arquivo**src/taskpane/taskpane.html**.
2. Após o elemento de fechamento `</main>`, adicione o seguinte HTML. O HTML cria duas caixas de texto e botões usados para obter ou armazenar dados globais.

    ```html
    <ol>
    <li>Enter a value to send to the custom function and select <strong>Store</strong>.</li>
    <li>Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve it.</li>
    <li>To send data to the task pane, in a cell, enter <strong>=CONTOSO.STOREVALUE("new value")</strong></li>
    <li>Select <strong>Get</strong> to display the value in the task pane.</li>
    </ol>
    <p>Store new value to shared state</p>
    <div>
    <input type="text" id="storeBox" />
    <button onclick="storeSharedValue()">Store</button>
    </div>
     
    <p>Get shared state value</p>
    <div>
    <input type="text" id="getBox" />
    <button onclick="getSharedValue()">Get</button>
    </div>
    ```
    
3. Antes do elemento `<body>`, adicione o seguinte script. Esse código manipulará os eventos de clique do botão quando o usuário desejar armazenar ou obter os dados globais.
    
    ```js
    <script>
    function storeSharedValue() {
    let sharedValue = document.getElementById('storeBox').value;
    window.sharedState = sharedValue;
    }
    
    function getSharedValue() {
    document.getElementById('getBox').value = window.sharedState;
    }</script>
    ```
    
4. Salve o arquivo.
5. Compilar o projeto
    
    ```command&nbsp;line
    npm run build 
    ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a>Experimente compartilhar dados entre as funções personalizadas e o painel de tarefas

- Inicie o projeto usando o comando a seguir.

    ```command&nbsp;line
    npm run start
    ```

Após a inicialização do Excel, você pode usar os botões do painel de tarefas para armazenar ou obter os dados compartilhados. Insira `=CONTOSO.GETVALUE()` em uma célula para que a função personalizada recupere os mesmos dados compartilhados. Ou use `=CONTOSO.STOREVALUE(“new value”)` para alterar os dados compartilhados para um novo valor.

> [!NOTE]
> A configuração do seu projeto, como mostrado neste artigo, compartilhará o contexto entre as funções personalizadas e o painel de tarefas. Não há suporte para chamar APIs do Office a partir de funções personalizadas na visualização.

