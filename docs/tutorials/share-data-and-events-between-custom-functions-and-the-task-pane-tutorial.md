---
title: 'Tutorial: Compartilhar dados e eventos entre as funções personalizadas do Excel e do painel de tarefas'
description: Aprenda como compartilhar dados e eventos no Excel entre as funções personalizadas e o painel de tarefas.
ms.date: 08/04/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: de27ff675e8ef757e0b4b7c95a74a061e9cadee586ae6b7134b68c16184fdf9c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098539"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>Tutorial: Compartilhar dados e eventos entre as funções personalizadas do Excel e do painel de tarefas

Você pode configurar o suplemento do Excel para usar um tempo de execução compartilhado. Isso permite compartilhar dados globais ou enviar eventos entre o painel de tarefas e as funções personalizadas.

Para a maioria dos cenários de funções personalizadas, recomendamos usar um tempo de execução compartilhada, a menos que você tenha uma razão específica para usar uma função personalizada fora do painel de tarefa (sem IU).

Este tutorial presume que você esteja familiarizado com o uso do gerador Yo do Office para criar adicionais no projetos de. Considere concluir o [Tutorial de funções personalizadas do Excel](excel-tutorial-create-custom-functions.md), se ainda não o fez.

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

Use o gerador Yeoman para criar um projeto de suplemento do Excel. Execute o comando a seguir e responda aos prompts com as respostas a seguir.

```command line
yo office
```

- Escolha um tipo de projeto: **Projeto de suplemento de funções personalizadas do Excel**
- Escolha um tipo de script: **JavaScript**
- Qual será o nome do seu suplemento? **Meu suplemento do Office**

![Captura de tela mostrando os prompts e respostas para o gerador do Yeoman em uma interface de linha de comando.](../images/yo-office-excel-project.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

## <a name="configure-the-manifest"></a>Configurar o manifesto

1. Inicie o código do Visual Studio e abra o projeto **Meu suplemento do Office**.
2. Abra o arquivo **manifest.xml**.
3. Localize a seção `<VersionOverrides>` e adicione a seguinte seção `<Runtimes>`. O tempo de vida precisa ser **longo** para que as funções personalizadas ainda possam funcionar, mesmo quando o painel de tarefas estiver fechado.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

> [!NOTE]
> Se seu suplemento inclui o elemento `Runtimes` no manifesto, ele utiliza o Internet Explorer 11 independentemente da versão do Windows ou do Microsoft 365. Para mais informações, consulte [Runtimes](../reference/manifest/runtimes.md).

4. No elemento `<Page>`, altere o local de origem de **Functions.Page.Url** para **ContosoAddin.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. Na seção `<DesktopFormFactor>`, altere o **FunctionFile** de **Commands.Url** para usar **ContosoAddin.Url**.

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. Na seção `<Action>`, altere o local de origem de **Taskpane.Url** para **ContosoAddin.Url**.

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. Adicione um novo **ID de URL** para **ContosoAddin.Url** que aponte para **taskpane.html**.

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. Salve suas alterações e recompile o projeto.

   ```command line
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

1. Abra o arquivo **src/taskpane/taskpane.html**.
2. Adicionar o seguinte elemento do roteiro pouco antes do elemento `</head>` de fechamento.

   ```html
   <script src="functions.js"></script>
   ```

3. Após o elemento de fechamento `</main>`, adicione o seguinte HTML. O HTML cria duas caixas de texto e botões usados para obter ou armazenar dados globais.

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
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

4. Antes do elemento `</body>` fechamento, adicionar o script a seguir. Esse código manipulará os eventos de clique do botão quando o usuário desejar armazenar ou obter os dados globais.

   ```js
   <script>
   function storeSharedValue() {
     let sharedValue = document.getElementById('storeBox').value;
     window.sharedState = sharedValue;
   }

   function getSharedValue() {
     document.getElementById('getBox').value = window.sharedState;
   }
   </script>
   ```

5. Salve o arquivo.
6. Compilar o projeto

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a>Experimente compartilhar dados entre as funções personalizadas e o painel de tarefas

- Inicie o projeto usando o comando a seguir.

  ```command line
  npm run start
  ```

Após a inicialização do Excel, você pode usar os botões do painel de tarefas para armazenar ou obter os dados compartilhados. Insira `=CONTOSO.GETVALUE()` em uma célula para que a função personalizada recupere os mesmos dados compartilhados. Ou use `=CONTOSO.STOREVALUE("new value")` para alterar os dados compartilhados para um novo valor.

> [!NOTE]
> A configuração do seu projeto, como mostrado neste artigo, compartilhará o contexto entre as funções personalizadas e o painel de tarefas. É possível chamar algumas APIs do Office a partir de funções personalizadas. [Consulte chamada de APIs do Microsoft Excel a partir de uma função personalizada](../excel/call-excel-apis-from-custom-function.md) para mais detalhes.
