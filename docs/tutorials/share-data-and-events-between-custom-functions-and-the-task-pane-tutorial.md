---
title: 'Tutorial: Compartilhar dados e eventos entre as funções personalizadas do Excel e do painel de tarefas'
description: Aprenda como compartilhar dados e eventos no Excel entre as funções personalizadas e o painel de tarefas.
ms.date: 06/15/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: b61ac6305586e5de2f53a0950fd6a52a0503eafd
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958720"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>Tutorial: Compartilhar dados e eventos entre as funções personalizadas do Excel e do painel de tarefas

Compartilhe dados globais e envie eventos entre o painel de tarefas e funções personalizadas do suplemento do Excel com um runtime compartilhado. É recomendável usar um runtime compartilhado para a maioria dos cenários de funções personalizadas, a menos que você tenha um motivo específico para usar um suplemento personalizado somente de função. Este tutorial pressupõe que você esteja familiarizado com o uso do [Gerador do Yeoman para Suplementos do Office](../develop/yeoman-generator-overview.md) para criar projetos de suplementos. Considere concluir o [Tutorial de funções personalizadas do Excel](excel-tutorial-create-custom-functions.md), se ainda não o fez.

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

Use o [Gerador Yeoman para Suplementos do Office](../develop/yeoman-generator-overview.md) para criar o projeto de suplemento do Excel.

- Para gerar um suplemento do Excel com funções personalizadas, execute o comando.

    ```command&nbsp;line
    yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true
    ```

O gerador cria o projeto e instala componentes do Node com suporte.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Siga estas etapas para configurar o projeto de suplemento para usar um runtime compartilhado.

1. Inicie Visual Studio Code e abra o projeto de suplemento gerado.
1. Abra o arquivo **manifest.xml**.
1. Substitua (ou adicione) a seguinte **\<Requirements\>** seção XML para exigir o [ conjunto de requisitos de tempo de execução compartilhado](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets).

    ```xml
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    ```

    Após a atualização, o XML do manifesto deverá aparecer na ordem a seguir.

    ```xml
    <Hosts>
      <Host Name="..."/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. Encontre a seção **\<VersionOverrides\>** e adicione a seguinte seção **\<Runtimes\>**. A vida útil deve ser **longa** para que o código do suplemento possa ser executado mesmo quando o painel de tarefas está fechado. O `resid`valor é **Taskpane.Url**, que faz referência ao local do arquivo **taskpane.html** especificado na `<bt:Urls>`seção próxima à parte inferior do arquivo **manifest.xml**.

    ```xml
    <Runtimes>
      <Runtime resid="Taskpane.Url" lifetime="long" />
    </Runtimes>
    ```

    > [!IMPORTANT]
    > A seção **\<Runtimes\>** deve ser inserida após o elemento `<Host xsi:type="...">` na ordem exata mostrada no XML a seguir.

    ```xml
    <VersionOverrides ...>
      <Hosts>
        <Host xsi:type="...">
          <Runtimes>
            <Runtime resid="Taskpane.Url" lifetime="long" />
          </Runtimes>
        ...
        </Host>
    ```

    > [!NOTE]
    > Se o suplemento incluir o elemento `Runtimes` no manifesto (necessário para um runtime compartilhado) e as condições para usar o Microsoft Edge com WebView2 (baseado em Chromium) forem atendidas, ele usará esse controle WebView2. Se as condições não forem atendidas, ele usará o Internet Explorer 11, independentemente da versão do Windows ou Microsoft 365. Para obter mais informações, consulte [Runtimes](/javascript/api/manifest/runtimes) e [Navegadores usados pelos suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

1. Localize o elemento **\<Page\>**. Em seguida, altere o local de origem de **Functions.Page.Url** para **Taskpane.Url**.

    ```xml
    <AllFormFactors>
    ...
    <Page>
      <SourceLocation resid="Taskpane.Url"/>
    </Page>
    ...
    ```

1. Localize a marca`<FunctionFile ...>` e altere o `resid` de **Commands.Url** para **Taskpane.Url**.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Salve o arquivo **manifest.xml**.

## <a name="configure-the-webpackconfigjs-file"></a>Configurar o arquivo webpack.config.js

O **webpack.config.js** construirá vários carregadores de tempo de execução. É necessário modificá-lo para carregar apenas o tempo de execução JavaScript compartilhado por meio do arquivo **taskpane.html**.

1. Abra o arquivo **webpack.config.js**.
1. Vá para seção `plugins:`.
1. Remova o seguinte plugin `functions.html`, se ele existir.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. Remova o seguinte plugin `commands.html`, se ele existir.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. Se você removeu os plugins`functions`ou `commands`, adicione-os como `chunks`. O JavaScript a seguir mostra a entrada atualizada se você removeu os plugins `functions` e `commands`.

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Salvar suas alterações e reconstrua o projeto.

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Você também pode remover os arquivos **functions.html** e **commands.html**. O **taskpane.htm** l carregará o código **functions.js** e **commands.js** no tempo de execução JavaScript compartilhado por meio das atualizações do webpack que você acabou de fazer.

1. Salve suas alterações e execute o projeto. Verifique se ele é carregado e executado sem erros.

   ```command&nbsp;line
   npm run start
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a>Compartilhar o estado entre as funções personalizadas e o código do painel de tarefas

Agora que as funções personalizadas são executadas no mesmo contexto que o código do painel de tarefas, elas podem compartilhar o estado diretamente sem usar o objeto **Armazenamento**. As instruções a seguir mostram como compartilhar uma variável global entre as funções personalizadas e o código do painel de tarefas.

### <a name="create-custom-functions-to-get-or-store-shared-state"></a>Criar funções personalizadas para obter ou armazenar o estado compartilhado

1. No código do Visual Studio, abra o arquivo **src/functions/functions.js**.
1. Na linha 1, insira o código a seguir na parte superior. Isso inicializará uma variável global chamada **sharedState**.

    ```js
    window.sharedState = "empty";
    ```

1. Adicione o código a seguir para criar uma função personalizada que armazena valores para a variável **sharedState**.

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

1. Adicione o código a seguir para criar uma função personalizada que obtém o valor atual da variável **sharedState**.

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

1. Salve o arquivo.

### <a name="create-task-pane-controls-to-work-with-global-data"></a>Criar controles do painel de tarefas para trabalhar com dados globais

1. Abra o arquivo **src/taskpane/taskpane.html**.
1. Adicionar o seguinte elemento do roteiro pouco antes do elemento `</head>` de fechamento.

    ```HTML
    <script src="../functions/functions.js"></script>
    ```

1. Após o elemento de fechamento `</main>`, adicione o seguinte HTML. O HTML cria duas caixas de texto e botões usados para obter ou armazenar dados globais.

    ```HTML
    <ol>
      <li>
        Enter a value to send to the custom function and select
        <strong>Store</strong>.
      </li>
      <li>
        Enter <strong>=CONTOSO.GETVALUE()</strong> into a cell to retrieve it.
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

1. Antes do elemento `</body>`, adicione o script a seguir. Esse código manipulará os eventos de clique do botão quando o usuário quiser armazenar ou obter dados globais.

    ```HTML
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

1. Salve o arquivo.
1. Compile o projeto.

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

## <a name="see-also"></a>Confira também

- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
