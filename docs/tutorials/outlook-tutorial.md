---
title: 'Tutorial: criar uma mensagem para compor o suplemento do Outlook'
description: Neste tutorial, você criará um suplemento do Outlook que insere Gists do GitHub no corpo de uma nova mensagem.
ms.date: 02/23/2022
ms.prod: outlook
ms.localizationpriority: high
ms.openlocfilehash: 987084c16f3e8f1af1809866ac248b4f1a4995b0
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63511384"
---
# <a name="tutorial-build-a-message-compose-outlook-add-in"></a>Tutorial: criar uma mensagem para compor o suplemento do Outlook

Este tutorial ensina como criar um suplemento que pode ser usado em mensagens no modo de redação do Outlook para inserir conteúdo no corpo de uma mensagem.

Neste tutorial, você vai:

> [!div class="checklist"]
>
> - Criar um projeto de um suplemento do Outlook
> - Definir botões de renderização na janela de mensagem de texto
> - Implementar uma experiência de primeira execução que coleta informações do usuário e busca os dados de um serviço externo
> - Implementar um botão sem interface do usuário que chame uma função
> - Implementar um painel de tarefas que insere o conteúdo no corpo de uma mensagem

## <a name="prerequisites"></a>Pré-requisitos

[!INCLUDE [Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- [Visual Studio Code (VS Code)](https://code.visualstudio.com/) ou seu editor de código preferido

- Outlook 2016 ou posterior no Windows (conectado a uma conta do Microsoft 365) ou Outlook na Web

- Uma conta [GitHub](https://www.github.com) 

## <a name="setup"></a>Configurar

O suplemento que você criará neste tutorial lerá [gists](https://gist.github.com) da conta do GitHub do usuário e adicionará a essência selecionada ao corpo de uma mensagem. Conclua as etapas a seguir para criar duas gists novas que você pode usar para testar o suplemento que você vai criar.

1. [Faça logon no GitHub](https://github.com/login).

1. [Crie uma nova gist.](https://gist.github.com)

    - No campo **descrição do gist...**, insira **a Markdown Olá Mundo**.

    - No campo **nome do arquivo como extensão...** campo, insira **test.md**.

    - Adicione a seguinte marcação para a caixa de texto de várias linhas.

        ```markdown
        # Hello World

        This is content converted from Markdown!

        Here's a JSON sample:

          ```json
          {
            "foo": "bar"
          }
          ```
        ```

    - Selecione o botão **criar gist público**.

1. [Criar outro novo gist](https://gist.github.com).

    - No campo **descrição do gist...**, insira **Olá Mundo**.

    - No campo **nome do arquivo como extensão...** campo, insira **test.html**.

    - Adicione a seguinte marcação para a caixa de texto de várias linhas.

        ```HTML
        <html>
          <head>
            <style>
            h1 {
              font-family: Calibri;
            }
            </style>
          </head>
          <body>
            <h1>Hello World!</h1>
            <p>This is a test</p>
          </body>
        </html>
        ```

    - Selecione o botão **criar gist público**.

## <a name="create-an-outlook-add-in-project"></a>Criar um projeto de um suplemento do Outlook

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Escolha o tipo de projeto** - `Office Add-in Task Pane project`

    - **Escolha o tipo de script** - `JavaScript`

    - **Qual será o nome do suplemento?** - `Git the gist`

    - **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** - `Outlook`

    ![Captura de tela apresentando os avisos e respostas do gerador Yeoman em uma interface de linha de comando.](../images/yeoman-prompts-2.png)

    Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Navegue até o diretório raiz do projeto.

    ```command&nbsp;line
    cd "Git the gist"
    ```

1. Este suplemento usará as seguintes bibliotecas.

    - Biblioteca [Showdown](https://github.com/showdownjs/showdown) para converter Markdown em HTML.
    - Biblioteca [URI.js](https://github.com/medialize/URI.js) para criar URLs relativos.
    - Biblioteca [jquery](https://jquery.com/) para simplificar as interações com o DOM.

     Para instalar essas ferramentas para o seu projeto, execute o seguinte comando no diretório raiz do projeto.

    ```command&nbsp;line
    npm install showdown urijs jquery --save
    ```

1. Abra o projeto no VS Code ou no seu editor de código preferido.

    [!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

### <a name="update-the-manifest"></a>Atualizar o manifesto

O manifesto controla como o suplemento é exibido no Outlook. Ele define a maneira como o suplemento aparece na lista de suplementos e os botões que aparecem na faixa de opções, além de definir as URLs para os arquivos HTML e JavaScript usados pelo suplemento.

#### <a name="specify-basic-information"></a>Especifique as informações básicas

Faça as seguintes atualizações no arquivo **manifest.xml** para especificar algumas informações básicas sobre o suplemento.

1. Localize o elemento **ProviderName** e substitua o valor padrão pelo nome da empresa.

    ```xml
    <ProviderName>Contoso</ProviderName>
    ```

1. Localize o elemento **Description**, substitua o valor padrão com uma descrição do suplemento e salve o arquivo.

    ```xml
    <Description DefaultValue="Allows users to access their GitHub gists."/>
    ```

#### <a name="test-the-generated-add-in"></a>Testar o suplemento gerado

Antes de prosseguir, vamos testar o suplemento básico que criou o gerador para confirmar que o projeto está configurado corretamente.

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

1. Execute o seguinte comando no diretório raiz do seu projeto. Ao executar esse comando, o servidor Web local será iniciado e seu complemento será sideload.

    ```command&nbsp;line
    npm start
    ```

1. No Outlook, abra uma mensagem existente e selecione o botão **Mostrar Painel de Tarefas**.

1. Quando solicitado com a caixa de diálogo **Parar na Carga do Modo de Exibição da Web**, selecione **OK**.

    [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

    Se tudo tiver sido configurado corretamente, o painel de tarefas será aberto e exibirá a página de boas-vindas do suplemento.

    ![Captura de tela do botão "Mostrar Painel de Tarefas" e do Git do painel de tarefas de gist adicionado pelo exemplo.](../images/button-and-pane.png)

## <a name="define-buttons"></a>Definir botões

Agora que você verificou que o complemento básico funciona, você pode personalizá-lo para adicionar mais funcionalidades. Por padrão, o manifesto define apenas os botões para a janela de mensagem de leitura. Vamos atualizar o manifesto para remover os botões na janela de mensagem de leitura e definir dois novos botões para a janela de mensagem de texto:

- **Inserir gist**: um botão que abre um painel de tarefas

- **Inserir gist padrão**: um botão que invoca uma função

### <a name="remove-the-messagereadcommandsurface-extension-point"></a>Remover o ponto de extensão MessageReadCommandSurface

Abra o arquivo **manifest.xml** e localize o elemento **ExtensionPoint** com o tipo **MessageReadCommandSurface**. Exclua esse elemento **ExtensionPoint** (incluindo a marca de fechamento) para remover os botões da janela de mensagem de leitura.

### <a name="add-the-messagecomposecommandsurface-extension-point"></a>Adicionar o ponto de extensão MessageComposeCommandSurface

Encontre a seguinte linha no manifesto: `</DesktopFormFactor>`. Imediatamente antes dessa linha, insira a marcação XML a seguir. Observe o seguinte sobre esta marcação.

- O **ExtensionPoint** com `xsi:type="MessageComposeCommandSurface"` indica que você está definindo botões para adicionar à janela Redigir mensagem.

- Ao usar um elemento **OfficeTab** com `id="TabDefault"`, você indica que quer adicionar os botões à guia padrão da faixa de opções.

- O elemento **Group** define o agrupamento dos novos botões, com um rótulo definido pelo recurso **groupLabel**.

- O primeiro elemento **Control** contém um elemento **Action** com `xsi:type="ShowTaskPane"`, portanto, esse botão abre um painel de tarefas.

- O segundo elemento **Control** contém um elemento **Action** com `xsi:type="ExecuteFunction"`, o que indica que esse botão invoca uma função JavaScript contida no arquivo de função.

```xml
<!-- Message Compose -->
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgComposeCmdGroup">
      <Label resid="GroupLabel"/>
      <Control xsi:type="Button" id="msgComposeInsertGist">
        <Label resid="TaskpaneButton.Label"/>
        <Supertip>
          <Title resid="TaskpaneButton.Title"/>
          <Description resid="TaskpaneButton.Tooltip"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
          <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="Taskpane.Url"/>
        </Action>
      </Control>
      <Control xsi:type="Button" id="msgComposeInsertDefaultGist">
        <Label resid="FunctionButton.Label"/>
        <Supertip>
          <Title resid="FunctionButton.Title"/>
          <Description resid="FunctionButton.Tooltip"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
          <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
          <FunctionName>insertDefaultGist</FunctionName>
        </Action>
      </Control>
    </Group>
  </OfficeTab>
</ExtensionPoint>
```

### <a name="update-resources-in-the-manifest"></a>Atualização de recursos no manifesto

O código anterior faz referência a rótulos, dicas de ferramentas e URLs que você precisa definir antes que o manifesto seja válido. Você especificará estas informações na seção **Resources** do manifesto.

1. Localize o elemento **Resources** no arquivo do manifesto e exclua o elemento inteiro (incluindo sua marca de fechamento).

1. No mesmo local, adicione a seguinte marcação para substituir o elemento **Resources** que você acabou de remover.

    ```xml
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Git the gist"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Insert gist"/>
        <bt:String id="TaskpaneButton.Title" DefaultValue="Insert gist"/>
        <bt:String id="FunctionButton.Label" DefaultValue="Insert default gist"/>
        <bt:String id="FunctionButton.Title" DefaultValue="Insert default gist"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Displays a list of your gists and allows you to insert their contents into the current message."/>
        <bt:String id="FunctionButton.Tooltip" DefaultValue="Inserts the content of the gist you mark as default into the current message."/>
      </bt:LongStrings>
    </Resources>
    ```

1. Salve suas alterações no manifesto.

### <a name="reinstall-the-add-in"></a>Reinstalar o suplemento

Você deve reinstalar o complemento para que as alterações de manifesto entrem em vigor.

1. Se o servidor Web estiver em execução, feche a janela de comando do nó.

1. Execute o comando a seguir para iniciar o servidor Web local e realizar o sideload automático do suplemento.

    ```command&nbsp;line
    npm start
    ```

Depois de reinstalar o suplemento, você pode verificar se ele foi instalado com êxito verificando os comandos **Inserir gist** e **Inserir gist padrão** na janela de composição de mensagem. Observe que nada acontece quando você escolhe um destes itens, porque você ainda não terminou de criar este suplemento.

- Se você estiver executando este suplemento no Outlook 2016 ou posterior no Windows, deverá ver dois novos botões na faixa de opções da janela de composição da mensagem: **Inserir gist** e **Inserir gist padrão**.

    ![Captura de tela do menu excedente da faixa de opções do Outlook no Windows com os botões do suplemento em destaque.](../images/add-in-buttons-in-windows.png)

- Se você estiver usando este suplemento no Outlook na Web, você verá um botão na parte inferior da janela de composição de mensagem. Selecione esse botão para ver as opções **Insert Gist** e **Insert Default Gist**.

    ![Captura de tela do formulário de composição de mensagem no Outlook na Web com o botão suplemento e o menu pop-up em destaque.](../images/add-in-buttons-in-owa.png)

## <a name="implement-a-first-run-experience"></a>Implementando uma experiência de primeira execução

Este suplemento precisa ser capaz de ler gists da conta do GitHub do usuário e identificar qual deles o usuário escolheu como a essência padrão. Para obter esses objetivos, o suplemento deverá solicitar ao usuário para fornecer o nome de usuário do GitHub e escolher uma essência padrão do seu conjunto de gists existentes. Conclua as etapas nesta seção para implementar uma experiência de primeira execução que será exibida uma caixa de diálogo para obter essas informações do usuário.

### <a name="collect-data-from-the-user"></a>Coletar dados do usuário

Para começar, vamos criar o UI para a caixa de diálogo. Dentro da pasta **./src**, crie uma nova subpasta chamada **configurações**. Na pasta **./src/settings**, crie um arquivo chamado **dialog.html** e adicione a marcação a seguir para definir um formulário básico com uma entrada de texto de um nome de usuário do GitHub e uma lista vazia de gists que será preenchida via JavaScript.

```html
<!DOCTYPE html>
<html>

<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
  <title>Settings</title>

  <!-- Office JavaScript API -->
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

<!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui. -->
  <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

  <!-- Template styles -->
  <link href="dialog.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-l">
  <main>
    <section class="ms-font-m ms-fontColor-neutralPrimary">
      <div class="not-configured-warning ms-MessageBar ms-MessageBar--warning">
        <div class="ms-MessageBar-content">
          <div class="ms-MessageBar-icon">
            <i class="ms-Icon ms-Icon--Info"></i>
          </div>
          <div class="ms-MessageBar-text">
            Oops! It looks like you haven't configured <strong>Git the gist</strong> yet.
            <br/>
            Please configure your GitHub username and select a default gist, then try that action again!
          </div>
        </div>
      </div>
      <div class="ms-font-xxl">Settings</div>
      <div class="ms-Grid">
        <div class="ms-Grid-row">
          <div class="ms-TextField">
            <label class="ms-Label">GitHub Username</label>
            <input class="ms-TextField-field" id="github-user" type="text" value="" placeholder="Please enter your GitHub username">
          </div>
        </div>
        <div class="error-display ms-Grid-row">
          <div class="ms-font-l ms-fontWeight-semibold">An error occurred:</div>
          <pre><code id="error-text"></code></pre>
        </div>
        <div class="gist-list-container ms-Grid-row">
          <div class="list-title ms-font-xl ms-fontWeight-regular">Choose Default Gist</div>
          <form>
            <div id="gist-list">
            </div>
          </form>
        </div>
      </div>
      <div class="ms-Dialog-actions">
        <div class="ms-Dialog-actionsRight">
          <button class="ms-Dialog-action ms-Button ms-Button--primary" id="settings-done" disabled>
            <span class="ms-Button-label">Done</span>
          </button>
        </div>
      </div>
    </section>
  </main>
  <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
  <script type="text/javascript" src="../helpers/gist-api.js"></script>
  <script type="text/javascript" src="dialog.js"></script>
</body>

</html>
```

Você deve ter notado que o arquivo HTML faz referência a um arquivo JavaScript, **gist-api.js**, que ainda não existe. Esse arquivo será criado na seção abaixo [Buscar dados do GitHub](#fetch-data-from-github).

Em seguida, crie um arquivo na pasta **./src/settings** chamado **dialog.css** e adicione o seguinte código para especificar os estilos que são usados pelo **dialog.html**.

```CSS
section {
  margin: 10px 20px;
}

.not-configured-warning {
  display: none;
}

.error-display {
  display: none;
}

.gist-list-container {
  margin: 10px -8px;
  display: none;
}

.list-title {
  border-bottom: 1px solid #a6a6a6;
  padding-bottom: 5px;
}

ul {
  margin-top: 10px;
}

.ms-ListItem-secondaryText,
.ms-ListItem-tertiaryText {
  padding-left: 15px;
}
```

Agora que você definiu a IU da caixa de diálogo, você pode escrever código que realmente faz alguma coisa. Crie um arquivo na pasta **./src/settings** chamado **dialog.js** e adicione o seguinte código. Observe que esse código usa jQuery para registrar eventos e usa a função **messageParent** para enviar as opções do usuário de volta ao chamador.

```js
(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      if (window.location.search) {
        // Check if warning should be displayed.
        var warn = getParameterByName('warn');
        if (warn) {
          $('.not-configured-warning').show();
        } else {
          // See if the config values were passed.
          // If so, pre-populate the values.
          var user = getParameterByName('gitHubUserName');
          var gistId = getParameterByName('defaultGistId');

          $('#github-user').val(user);
          loadGists(user, function(success){
            if (success) {
              $('.ms-ListItem').removeClass('is-selected');
              $('input').filter(function() {
                return this.value === gistId;
              }).addClass('is-selected').attr('checked', 'checked');
              $('#settings-done').removeAttr('disabled');
            }
          });
        }
      }

      // When the GitHub username changes,
      // try to load gists.
      $('#github-user').on('change', function(){
        $('#gist-list').empty();
        var ghUser = $('#github-user').val();
        if (ghUser.length > 0) {
          loadGists(ghUser);
        }
      });

      // When the Done button is selected, send the
      // values back to the caller as a serialized
      // object.
      $('#settings-done').on('click', function() {
        var settings = {};

        settings.gitHubUserName = $('#github-user').val();

        var selectedGist = $('.ms-ListItem.is-selected');
        if (selectedGist) {
          settings.defaultGistId = selectedGist.val();

          sendMessage(JSON.stringify(settings));
        }
      });
    });
  };

  // Load gists for the user using the GitHub API
  // and build the list.
  function loadGists(user, callback) {
    getUserGists(user, function(gists, error){
      if (error) {
        $('.gist-list-container').hide();
        $('#error-text').text(JSON.stringify(error, null, 2));
        $('.error-display').show();
        if (callback) callback(false);
      } else {
        $('.error-display').hide();
        buildGistList($('#gist-list'), gists, onGistSelected);
        $('.gist-list-container').show();
        if (callback) callback(true);
      }
    });
  }

  function onGistSelected() {
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
    $('.not-configured-warning').hide();
    $('#settings-done').removeAttr('disabled');
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }
})();
```

#### <a name="update-webpack-config-settings"></a>Atualizar as configurações webpack config

Por fim, abra o arquivo **webpack.config.js** encontrado no diretório raiz do projeto e conclua as etapas a seguir.

1. Localize o objeto `entry` dentro do objeto `config` e adicione uma nova entrada para `dialog`.

    ```js
    dialog: "./src/settings/dialog.js",
    ```

    Após fazer isso, o novo objeto `entry` ficará assim:

    ```js
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      dialog: "./src/settings/dialog.js",
    },
    ```

1. Localize a `plugins` matriz dentro do `config` objeto. Na matriz `patterns` do objeto `new CopyWebpackPlugin`, adicione novas entradas para **taskpane.css** e **dialog.css**.

    ```js
    {
      from: "./src/taskpane/taskpane.css",
      to: "taskpane.css",
    },
    {
      from: "./src/settings/dialog.css",
      to: "dialog.css",
    },
    ```

    Após fazer isso, o `new CopyWebpackPlugin` objeto terá a seguinte aparência:

    ```js
    new CopyWebpackPlugin({
      patterns: [
      {
        from: "./src/taskpane/taskpane.css",
        to: "taskpane.css",
      },
      {
        from: "./src/settings/dialog.css",
        to: "dialog.css",
      },
      {
        from: "assets/*",
        to: "assets/[name][ext][query]",
      },
      {
        from: "manifest*.xml",
        to: "[name]." + buildType + "[ext]",
        transform(content) {
          if (dev) {
            return content;
          } else {
            return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
          }
        },
      },
    ]}),
    ```

1. Na mesma matriz `plugins` do objeto `config`, adicione esse novo objeto ao final da matriz.

    ```js
    new HtmlWebpackPlugin({
      filename: "dialog.html",
      template: "./src/settings/dialog.html",
      chunks: ["polyfill", "dialog"]
    })
    ```

    Após fazer isso, a nova matriz `plugins` ficará assim:

    ```js
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "./src/taskpane/taskpane.css",
            to: "taskpane.css",
          },
          {
            from: "./src/settings/dialog.css",
            to: "dialog.css",
          },
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]." + buildType + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/settings/dialog.html",
        chunks: ["polyfill", "dialog"]
      })
    ],
    ```

### <a name="fetch-data-from-github"></a>Buscar dados do GitHub

O arquivo **dialog.js** que você acabou de criar especifica que o suplemento deve carregar gists quando o evento de **alteração** for acionado para o campo de nome de usuário do GitHub. Para recuperar gists do usuário do GitHub, você usará o [GitHub Gists API](https://developer.github.com/v3/gists/).

Dentro da pasta **./src**, crie uma nova subpasta nomeada **helpers**. Na pasta **./src/helpers**, crie um arquivo nomeado **gist-api.js** e adicione o seguinte código para recuperar os gists do usuário do GitHub e criar a lista de gists.

```js
function getUserGists(user, callback) {
  var requestUrl = 'https://api.github.com/users/' + user + '/gists';

  $.ajax({
    url: requestUrl,
    dataType: 'json'
  }).done(function(gists){
    callback(gists);
  }).fail(function(error){
    callback(null, error);
  });
}

function buildGistList(parent, gists, clickFunc) {
  gists.forEach(function(gist) {

    var listItem = $('<div/>')
      .appendTo(parent);

    var radioItem = $('<input>')
      .addClass('ms-ListItem')
      .addClass('is-selectable')
      .attr('type', 'radio')
      .attr('name', 'gists')
      .attr('tabindex', 0)
      .val(gist.id)
      .appendTo(listItem);

    var desc = $('<span/>')
      .addClass('ms-ListItem-primaryText')
      .text(gist.description)
      .appendTo(listItem);

    var desc = $('<span/>')
      .addClass('ms-ListItem-secondaryText')
      .text(' - ' + buildFileList(gist.files))
      .appendTo(listItem);

    var updated = new Date(gist.updated_at);

    var desc = $('<span/>')
      .addClass('ms-ListItem-tertiaryText')
      .text(' - Last updated ' + updated.toLocaleString())
      .appendTo(listItem);

    listItem.on('click', clickFunc);
  });  
}

function buildFileList(files) {

  var fileList = '';

  for (var file in files) {
    if (files.hasOwnProperty(file)) {
      if (fileList.length > 0) {
        fileList = fileList + ', ';
      }

      fileList = fileList + files[file].filename + ' (' + files[file].language + ')';
    }
  }

  return fileList;
}
```

Execute o seguinte comando para recriar o projeto.

```command&nbsp;line
npm run build
```

## <a name="implement-a-ui-less-button"></a>Implementar um botão sem interface do usuário

O botão **Inserir gist padrão** do suplemento é um botão sem interface do usuário que invocará uma função JavaScript, em vez de abrir um painel de tarefas como muitos dos botões suplementares. Quando o usuário seleciona o botão **Inserir gist padrão**, a função JavaScript correspondente verificará se o suplemento foi configurado.

- Se o suplemento já tiver sido configurado, a função carregará o conteúdo do gist selecionado pelo usuário como padrão e o inserirá no corpo da mensagem.

- Se o suplemento ainda não tiver sido configurado, a caixa de diálogo de configurações solicitará ao usuário que forneça as informações necessárias.

### <a name="update-the-function-file-html"></a>Atualizar o arquivo de função (HTML)

Uma função que é invocada por um botão sem interface do usuário deve ser definida no arquivo que é especificado pelo elemento **FunctionFile** no manifesto do fator de forma correspondente. O manifesto deste suplemento especifica `https://localhost:3000/commands.html` como o arquivo de função.

Abra o arquivo **./src/commands/commands.html** e substitua todo o conteúdo pela marcação a seguir.

```html
<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

    <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
    <script type="text/javascript" src="../../node_modules/showdown/dist/showdown.min.js"></script>
    <script type="text/javascript" src="../../node_modules/urijs/src/URI.min.js"></script>
    <script type="text/javascript" src="../helpers/addin-config.js"></script>
    <script type="text/javascript" src="../helpers/gist-api.js"></script>
</head>

<body>
  <!-- NOTE: The body is empty on purpose. Since functions in commands.js are
       invoked via a button, there is no UI to render. -->
</body>

</html>
```

Você deve ter notado que o arquivo HTML faz referência a um arquivo JavaScript, **addin-config.js**, que ainda não existe. Esse arquivo será criado na seção [Criar um arquivo para gerenciar as definições de configuração](#create-a-file-to-manage-configuration-settings) posteriormente neste tutorial.

### <a name="update-the-function-file-javascript"></a>Atualizar o arquivo de função (JavaScript)

Abra o arquivo **./src/commands/commands.js** e substitua todo o conteúdo pelo código a seguir. Observe que, se a função **insertDefaultGist** determinar que o suplemento ainda não foi configurado, ele adicionará o parâmetro `?warn=1` à URL da caixa de diálogo. Isso faz com que a caixa de diálogo de configurações renderize a barra de mensagens definida no **./src/settings/dialog.html**, para informar ao usuário por que ele vê a caixa de diálogo.

```js
var config;
var btnEvent;

// The initialize function must be run each time a new page is loaded.
Office.initialize = function (reason) {
};

// Add any UI-less function here.
function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('github-error', {
    type: 'errorMessage',
    message: error
  }, function(result){
  });
}

var settingsDialog;

function insertDefaultGist(event) {

  config = getConfig();

  // Check if the add-in has been configured.
  if (config && config.defaultGistId) {
    // Get the default gist content and insert.
    try {
      getGist(config.defaultGistId, function(gist, error) {
        if (gist) {
          buildBodyContent(gist, function (content, error) {
            if (content) {
              Office.context.mailbox.item.body.setSelectedDataAsync(content,
                {coercionType: Office.CoercionType.Html}, function(result) {
                  event.completed();
              });
            } else {
              showError(error);
              event.completed();
            }
          });
        } else {
          showError(error);
          event.completed();
        }
      });
    } catch (err) {
      showError(err);
      event.completed();
    }

  } else {
    // Save the event object so we can finish up later.
    btnEvent = event;
    // Not configured yet, display settings dialog with
    // warn=1 to display warning.
    var url = new URI('dialog.html?warn=1').absoluteTo(window.location).toString();
    var dialogOptions = { width: 20, height: 40, displayInIframe: true };

    Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
      settingsDialog = result.value;
      settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
      settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
    });
  }
}

function receiveMessage(message) {
  config = JSON.parse(message.message);
  setConfig(config, function(result) {
    settingsDialog.close();
    settingsDialog = null;
    btnEvent.completed();
    btnEvent = null;
  });
}

function dialogClosed(message) {
  settingsDialog = null;
  btnEvent.completed();
  btnEvent = null;
}

function getGlobal() {
  return (typeof self !== "undefined") ? self :
    (typeof window !== "undefined") ? window :
    (typeof global !== "undefined") ? global :
    undefined;
}

var g = getGlobal();

// The add-in command functions need to be available in global scope.
g.insertDefaultGist = insertDefaultGist;
```

### <a name="create-a-file-to-manage-configuration-settings"></a>Criar um arquivo para gerenciar configurações

O arquivo de função HTML faz referência a um arquivo chamado **suplemento config.js**, que ainda não existe. Na pasta **./src/helpers**, crie um arquivo chamado **addin-config.js** e adicione o código a seguir. O código usa o [Objeto RoamingSettings](/javascript/api/outlook/office.roamingsettings) para definir valores de configuração.

```js
function getConfig() {
  var config = {};

  config.gitHubUserName = Office.context.roamingSettings.get('gitHubUserName');
  config.defaultGistId = Office.context.roamingSettings.get('defaultGistId');

  return config;
}

function setConfig(config, callback) {
  Office.context.roamingSettings.set('gitHubUserName', config.gitHubUserName);
  Office.context.roamingSettings.set('defaultGistId', config.defaultGistId);

  Office.context.roamingSettings.saveAsync(callback);
}
```

### <a name="create-new-functions-to-process-gists"></a>Criar novas funções para processar gists

Em seguida, abra o arquivo **./src/helpers/gist-api.js** e adicione as seguintes funções. Observe o seguinte:

- Se a essência contiver HTML, o suplemento inserirá o HTML como está no corpo da mensagem.

- Se o gist contiver redução, ele usará a biblioteca[Showdown](https://github.com/showdownjs/showdown) para converter a redução em HTML e, em seguida, inserirá o HTML resultante no corpo da mensagem.

- Se a essência contiver algo diferente de HTML ou redução, o suplemento a inserirá no corpo da mensagem como um trecho de código.

```js
function getGist(gistId, callback) {
  var requestUrl = 'https://api.github.com/gists/' + gistId;

  $.ajax({
    url: requestUrl,
    dataType: 'json'
  }).done(function(gist){
    callback(gist);
  }).fail(function(error){
    callback(null, error);
  });
}

function buildBodyContent(gist, callback) {
  // Find the first non-truncated file in the gist
  // and use it.
  for (var filename in gist.files) {
    if (gist.files.hasOwnProperty(filename)) {
      var file = gist.files[filename];
      if (!file.truncated) {
        // We have a winner.
        switch (file.language) {
          case 'HTML':
            // Insert as is.
            callback(file.content);
            break;
          case 'Markdown':
            // Convert Markdown to HTML.
            var converter = new showdown.Converter();
            var html = converter.makeHtml(file.content);
            callback(html);
            break;
          default:
            // Insert contents as a <code> block.
            var codeBlock = '<pre><code>';
            codeBlock = codeBlock + file.content;
            codeBlock = codeBlock + '</code></pre>';
            callback(codeBlock);
        }
        return;
      }
    }
  }
  callback(null, 'No suitable file found in the gist');
}
```

### <a name="test-the-insert-default-gist-button"></a>Testar o botão Inserir gist padrão

Salvar todas as suas alterações e executar `npm start` do prompt de comando, se o servidor não estiver sendo executado. Conclua as seguintes etapas para testar o botão **Inserir Gist Padrão**.

1. Abra o Outlook e redija uma nova mensagem.

1. Na janela de mensagem de texto, selecione o botão **Inserir Gist Padrão**. Você verá uma caixa de diálogo na qual é possível configurar o suplemento, começando com o prompt para definir seu nome de usuário do GitHub.

    ![Captura de tela do prompt de diálogo para configurar o suplemento.](../images/addin-prompt-configure.png)

1. Na caixa de diálogo de configurações, insira seu nome de usuário do GitHub e, em seguida, **Tab** ou clique em qualquer lugar na caixa de diálogo para invocar o evento **change**, que deve carregar sua lista de gists públicos. Selecione um gist para ser o padrão e selecione **Concluído**.

    ![Captura de tela de caixa de diálogo de configurações do suplemento.](../images/addin-settings.png)

1. Selecione o botão **Inserir gist padrão** novamente. Desta vez, você deve ver o conteúdo do gist inserido no corpo do email.

   > [!NOTE]
   > Outlook no Windows: Para selecionar as configurações mais recentes, talvez seja necessário fechar e reabrir a janela de composição de mensagens.

## <a name="implement-a-task-pane"></a>Implementar um painel de tarefas

O botão **Inserir gist** deste suplemento abrirá o painel de tarefas e exibirá os gists do usuário. Em seguida, o usuário pode selecionar uma das gists para inserir no corpo da mensagem. Se o usuário ainda não tiver configurado o suplemento, ele será solicitado a fazê-lo.

### <a name="specify-the-html-for-the-task-pane"></a>Especificar o arquivo HTML para o painel de tarefas

No projeto que você criou, o painel de tarefas HTML é especificado no arquivo **./src/taskpane/taskpane.html**. Abra o arquivo e substitua todo o conteúdo pela seguinte marcação.

```html
<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Contoso Task Pane Add-in</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

   <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui. -->
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

    <!-- Template styles -->
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-l ms-landing-page">
  <main class="ms-landing-page__main">
    <section class="ms-landing-page__content ms-font-m ms-fontColor-neutralPrimary">
      <div id="not-configured" style="display: none;">
        <div class="centered ms-font-xxl ms-u-textAlignCenter">Welcome!</div>
        <div class="ms-font-xl" id="settings-prompt">Please choose the <strong>Settings</strong> icon at the bottom of this window to configure this add-in.</div>
      </div>
      <div id="gist-list-container" style="display: none;">
        <form>
          <div id="gist-list">
          </div>
        </form>
      </div>
      <div id="error-display" style="display: none;" class="ms-u-borderBase ms-fontColor-error ms-font-m ms-bgColor-error ms-borderColor-error">
      </div>
    </section>
    <button class="ms-Button ms-Button--primary" id="insert-button" tabindex=0 disabled>
      <span class="ms-Button-label">Insert</span>
    </button>
  </main>
  <footer class="ms-landing-page__footer ms-bgColor-themePrimary">
    <div class="ms-landing-page__footer--left">
      <img src="../../assets/logo-filled.png" />
      <h1 class="ms-font-xl ms-fontWeight-semilight ms-fontColor-white">Git the gist</h1>
    </div>
    <div id="settings-icon" class="ms-landing-page__footer--right" aria-label="Settings" tabindex=0>
      <i class="ms-Icon enlarge ms-Icon--Settings ms-fontColor-white"></i>
    </div>
  </footer>
  <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
  <script type="text/javascript" src="../../node_modules/showdown/dist/showdown.min.js"></script>
  <script type="text/javascript" src="../../node_modules/urijs/src/URI.min.js"></script>
  <script type="text/javascript" src="../helpers/addin-config.js"></script>
  <script type="text/javascript" src="../helpers/gist-api.js"></script>
  <script type="text/javascript" src="taskpane.js"></script>
</body>

</html>
```

### <a name="specify-the-css-for-the-task-pane"></a>Especificar o CSS para o painel de tarefas

No projeto que você criou, o painel de tarefas CSS é especificado no arquivo **./src/taskpane/taskpane.css**. Abra o arquivo e substitua todo o conteúdo pelo seguinte código.

```css
/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. */
html, body {
  width: 100%;
  height: 100%;
  margin: 0;
  padding: 0;
  overflow: auto; }

body {
  position: relative;
  font-size: 16px; }

main {
  height: 100%;
  overflow-y: auto; }

footer {
  width: 100%;
  position: relative;
  bottom: 0;
  margin-top: 10px;}

p, h1, h2, h3, h4, h5, h6 {
  margin: 0;
  padding: 0; }

ul {
  padding: 0; }

#settings-prompt {
  margin: 10px 0;
}

#error-display {
  padding: 10px;
}

#insert-button {
  margin: 0 10px;
}

.clearfix {
  display: block;
  clear: both;
  height: 0; }

.pointerCursor {
  cursor: pointer; }

.invisible {
  visibility: hidden; }

.undisplayed {
  display: none; }

.ms-Icon.enlarge {
  position: relative;
  font-size: 20px;
  top: 4px; }

.ms-ListItem-secondaryText,
.ms-ListItem-tertiaryText {
  padding-left: 15px;
}

.ms-landing-page {
  display: -webkit-flex;
  display: flex;
  -webkit-flex-direction: column;
          flex-direction: column;
  -webkit-flex-wrap: nowrap;
          flex-wrap: nowrap;
  height: 100%; }
  .ms-landing-page__main {
    display: -webkit-flex;
    display: flex;
    -webkit-flex-direction: column;
            flex-direction: column;
    -webkit-flex-wrap: nowrap;
            flex-wrap: nowrap;
    -webkit-flex: 1 1 0;
            flex: 1 1 0;
    height: 100%; }

  .ms-landing-page__content {
    display: -webkit-flex;
    display: flex;
    -webkit-flex-direction: column;
            flex-direction: column;
    -webkit-flex-wrap: nowrap;
            flex-wrap: nowrap;
    height: 100%;
    -webkit-flex: 1 1 0;
            flex: 1 1 0;
    padding: 20px; }
    .ms-landing-page__content h2 {
      margin-bottom: 20px; }
  .ms-landing-page__footer {
    display: -webkit-inline-flex;
    display: inline-flex;
    -webkit-justify-content: center;
            justify-content: center;
    -webkit-align-items: center;
            align-items: center; }
    .ms-landing-page__footer--left {
      transition: background ease 0.1s, color ease 0.1s;
      display: -webkit-inline-flex;
      display: inline-flex;
      -webkit-justify-content: flex-start;
              justify-content: flex-start;
      -webkit-align-items: center;
              align-items: center;
      -webkit-flex: 1 0 0px;
              flex: 1 0 0px;
      padding: 20px; }
      .ms-landing-page__footer--left:active {
        cursor: default; }
      .ms-landing-page__footer--left--disabled {
        opacity: 0.6;
        pointer-events: none;
        cursor: not-allowed; }
        .ms-landing-page__footer--left--disabled:active, .ms-landing-page__footer--left--disabled:hover {
          background: transparent; }
      .ms-landing-page__footer--left img {
        width: 40px;
        height: 40px; }
      .ms-landing-page__footer--left h1 {
        -webkit-flex: 1 0 0px;
                flex: 1 0 0px;
        margin-left: 15px;
        text-align: left;
        width: auto;
        max-width: auto;
        overflow: hidden;
        white-space: nowrap;
        text-overflow: ellipsis; }
    .ms-landing-page__footer--right {
      transition: background ease 0.1s, color ease 0.1s;
      padding: 29px 20px; }
      .ms-landing-page__footer--right:active, .ms-landing-page__footer--right:hover {
        background: #005ca4;
        cursor: pointer; }
      .ms-landing-page__footer--right:active {
        background: #005ca4; }
      .ms-landing-page__footer--right--disabled {
        opacity: 0.6;
        pointer-events: none;
        cursor: not-allowed; }
        .ms-landing-page__footer--right--disabled:active, .ms-landing-page__footer--right--disabled:hover {
          background: transparent; }
```

### <a name="specify-the-javascript-for-the-task-pane"></a>Especificar o JavaScript para o painel de tarefas

No projeto que você criou, o painel de tarefas JavaScript é especificado no arquivo **./src/taskpane/taskpane.js**. Abra o arquivo e substitua todo o conteúdo pelo seguinte código.

```js
(function(){
  'use strict';

  var config;
  var settingsDialog;

  Office.initialize = function(reason){

    jQuery(document).ready(function(){

      config = getConfig();

      // Check if add-in is configured.
      if (config && config.gitHubUserName) {
        // If configured, load the gist list.
        loadGists(config.gitHubUserName);
      } else {
        // Not configured yet.
        $('#not-configured').show();
      }

      // When insert button is selected, build the content
      // and insert into the body.
      $('#insert-button').on('click', function(){
        var gistId = $('.ms-ListItem.is-selected').val();
        getGist(gistId, function(gist, error) {
          if (gist) {
            buildBodyContent(gist, function (content, error) {
              if (content) {
                Office.context.mailbox.item.body.setSelectedDataAsync(content,
                  {coercionType: Office.CoercionType.Html}, function(result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                      showError('Could not insert gist: ' + result.error.message);
                    }
                });
              } else {
                showError('Could not create insertable content: ' + error);
              }
            });
          } else {
            showError('Could not retrieve gist: ' + error);
          }
        });
      });

      // When the settings icon is selected, open the settings dialog.
      $('#settings-icon').on('click', function(){
        // Display settings dialog.
        var url = new URI('dialog.html').absoluteTo(window.location).toString();
        if (config) {
          // If the add-in has already been configured, pass the existing values
          // to the dialog.
          url = url + '?gitHubUserName=' + config.gitHubUserName + '&defaultGistId=' + config.defaultGistId;
        }

        var dialogOptions = { width: 20, height: 40, displayInIframe: true };

        Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
          settingsDialog = result.value;
          settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
          settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        });
      })
    });
  };

  function loadGists(user) {
    $('#error-display').hide();
    $('#not-configured').hide();
    $('#gist-list-container').show();

    getUserGists(user, function(gists, error) {
      if (error) {

      } else {
        $('#gist-list').empty();
        buildGistList($('#gist-list'), gists, onGistSelected);
      }
    });
  }

  function onGistSelected() {
    $('#insert-button').removeAttr('disabled');
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
  }

  function showError(error) {
    $('#not-configured').hide();
    $('#gist-list-container').hide();
    $('#error-display').text(error);
    $('#error-display').show();
  }

  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function(result) {
      settingsDialog.close();
      settingsDialog = null;
      loadGists(config.gitHubUserName);
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();
```

### <a name="test-the-insert-gist-button"></a>Testar o botão Inserir gist

Salvar todas as suas alterações e executar `npm start` do prompt de comando, se o servidor não estiver sendo executado. Conclua as seguintes etapas para testar o botão **Inserir gist** botão.

1. Abra o Outlook e redija uma nova mensagem.

1. Na janela de mensagem de texto, selecione o botão **Inserir gist**. Você verá um painel de tarefas aberto à direita do formulário de texto.

1. No painel de tarefas, selecione a gist **Olá mundo Html** e selecione **Inserir** para inserir esse gist no corpo da mensagem.

![Captura de tela do painel de tarefas do suplemento e o conteúdo gist selecionado exibido no corpo da mensagem.](../images/addin-taskpane.png)

## <a name="next-steps"></a>Próximas etapas

Neste tutorial, você criou um suplemento do Outlook que pode ser usado no modo de composição de mensagens para inserir conteúdo no corpo de uma mensagem. Para saber mais sobre o desenvolvimento de suplementos do Outlook, continue no seguinte artigo.

> [!div class="nextstepaction"]
> [APIs de suplemento do Outlook](../outlook/apis.md)
