---
ms.date: 07/27/2021
title: Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado
ms.prod: non-product-specific
description: Configure seu suplemento do Office para usar um tempo de execução de JavaScript compartilhado para oferecer suporte à faixa de opções adicional, painel de tarefas e recursos de funções personalizadas.
localization_priority: Priority
ms.openlocfilehash: 9e24545bac2b2aaad58c2441ed0a5741c78c053d
ms.sourcegitcommit: 3cc8f6adee0c7c68c61a42da0d97ed5ea61be0ac
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53661136"
---
# <a name="configure-your-office-add-in-to-use-a-shared-javascript-runtime"></a>Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

É possível configurar o Suplemento do Office para executar todo o seu código em um único tempo de execução JavaScript compartilhado (também conhecido como tempo de execução compartilhado). Isso permite uma melhor coordenação em seu suplemento e acesso ao DOM e CORS de todas as partes de seu suplemento. Ele também ativa recursos adicionais, como a execução de código quando o documento é aberto ou a ativação ou desativação de botões da faixa de opções. Para configurar seu suplemento para usar um tempo de execução JavaScript compartilhado, siga as instruções neste artigo.

## <a name="create-the-add-in-project"></a>Criar o projeto de suplemento

Se você estiver iniciando um novo projeto, siga estas etapas para usar o [ gerador Yeoman para suplementos do Office ](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento do Excel ou PowerPoint.

Faça um dos seguintes:

- Para gerar um suplemento do Excel com funções personalizadas, execute o comando `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`.

    ou

- Para gerar um suplemento do PowerPoint, execute o comando `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true`.

O gerador criará o projeto e instalará os componentes do Node de suporte.

> [!NOTE]
> Você também pode usar as etapas neste artigo para atualizar um projeto existente do Visual Studio para usar o runtime compartilhado. No entanto, talvez seja necessário atualizar os esquemas XML para o manifesto. Para obter mais informações, confira [Solucionar erros de desenvolvimento com Suplementos do Office](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

## <a name="configure-the-manifest"></a>Configurar o manifesto

Siga estas etapas para um projeto novo ou existente para configurá-lo para usar um tempo de execução compartilhado. Estas etapas pressupõem que você gerou seu projeto usando o [Gerador Yeoman para Suplementos do Office ](https://github.com/OfficeDev/generator-office).

1. Inicie o Visual Studio Code e abra o projeto de suplemento do Excel ou PowerPoint que você gerou.
1. Abra o arquivo **manifest.xml**.
1. Se você gerou um suplemento do Excel, atualize a seção de requisitos para usar o [tempo de execução compartilhado](../reference/requirement-sets/shared-runtime-requirement-sets.md) em vez do tempo de execução da função personalizada. O XML deve aparecer da seguinte maneira.

    ```xml
    <Hosts>
      <Host Name="Workbook"/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. Localize a `<VersionOverrides>`seção e adicione a seguinte`<Runtimes>` seção apenas dentro da `<Host ...>`marca. A vida útil deve ser **longa** para que o código do suplemento possa ser executado mesmo quando o painel de tarefas está fechado. O `resid`valor é **Taskpane.Url**, que faz referência ao local do arquivo **taskpane.html** especificado na ` <bt:Urls>`seção próxima à parte inferior do arquivo **manifest.xml**.

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
       <Runtimes>
         <Runtime resid="Taskpane.Url" lifetime="long" />
       </Runtimes>
       ...
   ```

1. Se você gerou um Suplemento do Excel com funções personalizadas, localize o elemento `<Page>`. Em seguida, altere o local de origem de **Functions.Page.Url** para **Taskpane.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. Localize a marca`<FunctionFile ...>` e altere o `resid` de **Commands.Url** para **Taskpane.Url**. Observe que, se você não tiver comandos de ação, não terá uma entrada **FunctionFile** e pode pular esta etapa.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Salve o arquivo **manifest.xml**.

## <a name="configure-the-webpackconfigjs-file"></a>Configurar o arquivo webpack.config.js

O **webpack.config.js** construirá vários carregadores de tempo de execução. É necessário modificá-lo para carregar apenas o tempo de execução JavaScript compartilhado por meio do arquivo **taskpane.html**.

1. Inicie o Visual Studio Code e abra o projeto de suplemento do Excel ou PowerPoint que você gerou.
1. Abra o arquivo **webpack.config.js**.
1. Se o arquivo **webpack.config.js** tiver o seguinte código de plug-in **functions.html**, remova-o.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. Se o seu arquivo **webpack.config.js** tiver o seguinte código de plug-in **functions.html**, remova-o.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. Se o seu projeto usou as **functions** ou os blocos de **commands**, adicione-os à lista de blocos conforme mostrado a seguir (o código a seguir é para se o seu projeto usou os dois blocos).

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Salvar suas alterações e reconstrua o projeto.

   ```command line
   npm run build
   ```

> [!NOTE]
> Se o seu projeto tiver um arquivo **functions.html** ou um arquivo **commands.html**, eles podem ser removidos. O **taskpane.html** carregará o código **functions.js** e **commands.js** no tempo de execução JavaScript compartilhado por meio das atualizações do webpack que você acabou de fazer.

## <a name="test-your-office-add-in-changes"></a>Teste as alterações do Suplemento do Office

É possível confirmar que está usando o tempo de execução de JavaScript compartilhado corretamente usando as instruções a seguir.

1. Abra o arquivo **manifest.xml**.
1. Localize a seção `<Control xsi:type="Button" id="TaskpaneButton">` e altere o seguinte `<Action ...>` XML.

    de:

    ```xml
    <Action xsi:type="ShowTaskpane">
      <TaskpaneId>ButtonId1</TaskpaneId>
      <SourceLocation resid="Taskpane.Url"/>
    </Action>
    ```

    para:

    ```xml
    <Action xsi:type="ExecuteFunction">
      <FunctionName>action</FunctionName>
    </Action>
    ```

1. Abra o arquivo **./src/commands/commands.js**.
1. Substitua a função **ação** pelo código abaixo. Isso atualizará a função para abrir e modificar o botão do painel de tarefas para incrementar um contador. Abrir e acessar o DOM do painel de tarefas a partir de um comando só funciona com o tempo de execução JavaScript compartilhado.

    ```javascript
    var _count=0;
    
    function action(event) {
      // Your code goes here.
      _count++;
      Office.addin.showAsTaskpane();
      document.getElementById("run").textContent="Go"+_count;
    
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
    ```

1. Salve suas alterações e execute o projeto.

   ```command line
   npm start
   ```

Cada vez que você selecionar o botão suplementos, ele mudará o texto do botão **executar** para **ir** e incrementará um contador após ele.

## <a name="runtime-lifetime"></a>Duração do tempo de execução

Ao adicionar o elemento `Runtime`, você também especifica uma vida útil com um valor de `long` ou `short`. Defina esse valor como `long` para aproveitar os recursos, como iniciar o suplemento quando o documento for aberto, continuar executando o código após o fechamento do painel de tarefas ou usar o CORS e o DOM nas funções personalizadas.

> [!NOTE]
> O valor padrão de tempo de vida é `short`, mas recomendamos usar o `long` em suplementos do Excel. Se você definir o tempo de execução como `short` neste exemplo, o suplemento do Excel será iniciado quando um dos botões da faixa de opções for pressionado, mas poderá ser encerrado depois que o manipulador da faixa de opções for concluído. Da mesma forma, o suplemento será iniciado quando o painel de tarefas for aberto, mas pode ser encerrado quando o painel de tarefas for fechado.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> Se seu suplemento inclui o elemento `Runtimes` no manifesto (necessário para um tempo de execução compartilhado), ele utiliza o Internet Explorer 11 independentemente da versão do Windows ou do Microsoft 365. Para mais informações, consulte [Runtimes](../reference/manifest/runtimes.md).

## <a name="about-the-shared-javascript-runtime"></a>Sobre o tempo de execução de JavaScript compartilhado

No Windows ou Mac, seu suplemento executará o código para botões da faixa de opções, funções personalizadas e o painel de tarefas em ambientes de tempo de execução JavaScript separados. Isso cria limitações, como não poder compartilhar facilmente dados globais e não poder acessar todas as funcionalidades do CORS a partir de uma função customizada.

No entanto, você pode configurar o Suplemento do Office para compartilhar código no mesmo tempo de execução JavaScript (também conhecido como tempo de execução compartilhado). Isso permite uma melhor coordenação entre o suplemento e o acesso ao DOM e CORS do painel de tarefas de todas as partes do suplemento.

Configurar um tempo de execução compartilhado permite os seguintes cenários.

- Seu Suplemento do Office pode usar recursos adicionais da IU:
  - [Adicionar atalhos de teclado Personalizados aos Suplementos do Office (pré-visualização)](../design/keyboard-shortcuts.md)
  - [Crie guias contextuais Personalizadas em Suplementos do Office (pré-visualização)](../design/contextual-tabs.md)
  - [Ativar e Desativar Comandos de Suplemento](../design/disable-add-in-commands.md)
  - [Execute o código em seu Suplemento do Office quando o documento for aberto](run-code-on-document-open.md)
  - [Mostre ou oculte o painel de tarefas de seu Suplemento do Office ](show-hide-add-in.md)
- Para suplementos do Excel:
  - As funções personalizadas terão suporte CORS completo.
  - Funções personalizadas podem chamar APIs Office.js para ler dados de documentos de planilhas.

Para Office no Windows, o tempo de execução compartilhado usa Microsoft Edge com WebView2 (baseado em Chromium) se as condições para usá-lo forem atendidas conforme explicado em [Navegadores usados por suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Caso contrário, ele usa o Internet Explorer 11. Além disso, todos os botões que seu suplemento exibir na faixa de opções serão executados no mesmo tempo de execução compartilhado. A imagem a seguir mostra como as funções personalizadas, a interface do usuário da faixa de opções e o código do painel de tarefas serão executados no mesmo tempo de execução do JavaScript.

![Diagrama de uma função personalizada, painel de tarefas e botões da faixa de opções em execução em um tempo de execução de navegador compartilhado no Excel.](../images/custom-functions-in-browser-runtime.png)

### <a name="debugging"></a>Depuração

Ao usar um tempo de execução compartilhado, não é possível usar o Código do Visual Studio para depurar funções personalizadas no Excel no Windows no momento. Em vez disso, você precisará usar as ferramentas de desenvolvedor. Para obter mais informações, consulte [Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

### <a name="multiple-task-panes"></a>Vários painéis de tarefas

Não projete seu suplemento para usar vários painéis de tarefas se você planeja usar um tempo de execução compartilhado. Um tempo de execução compartilhado tem suporte para o uso de apenas um único painel de tarefas. Observe que qualquer painel de tarefas sem um `<TaskpaneID>` é considerado um painel de tarefas diferente.

## <a name="give-us-feedback"></a>Envie-nos seus comentários

Adoraríamos ouvir seus comentários sobre esse recurso. Se você encontrar algum bug ou problema, ou tiver solicitações sobre esse recurso, informe-nos criando um problema do GitHub no [repositório office-js](https://github.com/OfficeDev/office-js).

## <a name="see-also"></a>Confira também

- [Chamar APIs do Excel a partir de uma função personalizada](../excel/call-excel-apis-from-custom-function.md)
- [Adicione atalhos de teclado personalizados aos suplementos do Office (pré-visualização)](../design/keyboard-shortcuts.md)
- [Crie guias contextuais Personalizadas em Suplementos do Office (pré-visualização)](../design/contextual-tabs.md)
- [Ativar e Desativar Comandos de Suplemento](../design/disable-add-in-commands.md)
- [Execute o código em seu Suplemento do Office quando o documento for aberto](run-code-on-document-open.md)
- [Mostre ou oculte o painel de tarefas de seu Suplemento do Office ](show-hide-add-in.md)
- [Tutorial: compartilhar dados e eventos entre as funções personalizadas do Excel e do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
