---
title: Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado
description: Configure seu suplemento do Office para usar um tempo de execução de JavaScript compartilhado para oferecer suporte à faixa de opções adicional, painel de tarefas e recursos de funções personalizadas.
ms.date: 07/18/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: 70906199f27a5b84a9dcd71b2f36dcd16ff79f73
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889475"
---
# <a name="configure-your-office-add-in-to-use-a-shared-javascript-runtime"></a>Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

É possível configurar o Suplemento do Office para executar todo o seu código em um único tempo de execução JavaScript compartilhado (também conhecido como tempo de execução compartilhado). Isso permite uma melhor coordenação em seu suplemento e acesso ao DOM e CORS de todas as partes de seu suplemento. Ele também ativa recursos adicionais, como a execução de código quando o documento é aberto ou a ativação ou desativação de botões da faixa de opções. Para configurar seu suplemento para usar um tempo de execução JavaScript compartilhado, siga as instruções neste artigo.

## <a name="create-the-add-in-project"></a>Criar o projeto de suplemento

Se você estiver iniciando um novo projeto, use o [Gerador Yeoman para Suplementos do Office](yeoman-generator-overview.md) para criar um projeto de suplemento do Excel, PowerPoint ou Word.

Execute o comando `yo office --projectType taskpane --name "my office add in" --host <host> --js true`, onde `<host>` é um dos seguintes valores.

- excel
- powerpoint
- palavra

> [!IMPORTANT]
> O `--name` de argumento deve estar entre aspas duplas, mesmo que não tenha espaços.

Você pode usar opções diferentes para as opções de linha de comando **--projecttype**, **--name** e **--js**. Para a lista completa de opções, veja [Gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office).

O gerador criará o projeto e instalará os componentes do Node de suporte. Você também pode usar as etapas neste artigo para atualizar um projeto do Visual Studio para usar o tempo de execução compartilhado. No entanto, talvez seja necessário atualizar os esquemas XML para o manifesto. Para obter mais informações, confira [Solucionar erros de desenvolvimento com Suplementos do Office](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

## <a name="configure-the-manifest"></a>Configurar o manifesto

Siga estas etapas para um projeto novo ou existente para configurá-lo para usar um tempo de execução compartilhado. Estas etapas pressupõem que você gerou seu projeto usando o [Gerador Yeoman para Suplementos do Office ](yeoman-generator-overview.md).

1. Inicie o Visual Studio Code e abra seu projeto de suplemento.
1. Abra o arquivo **manifest.xml**.
1. Para um suplemento do Excel ou PowerPoint, atualize a seção de requisitos para incluir o [tempo de execução compartilhado](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets). Certifique-se de remover o requisito `CustomFunctionsRuntime` se estiver presente. O XML deve aparecer da seguinte maneira.

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

    > [!NOTE]
    > Não adicione o requisito `SharedRuntime` definido ao manifesto para um suplemento do Word. Isso causará um erro ao carregar o suplemento, que é um problema conhecido no momento.

1. Encontre a seção **\<VersionOverrides\>** e adicione a seguinte seção **\<Runtimes\>**. A vida útil deve ser **longa** para que o código do suplemento possa ser executado mesmo quando o painel de tarefas está fechado. O `resid`valor é **Taskpane.Url**, que faz referência ao local do arquivo **taskpane.html** especificado na `<bt:Urls>`seção próxima à parte inferior do arquivo **manifest.xml**.

    > [!IMPORTANT]
    > A seção **\<Runtimes\>** deve ser inserida após o elemento **\<Host\>** na ordem exata mostrada no XML a seguir.

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
         <Runtimes>
           <Runtime resid="Taskpane.Url" lifetime="long" />
         </Runtimes>
       ...
       </Host>
   ```

1. Se você gerou um suplemento do Excel com funções personalizadas, encontre o elemento **\<Page\>**. Em seguida, altere o local de origem de **Functions.Page.Url** para **Taskpane.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. Encontre a marca **\<FunctionFile\>** e altere a marca `resid` de **Commands.Url** para  **Taskpane.Url**. Observe que, se você não tiver comandos de ação, não terá uma entrada **\<FunctionFile\>** e poderá pular esta etapa.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Salve o arquivo **manifest.xml**.

## <a name="configure-the-webpackconfigjs-file"></a>Configurar o arquivo webpack.config.js

O **webpack.config.js** construirá vários carregadores de tempo de execução. É necessário modificá-lo para carregar apenas o tempo de execução JavaScript compartilhado por meio do arquivo **taskpane.html**.

1. Inicie Visual Studio Code e abra o projeto de suplemento gerado.
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

1. Abra o arquivo **taskpane.js**.
1. Substitua todo o conteúdo do arquivo pelo código a seguir. Isso exibirá uma contagem de quantas vezes o painel de tarefas foi aberto. A adição do evento onVisibilityModeChanged só tem suporte em um tempo de execução JavaScript compartilhado.

    ```javascript
    /*global document, Office*/

    let _count = 0;

    Office.onReady(() => {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";

      updateCount(); // Update count on first open.
      Office.addin.onVisibilityModeChanged(function (args) {
        if (args.visibilityMode === "Taskpane") {
          updateCount(); // Update count on subsequent opens.
        }
      });
    });

    function updateCount() {
      _count++;
      document.getElementById("run").textContent = "Task pane opened " + _count + " times.";
    }
    ```

1. Salve suas alterações e execute o projeto.

   ```command line
   npm start
   ```

Cada vez que você abre o painel de tarefas, a contagem de quantas vezes ele foi aberto será incrementada. O valor de **_count** não será perdido porque o tempo de execução compartilhado mantém seu código em execução mesmo quando o painel de tarefas é fechado.

## <a name="runtime-lifetime"></a>Duração do tempo de execução

Ao adicionar o elemento `Runtime`, você também especifica um tempo de vida com um valor de `long` ou `short`. Defina esse valor como `long` para aproveitar as vantagens dos recursos, como iniciar seu suplemento quando o documento for aberto, continuar a executar o código depois que o painel de tarefas for fechado ou usar CORS e DOM de funções personalizadas.

> [!NOTE]
> O valor de vida útil padrão é `short`, mas recomendamos usar `long` em suplementos do Excel, PowerPoint e Word. Se você definir seu tempo de execução como `short` neste exemplo, seu suplemento será iniciado quando um dos botões da faixa de opções for pressionado, mas poderá ser encerrado após a execução do manipulador de faixa de opções. Da mesma forma, o suplemento será iniciado quando o painel de tarefas for aberto, mas pode ser encerrado quando o painel de tarefas for fechado.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> Se o suplemento incluir o elemento `Runtimes` no manifesto (necessário para um runtime compartilhado) e as condições para usar o Microsoft Edge com WebView2 (baseado em Chromium) forem atendidas, ele usará esse controle WebView2. Se as condições não forem atendidas, ele usará o Internet Explorer 11, independentemente da versão do Windows ou Microsoft 365. Para obter mais informações, consulte [Runtimes](/javascript/api/manifest/runtimes) e [Navegadores usados pelos suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="about-the-shared-javascript-runtime"></a>Sobre o tempo de execução de JavaScript compartilhado

No Windows ou Mac, seu suplemento executará o código para botões da faixa de opções, funções personalizadas e o painel de tarefas em ambientes de tempo de execução JavaScript separados. Isso cria limitações, como não poder compartilhar facilmente dados globais e não poder acessar todas as funcionalidades do CORS a partir de uma função customizada.

No entanto, você pode configurar o Suplemento do Office para compartilhar código no mesmo tempo de execução JavaScript (também conhecido como tempo de execução compartilhado). Isso permite uma melhor coordenação entre o suplemento e o acesso ao DOM e CORS do painel de tarefas de todas as partes do suplemento.

Configurar um tempo de execução compartilhado permite os seguintes cenários.

- Seu Suplemento do Office pode usar recursos de UI adicionais.
  - [Ativar e Desativar Comandos de Suplemento](../design/disable-add-in-commands.md)
  - [Execute o código em seu Suplemento do Office quando o documento for aberto](run-code-on-document-open.md)
  - [Mostre ou oculte o painel de tarefas de seu Suplemento do Office ](show-hide-add-in.md)
- Os seguintes itens estão disponíveis somente para suplementos do Excel.
  - [Adicionar atalhos de teclado Personalizados aos Suplementos do Office (pré-visualização)](../design/keyboard-shortcuts.md)
  - [Crie guias contextuais Personalizadas em Suplementos do Office (pré-visualização)](../design/contextual-tabs.md)
  - As funções personalizadas terão suporte CORS completo.
  - Funções personalizadas podem chamar APIs Office.js para ler dados de documentos de planilhas.

Para Office no Windows, o tempo de execução compartilhado usa Microsoft Edge com WebView2 (baseado em Chromium) se as condições para usá-lo forem atendidas conforme explicado em [Navegadores usados por suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Caso contrário, ele usa o Internet Explorer 11. Além disso, todos os botões que seu suplemento exibir na faixa de opções serão executados no mesmo tempo de execução compartilhado. A imagem a seguir mostra como as funções personalizadas, a interface do usuário da faixa de opções e o código do painel de tarefas serão executados no mesmo tempo de execução do JavaScript.

![Diagrama de uma função personalizada, painel de tarefas e botões da faixa de opções em execução em um tempo de execução de navegador compartilhado no Excel.](../images/custom-functions-in-browser-runtime.png)

### <a name="debug"></a>Depurar

Ao usar um tempo de execução compartilhado, não é possível usar o Código do Visual Studio para depurar funções personalizadas no Excel no Windows no momento. Em vez disso, você precisará usar as ferramentas de desenvolvedor. Para obter mais informações, consulte [Depurar suplementos usando ferramentas de desenvolvedor para Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md) ou [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](../testing/debug-add-ins-using-devtools-edge-chromium.md).

### <a name="multiple-task-panes"></a>Vários painéis de tarefas

Não projete seu suplemento para usar vários painéis de tarefas se você planeja usar um tempo de execução compartilhado. Um tempo de execução compartilhado tem suporte para o uso de apenas um único painel de tarefas. Observe que qualquer painel de tarefas sem um `<TaskpaneID>` é considerado um painel de tarefas diferente.

## <a name="see-also"></a>Confira também

- [Chamar APIs do Excel a partir de uma função personalizada](../excel/call-excel-apis-from-custom-function.md)
- [Adicione atalhos de teclado personalizados aos suplementos do Office (pré-visualização)](../design/keyboard-shortcuts.md)
- [Crie guias contextuais Personalizadas em Suplementos do Office (pré-visualização)](../design/contextual-tabs.md)
- [Ativar e Desativar Comandos de Suplemento](../design/disable-add-in-commands.md)
- [Execute o código em seu Suplemento do Office quando o documento for aberto](run-code-on-document-open.md)
- [Mostre ou oculte o painel de tarefas de seu Suplemento do Office ](show-hide-add-in.md)
- [Tutorial: compartilhar dados e eventos entre as funções personalizadas do Excel e do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
