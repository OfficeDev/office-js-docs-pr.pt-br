---
title: Criar um suplemento de Project que usa REST com um serviço OData local do Project Server
description: Saiba como criar um suplemento do painel de tarefas para o Project Professional 2013 que compara os dados de custo e trabalho no projeto ativo com as médias de todos os projetos na instância atual do Project Web App.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: ed0df5aa9957a0ab326f2d2191421087061aeae2
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889566"
---
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a>Criar um suplemento do Project que usa REST com um serviço OData local do Project Server

Este artigo descreve como criar um suplemento de painel tarefas do Project Professional 2013 que compara dados de custo e de trabalho no projeto ativo com médias de todos os projetos da instância atual do Project Web App. O suplemento usa REST com a biblioteca jQuery para acessar o serviço de relatório OData do **ProjectData** no Project Server 2013.

O código deste artigo é baseado em um exemplo desenvolvido por Saurabh Sanghvi e Arvind Iyer, da Microsoft Corporation.

## <a name="prerequisites"></a>Pré-requisitos

A seguir estão os pré-requisitos para a criação de um suplemento do painel de tarefas do Project que lê o serviço **ProjectData** de uma instância do Project Web App em uma instalação local do Project Server 2013.

- Verifique se você instalou os service packs e as atualizações mais recentes do Windows em seu computador de desenvolvimento local. O sistema operacional pode ser Windows 7, Windows 8, Windows Server 2008 ou Windows Server 2012.

- O Project Professional 2013 é necessário para a conexão com o Project Web App. O computador de desenvolvimento deve ter Project Professional 2013 instalado para habilitar a depuração **F5** com o Visual Studio.

    > [!NOTE]
    > Project Standard 2013 também pode hospedar suplementos do painel de tarefas, mas não pode entrar no Project Web App.

- O Visual Studio 2015 com Office Developer Tools para Visual Studio inclui modelos para criar suplementos do Office e do SharePoint. Verifique se você instalou a versão mais recente do Office Developer Tools. Confira a seção _Ferramentas_ de [Download de suplementos do Office e do SharePoint](https://developer.microsoft.com/office/docs).

- Os procedimentos e exemplos de código neste artigo acessam o serviço **ProjectData** do Project Server 2013 em um domínio local. Os métodos jQuery neste artigo não funcionam com o Project na Web.

    Verifique se o serviço **ProjectData** está acessível no computador de desenvolvimento.

### <a name="procedure-1-verify-that-the-projectdata-service-is-accessible"></a>Procedimento 1. Verifique se o serviço ProjectData está acessível

1. Para permitir que o navegador mostre diretamente os dados XML de uma consulta REST, desative o modo de exibição de leitura do feed. Para obter informações sobre como fazer isso no Internet Explorer, consulte o Procedimento 1, etapa 4 em [Feeds OData de consulta para dados de relatório do Project](/previous-versions/office/project-odata/jj163048(v=office.15)).

1. Consulte o **serviço ProjectData** usando o navegador com a seguinte URL: **http://ServerName /ProjectServerName /_api/ProjectData**. Por exemplo, se a instância do Project Web App for  `http://MyServer/pwa`, o navegador mostrará os resultados a seguir.

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/"
        xmlns="https://www.w3.org/2007/app"
        xmlns:atom="https://www.w3.org/2005/Atom">
        <workspace>
            <atom:title>Default</atom:title>
            <collection href="Projects">
                <atom:title>Projects</atom:title>
            </collection>
            <collection href="ProjectBaselines">
                <atom:title>ProjectBaselines</atom:title>
            </collection>
            <!-- ... and 33 more collection elements -->
        </workspace>
        </service>
    ```

1. Pode ser necessário fornecer as credenciais de rede para ver os resultados. Se o navegador exibir "Erro 403, acesso negado", você não tem permissão de logon para essa instância do Project Web App ou há algum problema de rede que exige ajuda administrativa.

## <a name="use-visual-studio-to-create-a-task-pane-add-in-for-project"></a>Usar o Visual Studio para criar um suplemento do painel de tarefas para o Project

O Office Developer Tools para Visual Studio inclui um modelo de suplemento de painel de tarefas para o Project 2013. Se você criar uma solução chamada **HelloProjectOData**, a solução conterá os dois projetos do Visual Studio a seguir:

- O projeto de suplemento usa o nome da solução. Ele inclui o arquivo de manifesto XML para o suplemento e serve para o .NET Framework 4.5. O procedimento 3 mostra as etapas para modificar o manifesto para o suplemento **HelloProjectOData** .

- O projeto Web é denominado **HelloProjectODataWeb**. Ele inclui as páginas da Web, os arquivos JavaScript, os arquivos CSS, as imagens, as referências e os arquivos de configuração para o conteúdo Web no painel de tarefas. O projeto Web serve para o .NET Framework 4. Os procedimentos 4 e 5 mostram como modificar os arquivos no projeto Web para criar a funcionalidade do suplemento **HelloProjectOData** .

### <a name="procedure-2-create-the-helloprojectodata-add-in-for-project"></a>Procedimento 2. Criar o suplemento HelloProjectOData para Project

1. Execute o Visual Studio 2015 como administrador e selecione **Novo Projeto** na página Inicial.

1. Na caixa **de diálogo** Novo Projeto, expanda os nós **Modelos,** **Visual C#** e **Office/SharePoint** e, em seguida, selecione **Suplementos do Office**. Selecione **.NET Framework 4.5.2** na lista suspensa da estrutura de destino na parte superior do painel central e, em seguida, selecione o Suplemento do **Office** (confira a próxima captura de tela).

1. Para colocar os dois projetos do Visual Studio no mesmo diretório, selecione Criar **diretório** para a solução e navegue até o local desejado.

1. No campo **Nome** , digiteHelloProjectOData e escolha **OK**.

    *Figura 1. Criar um Suplemento do Office*

    ![Criar um Suplemento do Office.](../images/pj15-hello-project-o-data-creating-app.png)

1. Na caixa **de diálogo Escolher o tipo de** suplemento, selecione Painel **de tarefas** e escolha **Avançar** (confira a próxima captura de tela).

    *Figura 2. Como escolher o tipo de suplemento a criar*

    ![Escolha o tipo de suplemento a ser criado.](../images/pj15-hello-project-o-data-choose-project.png)

1. Na caixa **de diálogo Escolher os aplicativos** host, desmarque todas  as caixas de seleção, exceto a caixa de seleção Projeto (veja a próxima captura de tela) e escolha **Concluir**.

    *Figura 3. Como escolher o aplicativo host*

    ![Escolha Project como o único aplicativo host.](../images/create-office-add-in.png)

    O Visual Studio cria o **projeto HelloProjectOdata** e o **projeto HelloProjectODataWeb** .

A **pasta AddIn** (consulte a próxima captura de tela) contém o arquivo App.css para estilos CSS personalizados. Na subpasta **Home**, o arquivo Home.html contém referências para arquivos CSS e JavaScript que o suplemento usa, e o conteúdo HTML5 para o suplemento. Além disso, o arquivo Home.js é para o seu código JavaScript personalizado. A pasta **Scripts** inclui os arquivos da biblioteca jQuery. A subpasta **Office** inclui as bibliotecas JavaScript, como office.js e project-15.js, além das bibliotecas de linguagem para cadeias de caracteres padrão nos suplementos do Office. Na pasta **Content**, o arquivo Office.css contém os estilos padrão de todos os Suplementos do Office.

*Figura 4. Exibir os arquivos de projeto Web padrão Gerenciador de Soluções*

![Exiba os arquivos de projeto Web no Gerenciador de Soluções.](../images/pj15-hello-project-o-data-initial-solution-explorer.png)

O manifesto do projeto **HelloProjectOData** é o HelloProjectOData.xml arquivo. Opcionalmente, você pode modificar o manifesto para adicionar uma descrição do suplemento, uma referência a um ícone, informações de linguagem adicionais e outras configurações. O Procedimento 3 simplesmente modifica o nome de exibição e a descrição do suplemento e adiciona um ícone.

Para saber mais sobre o manifesto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md) e [Referência de esquema para manifestos de suplementos do Office (versão 1.1)](../develop/add-in-manifests.md#see-also).

### <a name="procedure-3-modify-the-add-in-manifest"></a>Procedimento 3. Modificar o manifesto do suplemento

1. No Visual Studio, abra o arquivo HelloProjectOData.xml.

1. O nome de exibição padrão é o nome do projeto do Visual Studio ("HelloProjectOData"). Por exemplo, altere o valor padrão do elemento **\<DisplayName\>** para "Hello ProjectData".

1. A descrição padrão também é "HelloProjectOData". Por exemplo, altere o valor padrão do elemento Description para "Testar consultas REST do serviço ProjectData".

1. Adicione um ícone para mostrar na lista **suspensa Suplementos do Office** na **guia PROJETO** da faixa de opções. Você pode adicionar um arquivo de ícone na solução do Visual Studio ou usar uma URL para um ícone.

As etapas a seguir mostram como adicionar um arquivo de ícone à solução do Visual Studio.

1. Em **Gerenciador de Soluções**, vá para a pasta chamada Imagens.

1. Para ser exibido na lista **suspensa Suplementos do Office** , o ícone deve ter 32 x 32 pixels. Por exemplo, instale o SDK do Project 2013, escolha a pasta **Imagens** e adicione o seguinte arquivo do SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`

    Como alternativa, use seu próprio ícone 32 x 32; ou copie a imagem a seguir para um arquivo chamado NewIcon.png e adicione esse arquivo à  `HelloProjectODataWeb\Images` pasta.

    ![Ícone do aplicativo HelloProjectOData.](../images/pj15-hello-project-data-new-icon.jpg)

1. No manifesto HelloProjectOData.xml, **\<IconUrl\>** **\<Description\>** adicione um elemento abaixo do elemento, em que o valor da URL do ícone é o caminho relativo para o arquivo de ícone 32x32. Por exemplo, adicione a seguinte linha: `<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />`. O HelloProjectOData.xml arquivo de manifesto agora contém o seguinte (seu **\<Id\>** valor será diferente):

    ```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82</Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="create-the-html-content-for-the-helloprojectodata-add-in"></a>Criar o conteúdo HTML para o suplemento HelloProjectOData

O **suplemento HelloProjectOData** é um exemplo que inclui depuração e saída de erro; não se destina ao uso em produção. Antes de começar a codificar o conteúdo HTML, projete a interface do usuário e a experiência do usuário para o suplemento e descreva as funções JavaScript que interagem com o código HTML. Para obter mais informações, consulte [As diretrizes de design para suplementos do Office](../design/add-in-design.md).

O painel de tarefas mostra o nome de exibição do suplemento na parte superior, que é o valor **\<DisplayName\>** do elemento no manifesto. O elemento **body** no arquivo HelloProjectOData.html contém outros elementos de interface do usuário, da seguinte maneira:

- Um subtítulo indica a funcionalidade geral ou o tipo de operação, por exemplo, **CONSULTA REST ODATA**.

- O **botão Obter Ponto de Extremidade do** `setOdataUrl` ProjectData chama a função para obter o ponto de extremidade do serviço **ProjectData** e exibi-la em uma caixa de texto. Se o projeto não estiver conectado ao Project Web App, o suplemento chama um identificador de erro para exibir uma mensagem de erro pop-up.

- O **botão Comparar Todos os Projetos** é desabilitado até que o suplemento obtenha um ponto de extremidade OData válido. Quando você seleciona o botão, ele `retrieveOData` chama a função, que usa uma consulta REST para obter o custo do projeto e os dados de trabalho do **serviço ProjectData** .

- Uma tabela exibe os valores médios de custo do projeto, custo real, trabalho e porcentagem concluída. A tabela também compara os valores atuais do projeto ativo com a média. Se o valor atual for maior que a média de todos os projetos, será exibido em vermelho. Se o valor atual for menor que a média, será exibido em verde. Se o valor atual não estiver disponível, a tabela exibirá um **NA azul**.

    A `retrieveOData` função chama a `parseODataResult` função, que calcula e exibe valores para a tabela.

    > [!NOTE]
    > Neste exemplo, os dados de custo e trabalho para o projeto ativo são derivados dos valores publicados. Se você alterar valores no Project, o serviço **ProjectData** não terá as alterações até que o projeto seja publicado.

### <a name="procedure-4-create-the-html-content"></a>Procedimento 4. Criar o conteúdo HTML

1. No elemento **principal** do arquivo Home.html, adicione elementos de **link** adicionais para arquivos CSS usados pelo suplemento. O modelo de projeto do Visual Studio inclui um link para o arquivo App.css que você pode usar para os estilos CSS personalizados.

1. Adicione elementos de **script** adicionais para bibliotecas JavaScript que seu suplemento usa. O modelo de projeto inclui links para os arquivos jQuery- _[versão]_.js, office.js e MicrosoftAjax.js na **pasta Scripts** .

    > [!NOTE]
    > Antes de implantar o suplemento, mude a referência office.js e a referência jQuery para a referência CDN (rede de distribuição de conteúdo). A referência CDN fornece a versão mais recente e melhora o desempenho.

    O **suplemento HelloProjectOData** também usa o arquivo SurfaceErrors.js, que exibe erros em uma mensagem pop-up. Você pode copiar o código da seção  Programação Robusta de Criar seu primeiro suplemento do painel de tarefas para [o Project 2013](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md) usando um editor de texto e, em seguida, adicionar um arquivo SurfaceErrors.js na pasta **Scripts\Office** do projeto **HelloProjectODataWeb**.

    A seguir está o código HTML atualizado para o **elemento de** cabeçalho, com a linha adicional para o SurfaceErrors.js arquivo.

    ```HTML
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

    <!-- Add your CSS styles to the following file. -->
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
    <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
    <script src="../Scripts/jquery-1.7.1.js"></script>

    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

    <!-- Use the local script references for Office.js to enable offline debugging -->
    <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/1.0/Office.js"></script>

    <!-- Add your JavaScript to the following files. -->
    <script src="../Scripts/HelloProjectOData.js"></script>
    <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
    ```

1. No elemento **do** corpo, exclua o código existente do modelo e adicione o código para a interface do usuário. Se um elemento deve ser preenchido com os dados ou manipulado por uma instrução jQuery, deve incluir um atributo **id** exclusivo. No código a seguir, os atributos **de ID** para os elementos **button**, **span** e **td** (definição de célula de tabela) que as funções jQuery usam são mostrados em negrito.

   The following HTML adds a graphic image, which could be a company logo. Você pode usar um logotipo de sua escolha ou copiar o arquivo NewLogo.png do download do SDK do Project 2013 e, em seguida, usar **Gerenciador de Soluções** `HelloProjectODataWeb\Images` para adicionar o arquivo à pasta.

    ```HTML
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br /><br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
                <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
            <table class="infoTable" aria-readonly="True" style="width: 100%;">
                <tr>
                    <td class="heading_leftCol"></td>
                    <td class="heading_midCol"><strong>Average</strong></td>
                    <td class="heading_rightCol"><strong>Current</strong></td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Work</strong></td>
                    <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project % Complete</strong></td>
                    <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
                </tr>
            </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
    ```

## <a name="create-the-javascript-code-for-the-add-in"></a>Criar o código JavaScript para o suplemento

O modelo para um suplemento de painel de tarefas do Project inclui código de inicialização padrão que foi projetado para demonstrar ações get e set básicas para dados em um documento no caso de um suplemento típico do Office 2013. Como o Project 2013 não dá suporte a ações que gravam no projeto ativo e o suplemento **HelloProjectOData** `getSelectedDataAsync` não usa o método, você pode excluir o script `Office.initialize` `setData` `getData` dentro da função e excluir a função e a função no arquivo HelloProjectOData.js padrão.

O JavaScript inclui constantes globais para a consulta REST e variáveis globais que são usadas em várias funções. O **botão Obter Ponto de Extremidade do** `setOdataUrl` ProjectData chama a função, que inicializa as variáveis globais e determina se o Project está conectado ao Project Web App.

O restante do arquivo HelloProjectOData.js inclui duas funções: `retrieveOData` a função é chamada quando o usuário seleciona Comparar Todos os **Projetos;**`parseODataResult` e a função calcula médias e preenche a tabela de comparação com valores formatados para cor e unidades.

### <a name="procedure-5-create-the-javascript-code"></a>Procedimento 5. Criar o código JavaScript

1. Exclua todo o código no arquivo HelloProjectOData.js padrão e adicione as variáveis globais e a `Office.initialize` função. Nomes de variáveis que são todas maiúsculas implicam que são constantes; eles são usados posteriormente com a `_pwa` variável para criar a consulta REST neste exemplo.

    ```js
    var PROJDATA = "/_api/ProjectData";
    var PROJQUERY = "/Projects?";
    var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
    var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
    var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
    var _pwa;           // URL of Project Web App.
    var _projectUid;    // GUID of the active project.
    var _docUrl;        // Path of the project document.
    var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
        });
    }
    ```

1. Adicionar `setOdataUrl` e funções relacionadas. A `setOdataUrl` função chama `getProjectGuid` e `getDocumentUrl` inicializa as variáveis globais. No método [getProjectFieldAsync](/javascript/api/office/office.document),  `removeAttr` a função anônima para o parâmetro  de retorno de chamada habilita o botão Comparar Todos os Projetos usando o método na biblioteca jQuery e exibe a URL do serviço **ProjectData**. Se o Project não estiver conectado ao Project Web App, a função gera um erro e exibe uma mensagem de erro pop-up. O SurfaceErrors.js arquivo inclui o `throwError` método.

   > [!NOTE]
   > Se você executar o Visual Studio no computador do Project Server, para usar a depuração **F5** , `_pwa` remova o comentário do código após a linha que inicializa a variável global. Para habilitar o uso do método jQuery `ajax` ao depurar no computador do Project Server, `localhost` você deve definir o valor para a URL do PWA. Se você executar o Visual Studio em um computador remoto, `localhost` a URL não será necessária. Before you deploy the add-in, comment out that code.

    ```js
    function setOdataUrl() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _pwa = String(asyncResult.value.fieldValue);

                    // If you debug with Visual Studio on a local Project Server computer, 
                    // uncomment the following lines to use the localhost URL.
                    //var localhost = location.host.split(":", 1);
                    //var pwaStartPosition = _pwa.lastIndexOf("/");
                    //var pwaLength = _pwa.length - pwaStartPosition;
                    //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                    //_pwa = location.protocol + "//" + localhost + pwaName;

                    if (_pwa.substring(0, 4) == "http") {
                        _odataUrl = _pwa + PROJDATA;
                        $("#compareProjects").removeAttr("disabled");
                        getProjectGuid();
                    }
                    else {
                        _odataUrl = "No connection!";
                        throwError(_odataUrl, "You are not connected to Project Web App.");
                    }
                    getDocumentUrl();
                    $("#projectDataEndPoint").text(_odataUrl);
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the GUID of the active project.
    function getProjectGuid() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _projectUid = asyncResult.value.fieldValue;
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the path of the project in Project web app, which is in the form <>\ProjectName .
    function getDocumentUrl() {
        _docUrl = "Document path:\r\n" + Office.context.document.url;
    }
    ```

1. Adicione a `retrieveOData` função, que concatena valores para a consulta REST `ajax` e, em seguida, chama a função em jQuery para obter os dados solicitados do serviço **ProjectData** . A `support.cors` variável habilita o CORS (compartilhamento de recursos entre origens) com a `ajax` função. Se a `support.cors` instrução estiver ausente ou estiver definida como `false`, a `ajax` função retornará um **erro De transporte** .

   > [!NOTE]
   > O seguinte código funciona com uma instalação no local do Project Server 2013. Para o Project na Web, use o OAuth para autenticação baseada em token. Para saber mais, confira [Como lidar com limitações de política de mesma origem nos Suplementos do Office](../develop/addressing-same-origin-policy-limitations.md).

   Na chamada `ajax` , você pode usar o parâmetro _de cabeçalho_ ou o _parâmetro beforeSend_ . O _parâmetro_ completo é uma função anônima para que ele esteja no mesmo escopo que as variáveis em `retrieveOData`. A função para o  _parâmetro completo_ exibe os resultados `odataText` no controle `parseODataResult` e também chama o método para analisar e exibir a resposta JSON. O _parâmetro_ de erro especifica a função `getProjectDataErrorHandler` nomeada, `odataText` `throwError` que grava uma mensagem de erro no controle e também usa o método para exibir uma mensagem pop-up.

    ```js
    // Functions to get and parse the Project Server reporting data./

    // Get data about all projects on Project Server,
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();

        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project on the web.
        $.support.cors = true;

        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json",
            data: "",      // Empty string for the optional data.
            //headers: { "Accept": accept },
            beforeSend: function (xhr) {
                xhr.setRequestHeader("ACCEPT", accept);
            },
            complete: function (xhr, textStatus) {
                // Create a message to display in the text box.
                var message = "\r\ntextStatus: " + textStatus +
                    "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                    "\r\nStatus: " + xhr.status +
                    "\r\nResponseText:\r\n" + xhr.responseText;

                // xhr.responseText is the result from an XmlHttpRequest, which
                // contains the JSON response from the OData service.
                parseODataResult(xhr.responseText, _projectUid);

                // Write the document name, response header, status, and JSON to the odataText control.
                $("#odataText").text(_docUrl);
                $("#odataText").append("\r\nREST query:\r\n" + restUrl);
                $("#odataText").append(message);

                if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                    $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
                }
            },
            error: getProjectDataErrorHandler
        });
    }

    function getProjectDataErrorHandler(data, errorCode, errorMessage) {
        $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
        throwError(errorCode, errorMessage);
    }
    ```

1. Adicione o `parseODataResult` método, que desserializa e processa a resposta JSON do serviço OData. `parseODataResult` O método calcula os valores médios dos dados de custo e trabalho com uma precisão de uma ou duas casas decimais, formata valores com a cor correta e adiciona uma unidade ( **$**, **hrs****%** ou ), e exibe os valores em células de tabela especificadas.

   Se o GUID do projeto ativo corresponder ao `ProjectId` valor, a variável `myProjectIndex` será definida como o índice do projeto. Se `myProjectIndex` indicar que o projeto ativo foi publicado no Project Server, `parseODataResult` o método formata e exibe dados de custo e trabalho para esse projeto. If the active project is not published, values for the active project are displayed as a blue **NA**.

    ```js
    // Calculate the average values of actual cost, cost, work, and percent complete
    // for all projects, and compare with the values for the current project.
    function parseODataResult(oDataResult, currentProjectGuid) {
        // Deserialize the JSON string into a JavaScript object.
        var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
        var len = res.d.results.length;
        var projActualCost = 0;
        var projCost = 0;
        var projWork = 0;
        var projPercentCompleted = 0;
        var myProjectIndex = -1;
        for (i = 0; i < len; i++) {
            // If the current project GUID matches the GUID from the OData query,  
            // store the project index.
            if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
                myProjectIndex = i;
            }
            projCost += Number(res.d.results[i].ProjectCost);
            projWork += Number(res.d.results[i].ProjectWork);
            projActualCost += Number(res.d.results[i].ProjectActualCost);
            projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);
        }
        var avgProjCost = projCost / len;
        var avgProjWork = projWork / len;
        var avgProjActualCost = projActualCost / len;
        var avgProjPercentCompleted = projPercentCompleted / len;

        // Round off cost to two decimal places, and round off other values to one decimal place.
        avgProjCost = avgProjCost.toFixed(2);
        avgProjWork = avgProjWork.toFixed(1);
        avgProjActualCost = avgProjActualCost.toFixed(2);
        avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

        // Display averages in the table, with the correct units.
        document.getElementById("AverageProjectCost").innerHTML = "$"
            + avgProjCost;
        document.getElementById("AverageProjectActualCost").innerHTML
            = "$" + avgProjActualCost;
        document.getElementById("AverageProjectWork").innerHTML
            = avgProjWork + " hrs";
        document.getElementById("AverageProjectPercentComplete").innerHTML
            = avgProjPercentCompleted + "%";

        // Calculate and display values for the current project.
        if (myProjectIndex != -1) {
            var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
            var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
            var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
            var myProjPercentCompleted =
            Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

            myProjCost = myProjCost.toFixed(2);
            myProjWork = myProjWork.toFixed(1);
            myProjActualCost = myProjActualCost.toFixed(2);
            myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

            document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

            if (Number(myProjCost) <= Number(avgProjCost)) {
                document.getElementById("CurrentProjectCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectCost").style.color = "red"
            }

            document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

            if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
                document.getElementById("CurrentProjectActualCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectActualCost").style.color = "red"
            }

            document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

            if (Number(myProjWork) <= Number(avgProjWork)) {
                document.getElementById("CurrentProjectWork").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectWork").style.color = "green"
            }

            document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

            if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
                document.getElementById("CurrentProjectPercentComplete").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectPercentComplete").style.color = "green"
            }
        }
        else {
            document.getElementById("CurrentProjectCost").innerHTML = "NA";
            document.getElementById("CurrentProjectCost").style.color = "blue"

            document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
            document.getElementById("CurrentProjectActualCost").style.color = "blue"

            document.getElementById("CurrentProjectWork").innerHTML = "NA";
            document.getElementById("CurrentProjectWork").style.color = "blue"

            document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
            document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
        }
    }
    ```

## <a name="test-the-helloprojectodata-add-in"></a>Testar o suplemento HelloProjectOData

Para testar e depurar o suplemento **HelloProjectOData** com o Visual Studio 2015, o Project Professional 2013 deve ser instalado no computador de desenvolvimento. Para habilitar cenários de teste diferentes, certifique-se de poder escolher se o Project abre no caso de arquivos no computador local ou se ele se conecta ao Project Web App. Por exemplo, execute as etapas a seguir.

1. Na guia **Arquivo** , escolha a **guia Informações** no modo de exibição Backstage e, em seguida, **escolha Gerenciar Contas**.

1. Na caixa **de diálogo** Contas do aplicativo Web do Project, **a** lista de contas disponíveis pode ter várias contas do Project Web App, além da conta de **computador** local. Na seção **Ao iniciar**, selecione **Escolher uma conta**.

1. Feche o Project para que o Visual Studio possa iniciá-lo na depuração do suplemento.

Os testes básicos devem incluir o seguinte:

- Execute o suplemento no Visual Studio, e abra um projeto publicado do Project Web App que contém dados de custos e trabalho. Verifique se o suplemento exibe o ponto de extremidade **ProjectData** e exibe corretamente os dados de custo e trabalho na tabela. Você pode usar a saída no controle **odataText** para verificar a consulta REST e outras informações.

- Execute o suplemento novamente, onde você escolhe o perfil do computador local **na caixa de** diálogo Logon quando o Project é iniciado. Abra um arquivo .mpp local e teste o suplemento. Verifique se o suplemento exibe uma mensagem de erro ao tentar acessar o ponto de extremidade **ProjectData**.

- Execute o suplemento novamente e crie um projeto com tarefas com dados de custo e de trabalho. Você pode salvar o projeto no Project Web App, mas não o publique. Verifique se o suplemento exibe dados do Project Server, mas **NA** para o projeto atual.

### <a name="procedure-6-test-the-add-in"></a>Procedimento 6. Testar o suplemento

1. Execute o Project Professional 2013, conecte-se ao Project Web App e crie um projeto de teste. Atribua tarefas aos recursos locais ou a recursos da empresa, defina vários valores de porcentagem concluída em algumas tarefas e publique o projeto. Feche o projeto, o que permite que o Visual Studio inicie o Project para depurar o suplemento.

1. No Visual Studio, pressione **F5**. Faça logon no Project Web App e abra o projeto que você criou na etapa anterior. Você pode abrir o projeto no modo somente leitura ou no modo de edição.

1. Na guia **PROJETO** da faixa de opções, na lista suspensa **Suplementos do Office** , selecione **Hello ProjectData** (consulte a Figura 5). O botão **Comparar Todos os Projetos** deve estar desativado.

    *Figura 5. Iniciar o suplemento HelloProjectOData*

    ![Teste o aplicativo HelloProjectOData.](../images/pj15-hello-project-data-test-the-app.png)

1. No painel **de tarefas Hello ProjectData** , selecione **Obter Ponto de Extremidade do ProjectData**. A **linha projectDataEndPoint** deve mostrar a URL do serviço **ProjectData** e o botão Comparar  Todos os Projetos deve ser habilitado (consulte a Figura 6).

1. Selecione **Comparar Todos os Projetos**. O suplemento pode pausar enquanto recupera dados do serviço **ProjectData** e, em seguida, deve exibir a média formatada e os valores atuais na tabela.

    *Figura 6. Exibir resultados da consulta REST*

    ![Exiba os resultados da consulta REST.](../images/pj15-hello-project-data-rest-results.png)

1. Examine a saída na caixa de texto. Ele deve mostrar o caminho do documento, a consulta REST, as informações de status e os resultados JSON das chamadas para `ajax` e `parseODataResult`. A saída ajuda a entender, criar e depurar código no `parseODataResult` método, como `projCost += Number(res.d.results[i].ProjectCost);`.

    A seguir está um exemplo da saída com quebras de linha e espaços adicionados ao texto para maior clareza, para três projetos em uma instância do Project Web App.

    ```json
    Document path: <>\WinProj test1

    REST query:
    http://sphvm-37189/pwa/_api/ProjectData/Projects?$filter=ProjectName ne 'Timesheet Administrative Work Items'
        &amp;$select=ProjectId, ProjectName, ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost

    textStatus: success
    ContentType: application/json;odata=verbose;charset=utf-8
    Status: 200

    ResponseText:
    {"d":{"results":[
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "type":"ReportingData.Project"},
        "ProjectId":"ce3d0d65-3904-e211-96cd-00155d157123",
        "ProjectActualCost":"0.000000",
        "ProjectCost":"0.000000",
        "ProjectName":"Task list created in PWA",
        "ProjectPercentCompleted":0,
        "ProjectWork":"16.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"c31023fc-1404-e211-86b2-3c075433b7bd",
        "ProjectActualCost":"700.000000",
        "ProjectCost":"2400.000000",
        "ProjectName":"WinProj test 2",
        "ProjectPercentCompleted":29,
        "ProjectWork":"48.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"dc81fbb2-b801-e211-9d2a-3c075433b7bd",
        "ProjectActualCost":"1900.000000",
        "ProjectCost":"5200.000000",
        "ProjectName":"WinProj test1",
        "ProjectPercentCompleted":37,
        "ProjectWork":"104.000000"}
    ]}}
    ```

1. Pare a depuração (pressione **Shift + F5**) e pressione **F5** novamente para executar uma nova instância do Project. Na caixa de diálogo **Login**, escolha o perfil local **Computador**, e não o Project Web App. Crie ou abra um arquivo .mpp de projeto local, abra o painel de tarefas **Hello ProjectData** e selecione **Obter Ponto de Extremidade de ProjectData**. O suplemento deve mostrar uma **conexão Sem!** (consulte a Figura 7) e o **botão Comparar Todos os Projetos** deve permanecer desabilitado.

   *Figura 7. Usar o suplemento sem uma conexão de aplicativo Web do Project*

   ![Use o aplicativo sem uma conexão do Project Web App.](../images/pj15-hello-project-data-no-connection.png)

1. Pare a depuração e pressione **F5** novamente. Faça logon no Project Web App e crie um projeto com dados de custo e de trabalho. Você pode salvar o projeto, mas não o publique.

   No painel **de tarefas Hello ProjectData**, ao selecionar Comparar Todos os **Projetos, você** deverá ver um **NA** azul para campos na coluna Atual  (consulte a Figura 8).

   *Figura 8. Comparar um projeto não publicado com outros projetos*

   ![Compare um projeto não publicado com outras pessoas.](../images/pj15-hello-project-data-not-published.png)

Mesmo que seu suplemento tenha funcionado corretamente nos testes anteriores, há outros testes que devem ser executados. Por exemplo:

- Abra um projeto do Project Web App que não tenha nenhum dado de custo ou de trabalho para as tarefas. Você deve ver valores de zero nos campos na **coluna** Atual.

- Teste um projeto sem tarefas.

- Se você modificar o suplemento e publicá-lo, deve executar testes semelhantes novamente com o suplemento publicado. Para outras considerações, confira [Próximas etapas](#next-steps).

> [!NOTE]
> Há limites para a quantidade de dados que podem ser retornados em uma consulta do serviço **ProjectData** ; a quantidade de dados varia de acordo com a entidade. Por exemplo, o conjunto `Projects` de entidades tem um limite padrão de 100 projetos por consulta, `Risks` mas o conjunto de entidades tem um limite padrão de 200. For a production installation, the code in the **HelloProjectOData** example should be modified to enable queries of more than 100 projects. For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).

## <a name="example-code-for-the-helloprojectodata-add-in"></a>Exemplo de código para o suplemento de HelloProjectOData

### <a name="helloprojectodatahtml-file"></a>Arquivo HelloProjectOData.html

O código a seguir está no arquivo `Pages\HelloProjectOData.html` do projeto **HelloProjectODataWeb**.

```HTML
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Test ProjectData Service</title>

        <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

        <!-- Add your CSS styles to the following file. -->
        <link rel="stylesheet" type="text/css" href="../Content/App.css" />

        <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
        <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
        <script src="../Scripts/jquery-1.7.1.js"></script>

        <!-- Use the CDN reference to Office.js when deploying your add-in -->
        <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

        <!-- Use the local script references for Office.js to enable offline debugging -->
        <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
        <script src="../Scripts/Office/1.0/Office.js"></script>

        <!-- Add your JavaScript to the following files. -->
        <script src="../Scripts/HelloProjectOData.js"></script>
        <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br />
            <br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">
            Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
        <table class="infoTable" aria-readonly="True" style="width: 100%;">
            <tr>
            <td class="heading_leftCol"></td>
            <td class="heading_midCol"><strong>Average</strong></td>
            <td class="heading_rightCol"><strong>Current</strong></td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Cost</strong></td>
            <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
            <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Work</strong></td>
            <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project % Complete</strong></td>
            <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
            </tr>
        </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
</html>
```

### <a name="helloprojectodatajs-file"></a>Arquivo HelloProjectOData.js

O código a seguir está no arquivo `Scripts\Office\HelloProjectOData.js` do projeto **HelloProjectODataWeb**.

```js
/* File: HelloProjectOData.js
* JavaScript functions for the HelloProjectOData example task pane app.
* October 2, 2012
*/

var PROJDATA = "/_api/ProjectData";
var PROJQUERY = "/Projects?";
var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
var _pwa;           // URL of Project Web App.
var _projectUid;    // GUID of the active project.
var _docUrl;        // Path of the project document.
var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
    });
}

// Set the global variables, enable the Compare All Projects button,
// and display the URL of the ProjectData service.
// Display an error if Project is not connected with Project Web App.
function setOdataUrl() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.ProjectServerUrl,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _pwa = String(asyncResult.value.fieldValue);

                // If you debug with Visual Studio on a local Project Server computer,
                // uncomment the following lines to use the localhost URL.
                //var localhost = location.host.split(":", 1);
                //var pwaStartPosition = _pwa.lastIndexOf("/");
                //var pwaLength = _pwa.length - pwaStartPosition;
                //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                //_pwa = location.protocol + "//" + localhost + pwaName;

                if (_pwa.substring(0, 4) == "http") {
                    _odataUrl = _pwa + PROJDATA;
                    $("#compareProjects").removeAttr("disabled");
                    getProjectGuid();
                }
                else {
                    _odataUrl = "No connection!";
                    throwError(_odataUrl, "You are not connected to Project Web App.");
                }
                getDocumentUrl();
                $("#projectDataEndPoint").text(_odataUrl);
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the GUID of the active project.
function getProjectGuid() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.GUID,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _projectUid = asyncResult.value.fieldValue;
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the path of the project in Project web app, which is in the form <>\ProjectName .
function getDocumentUrl() {
    _docUrl = "Document path:\r\n" + Office.context.document.url;
}

//  Functions to get and parse the Project Server reporting data./

// Get data about all projects on Project Server,
// by using a REST query with the ajax method in jQuery.
function retrieveOData() {
    var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
    var accept = "application/json; odata=verbose";
    accept.toLocaleLowerCase();

    // Enable cross-origin scripting (required by jQuery 1.5 and later).
    // This does not work with Project on the web.
    $.support.cors = true;

    $.ajax({
        url: restUrl,
        type: "GET",
        contentType: "application/json",
        data: "",      // Empty string for the optional data.
        //headers: { "Accept": accept },
        beforeSend: function (xhr) {
            xhr.setRequestHeader("ACCEPT", accept);
        },
        complete: function (xhr, textStatus) {
            // Create a message to display in the text box.
            var message = "\r\ntextStatus: " + textStatus +
                "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                "\r\nStatus: " + xhr.status +
                "\r\nResponseText:\r\n" + xhr.responseText;

            // xhr.responseText is the result from an XmlHttpRequest, which 
            // contains the JSON response from the OData service.
            parseODataResult(xhr.responseText, _projectUid);

            // Write the document name, response header, status, and JSON to the odataText control.
            $("#odataText").text(_docUrl);
            $("#odataText").append("\r\nREST query:\r\n" + restUrl);
            $("#odataText").append(message);

            if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
            }
        },
        error: getProjectDataErrorHandler
    });
}

function getProjectDataErrorHandler(data, errorCode, errorMessage) {
    $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
    throwError(errorCode, errorMessage);
}

// Calculate the average values of actual cost, cost, work, and percent complete
// for all projects, and compare with the values for the current project.
function parseODataResult(oDataResult, currentProjectGuid) {
    // Deserialize the JSON string into a JavaScript object.
    var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
    var len = res.d.results.length;
    var projActualCost = 0;
    var projCost = 0;
    var projWork = 0;
    var projPercentCompleted = 0;
    var myProjectIndex = -1;

    for (i = 0; i < len; i++) {
        // If the current project GUID matches the GUID from the OData query,  
        // then store the project index.
        if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
            myProjectIndex = i;
        }
        projCost += Number(res.d.results[i].ProjectCost);
        projWork += Number(res.d.results[i].ProjectWork);
        projActualCost += Number(res.d.results[i].ProjectActualCost);
        projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);

    }
    var avgProjCost = projCost / len;
    var avgProjWork = projWork / len;
    var avgProjActualCost = projActualCost / len;
    var avgProjPercentCompleted = projPercentCompleted / len;

    // Round off cost to two decimal places, and round off other values to one decimal place.
    avgProjCost = avgProjCost.toFixed(2);
    avgProjWork = avgProjWork.toFixed(1);
    avgProjActualCost = avgProjActualCost.toFixed(2);
    avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

    // Display averages in the table, with the correct units. 
    document.getElementById("AverageProjectCost").innerHTML = "$"
        + avgProjCost;
    document.getElementById("AverageProjectActualCost").innerHTML
        = "$" + avgProjActualCost;
    document.getElementById("AverageProjectWork").innerHTML
        = avgProjWork + " hrs";
    document.getElementById("AverageProjectPercentComplete").innerHTML
        = avgProjPercentCompleted + "%";

    // Calculate and display values for the current project.
    if (myProjectIndex != -1) {

        var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
        var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
        var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
        var myProjPercentCompleted = Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

        myProjCost = myProjCost.toFixed(2);
        myProjWork = myProjWork.toFixed(1);
        myProjActualCost = myProjActualCost.toFixed(2);
        myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

        document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

        if (Number(myProjCost) <= Number(avgProjCost)) {
            document.getElementById("CurrentProjectCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectCost").style.color = "red"
        }

        document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

        if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
            document.getElementById("CurrentProjectActualCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectActualCost").style.color = "red"
        }

        document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

        if (Number(myProjWork) <= Number(avgProjWork)) {
            document.getElementById("CurrentProjectWork").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectWork").style.color = "green"
        }

        document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

        if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
            document.getElementById("CurrentProjectPercentComplete").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectPercentComplete").style.color = "green"
        }
    }
    else {    // The current project is not published.
        document.getElementById("CurrentProjectCost").innerHTML = "NA";
        document.getElementById("CurrentProjectCost").style.color = "blue"

        document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
        document.getElementById("CurrentProjectActualCost").style.color = "blue"

        document.getElementById("CurrentProjectWork").innerHTML = "NA";
        document.getElementById("CurrentProjectWork").style.color = "blue"

        document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
        document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
    }
}
```

### <a name="appcss-file"></a>Arquivo App.css

O código a seguir está no arquivo `Content\App.css` do projeto **HelloProjectODataWeb**.

```css
/*
*  File: App.css for the HelloProjectOData app.
*  Updated: 10/2/2012
*/

body
{
    font-size: 11pt;
}
h1
{
    font-size: 22pt;
}
h2
{
    font-size: 16pt;
}

/******************************************************************
Code label class
******************************************************************/

.rest 
{
    font-family: 'Courier New';
    font-size: 0.9em;
}

/******************************************************************
Button classes
******************************************************************/

.button-wide {
    width: 210px;
    margin-top: 2px;
}
.button-narrow 
{
    width: 80px;
    margin-top: 2px;
}

/******************************************************************
Table styles
******************************************************************/

.infoTable
{
    text-align: center; 
    vertical-align: middle
}
.heading_leftCol
{
    width: 20px;
    height: 20px;
}
.heading_midCol
{
    width: 100px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.heading_rightCol
{
    width: 101px;
    height: 20px;
    font-size: medium;
    font-weight: bold;
}
.row_leftCol
{
    width: 20px;
    font-size: small;
    font-weight: bold;
}
.row_midCol
{
    width: 100px;
}
.row_rightCol
{
    width: 101px;
}
.logo
{
    width: 135px;
    height: 53px;
}
```

### <a name="surfaceerrorsjs-file"></a>Arquivo SurfaceErrors.js

Você pode copiar o código para o arquivo SurfaceErrors.js da seção _Programação Robusta_ de [Criar seu primeiro suplemento de painel de tarefas para o Project 2013 usando um editor de texto](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).

## <a name="next-steps"></a>Próximas etapas

Se **HelloProjectOData** fosse um suplemento de produção a ser vendido no AppSource ou distribuído em um catálogo de aplicativos do SharePoint, ele seria projetado de forma diferente. Por exemplo, não haveria nenhuma saída de depuração em uma caixa de texto e provavelmente nenhum botão para obter o ponto de extremidade **ProjectData**. Você também precisaria reescrever a função `retireveOData` para lidar com instâncias do Project Web App que têm mais de 100 projetos.

O suplemento deveria conter mais verificações de erro, além de lógica para capturar e explicar ou mostrar casos extremos. Por exemplo, se uma instância do Project Web App tiver mil projetos com uma duração média de cinco dias e custo médio de US$ 2.400, e o projeto ativo for o único que tem uma duração de mais de 20 dias, a comparação de custo e trabalho poderá ficar desequilibrada. Isso poderia ser exibido com um gráfico de frequência. Você poderia adicionar opções para exibir a duração, comparar projetos de tamanhos semelhantes ou comparar projetos de um mesmo departamento ou de departamentos diferentes. Ou poderia adicionar uma forma de o usuário selecionar os campos a exibir em uma lista.

Para outras consultas do serviço **ProjectData** , há limites para o comprimento da cadeia de caracteres de consulta, o que afeta o número de etapas que uma consulta pode executar de uma coleção pai para um objeto em uma coleção filho. Por exemplo, uma consulta de duas etapas de **Projects** para **Tasks** para itens de tarefa funciona, mas uma consulta de três etapas, como **Projects** para **Tasks** para **Assignments** para itens de atribuição pode exceder o comprimento máximo de URL padrão. Para obter mais informações, consulte [Feeds OData de consulta para dados de relatório do Project](/previous-versions/office/project-odata/jj163048(v=office.15)).

Se você modificar o **suplemento HelloProjectOData** para uso em produção, execute as etapas a seguir.

- No arquivo HelloProjectOData.html, para obter melhor desempenho, mude a referência ao office.js do projeto local para a referência da CDN:

    ```HTML
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

- Reescreva `retrieveOData` a função para habilitar consultas de mais de 100 projetos. Por exemplo, você pode obter o número de projetos com uma consulta `~/ProjectData/Projects()/$count` e usar os operadores _$skip_ e _$top_ na consulta REST para dados de projeto. Execute várias consultas em sequência e tire a média dos dados de cada consulta. Cada consulta para dados do projeto seria do formato:

  `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`

  Para obter mais informações, consulte [as opções de consulta do sistema OData usando o ponto de extremidade REST](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps&preserve-view=true) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).

- Para implantar o suplemento, confira [Publicar seu suplemento do Office](../publish/publish.md).

## <a name="see-also"></a>Confira também

- [Suplementos do painel de tarefas para Project](project-add-ins.md)
- [Criar seu primeiro suplemento de painel de tarefas para o Project 2013 usando um editor de texto](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- [ProjectData - referência do serviço OData do Project](/previous-versions/office/project-odata/jj163015(v=office.15))
- [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)
