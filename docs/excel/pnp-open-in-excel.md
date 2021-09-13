---
title: Abra Excel página da Web e insiro seu Office Dep.
description: Abra Excel página da Web e insiro seu Office Add-in.
ms.date: 02/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0ac644de03c1f3a4c382dbe151c3224afffdbc81
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148592"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Abra Excel página da Web e insiro seu Office Dep.

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Imagem do Excel na página da Web abrindo um novo documento Excel com o seu add-in incorporado e abrindo automaticamente.":::

Estenda seu aplicativo Web SaaS para que seus clientes possam abrir seus dados de uma página da Web diretamente para Microsoft Excel. Um cenário comum é que os clientes trabalharão com dados em seu aplicativo Web. Em seguida, eles vão querer copiar os dados em um Excel documento. Por exemplo, eles podem querer executar análises adicionais usando Excel. Normalmente, o cliente é obrigado a exportar os dados para um arquivo, como um arquivo .csv, e depois importar esses dados para Excel. Eles também precisam adicionar manualmente o seu Office Add-in ao documento.

Reduza o número de etapas para um único botão clique em sua página da Web que gera e abre o Excel documento. Você também pode inserir seu Office de usuário dentro do documento e exibi-lo quando o documento for aberto. Isso garante que o cliente ainda tenha acesso aos recursos do aplicativo. Quando o documento é aberto, os dados selecionados pelo cliente e o seu Office Dedados já estão disponíveis para que eles continuem trabalhando.

Este artigo mostra o código e as técnicas para implementar esse cenário em seu próprio aplicativo Web SaaS.

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>Criar um novo Excel e inserir um Office Dep.

Primeiro, vamos aprender a criar um documento Excel de uma página da Web e inserir um complemento no documento. O Office de código de entrada [OOXML](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) mostra como inserir o Script Lab [do](https://appsource.microsoft.com/product/office/wa104380862) Script Lab em um novo documento Office. Embora o exemplo funcione com qualquer documento Office, vamos nos concentrar Excel planilhas neste artigo. Use as etapas a seguir para criar e executar o exemplo.

1. Extraia o código de exemplo  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip em uma pasta no computador.
2. Para criar e executar o exemplo, siga as etapas na seção **Para usar o** projeto do readme.
3. Quando você executar o exemplo, ela exibirá uma página da Web semelhante à captura de tela a seguir. Use a página da Web para criar um novo documento Excel que contém Script Lab quando ele é aberto.
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Captura de tela da página da Web que o exemplo de laboratório de scripts de incorporação exibe para selecionar um arquivo Excel e incorporar o complemento do laboratório de script nele.":::

### <a name="how-the-sample-works"></a>Como o exemplo funciona

O código de exemplo usa o SDK OOXML para incorporar o Script Lab do Script Lab ao documento de Excel que você escolher. As informações a seguir são retiradas da [ **seção Sobre o código**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) no arquivo readme.

O arquivo **Home.aspx.cs**:

- Fornece manipuladores de eventos de botão e manipulação básica da interface do usuário.
- Usa técnicas ASP.NET padrão para carregar e baixar o arquivo.
- Usa a extensão de nome de arquivo do arquivo carregado (xlsx, docx ou pptx) para determinar o tipo de arquivo. Isso precisa ser feito no início porque o SDK Open XML geralmente tem APIs distintas para cada tipo de arquivo.
- Chama o **OOXMLHelper** para validar o arquivo e chama o **AddInEmbedder** para inserir Script Lab no arquivo e definir para abrir automaticamente.

O arquivo **AddInEmbedder.cs**:

- Fornece a principal lógica de negócios, que neste exemplo é um método que incorpora Script Lab.
- Faz chamadas para o auxiliar OOXML com base no tipo do arquivo.

O arquivo **OOXMLHelper.cs**:

- Fornece toda a manipulação OOXML detalhada.
- Usa uma técnica padrão para validar o arquivo Office, que é simplesmente chamar o **método Document.Open** nele. Se o arquivo for inválido, o método lançará uma exceção.
- Contém principalmente o código gerado pelas Ferramentas de Produtividade do SDK open XML 2.5 que estão disponíveis no link para o [SDK Open XML 2.5](/office/open-xml/open-xml-sdk).

O **método GenerateWebExtensionPart1Content** no arquivo **OOXMLHelper.cs** define a referência à ID do Script Lab no Microsoft AppSource:

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- O **valor StoreType** é "OMEX", um alias do Microsoft AppSource.
- O **valor** da Loja é "en-US" encontrado na seção Cultura do Microsoft AppSource para Script Lab.
- O **valor da Id** é a ID de ativo do Microsoft AppSource para Script Lab.

Se você estiver configurando um complemento de um catálogo de compartilhamento de arquivos para abertura automática, usará valores diferentes:

O **valor StoreType** é "FileSystem".

- O **valor** da Loja é a URL do compartilhamento de rede; por exemplo, " \\ \\ MyComputer \\ MySharedFolder". Essa deve ser a URL exata que aparece como o Endereço de Catálogo Confiável do compartilhamento na central de Office Trust Center.
- O **valor de Id** é a ID do aplicativo no manifesto dos complementos.
> [!NOTE]
> Para obter mais informações sobre valores alternativos para esses atributos, consulte [Abrir automaticamente](../develop/automatically-open-a-task-pane-with-a-document.md)um painel de tarefas com um documento .

## <a name="use-the-fluent-ui"></a>Usar a interface Fluent interface do usuário

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Fluent Ícones de interface do usuário para Word, Excel e PowerPoint.":::

Uma prática prática é usar a interface do usuário Fluent para ajudar os usuários a fazer a transição entre os produtos Microsoft. Você sempre deve usar um ícone Office para indicar qual aplicativo Office será lançado em sua página da Web. Vamos modificar o código de exemplo para usar o ícone Excel para indicar que ele inicia o Excel aplicativo.

1. Abra o exemplo em Visual Studio.
1. Abra a **página Home.aspx.**
1. Encontre o código a seguir que é o botão de download no formulário.

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. Substitua o código do botão pela seguinte marca de imagem.

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. Pressione **F5** (ou **Depurar > Iniciar Depuração**). Você verá o ícone aparecer quando a home page for carregada.

Para obter mais informações, [consulte Office Ícones](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) de Marca no portal Fluent de desenvolvedores da interface do usuário.  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Upload o documento Excel para Microsoft OneDrive

Recomendamos carregar novos documentos para OneDrive se seu cliente usa OneDrive. Isso torna mais fácil para eles encontrar e trabalhar com os documentos. Vamos criar um novo exemplo de código e ver como você pode usar o SDK do Microsoft Graph para carregar um novo documento Excel para OneDrive.

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>Usar um início rápido para criar um novo aplicativo Web Graph Microsoft

1. Vá para e siga as etapas para criar e abrir um exemplo de código de início rápido que [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) interage com Office serviços.
1. Na **etapa 1: Escolher idioma ou plataforma,** escolha **ASP.NET MVC**. Embora as etapas deste procedimento usem a opção ASP.NET MVC, as etapas seguem um padrão que se aplica a qualquer idioma ou plataforma.
1. Na **etapa 2: Obter uma ID do aplicativo e** um segredo, escolha Obter uma **ID do aplicativo e segredo**.
1. Entre na sua conta Microsoft 365 de usuário.  
1. Na página Da Web Secreta do **aplicativo,** salve o segredo do aplicativo em um local de arquivo onde você pode recuperá-lo e usá-lo mais tarde.
1. Escolha **Got it, take me back to the quick start**.
1. Na **etapa 2: Registro bem-sucedido!** Insira o segredo do aplicativo gerado.
1. Na **etapa 3: Iniciar a codificação,** escolha Baixar o exemplo de código baseado em **SDK.**
1. Extraia a pasta zip de download em uma pasta local.  
1. Abra o arquivo graph-tutorial.sln no Visual Studio 2019.
1. Crie e execute a solução e confirme se ela está funcionando corretamente. Você deve poder usar a página da Web do calendário para exibir seu calendário Microsoft 365 calendário.

### <a name="upload-a-file-to-onedrive"></a>Upload um arquivo para OneDrive

1. Abra a **solução graph-tutorial.sln** no Visual Studio 2019 e abra o arquivo **PrivateSettings.config.**

1. Adicione um novo escopo **Files.ReadWrite** à chave   **ida:AppScopes** para que ela se pareça com o código a seguir.

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. Abra o **arquivo Index.cshtml.**
1. Insira o código ActionLink a seguir para criar um botão para carregar um arquivo no OneDrive.

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. Abra o **arquivo HomeController.cs.**
1. Insira o código a seguir para manipular a solicitação do link de ação.

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. Abra o **arquivo GraphHelper.cs.**
1. Insira o código a seguir para chamar a API do Microsoft Graph para criar um novo arquivo no OneDrive.

    ```csharp
    public static async Task UploadFile(string fileName, System.IO.MemoryStream stream)
        {
           var graphClient = GetAuthenticatedClient();
            await graphClient.Me
                .Drive
                .Root
                .ItemWithPath(fileName)
                .Content
                .Request()
                .PutAsync<DriveItem>(stream);
            return;
        }
    ```

1. Pressione **F5** (ou **Depurar > Iniciar Depuração**). O aplicativo Web será iniciar.
1. Escolha **Clique aqui para entrar** e entrar.
1. Escolha **Clique aqui para criar um novo arquivo em OneDrive**.
1. Abra uma nova guia do navegador e entre na sua OneDrive de usuário. Você verá o arquivo test.txt na pasta raiz.

Agora que você aprendeu a carregar um arquivo no OneDrive, você pode reutilizar esse código para carregar qualquer documento Excel que você criar.

## <a name="additional-considerations-for-your-solution"></a>Considerações adicionais para sua solução

A solução de todos é diferente em termos de tecnologias e abordagens. As considerações a seguir ajudarão você a planejar como modificar sua solução para abrir documentos e incorporar seu Office Add-in.

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Criar uma nova Excel na página da Web

O exemplo modifica um documento Excel existente. Um cenário mais comum é que você criará uma nova planilha Excel de sua página da Web. Você pode encontrar detalhes adicionais sobre como criar uma nova planilha em **Criar um documento de** planilha fornecendo um nome de arquivo. Este artigo mostra como criar o arquivo localmente, mas você também pode criar o arquivo em um fluxo usando uma sobrecarga no método SpreadsheetDocument.Create.

### <a name="read-custom-properties-when-your-add-in-starts"></a>Ler propriedades personalizadas quando o seu complemento for iniciado

O exemplo de código armazena uma ID de trecho no novo documento Excel usando o SDK OOXML. Script Lab lê a ID de trecho do documento Excel e exibe o código do trecho quando ele é aberto. Talvez seja necessário enviar propriedades personalizadas para seu próprio complemento (como uma cadeia de caracteres de consulta ou um token de autenticação temporária).) Consulte **Persistindo o estado** e as configurações do add-in para obter detalhes completos sobre como ler propriedades personalizadas quando o seu complemento for iniciado.

### <a name="initialize-the-excel-document-with-data"></a>Inicializar o documento Excel com dados

Normalmente, quando o cliente abre um documento Excel de seu site, ele espera que o documento contenha alguns dados do site. Há algumas maneiras de gravar dados no documento.

- **Use o SDK OOXML para gravar os dados**. Você pode usar o SDK para gravar diretamente quaisquer dados no documento. Essa abordagem será útil se você quiser que os dados sejam disponibilizados no momento em que o documento for aberto.
- **Passe uma propriedade de consulta personalizada para seu Office Add-in**. Ao gerar o documento, você incorpora uma propriedade personalizada para o Office que contém uma cadeia de caracteres de consulta que recupera todos os dados necessários. Quando o seu add-in é aberto, ele recupera a consulta, executa a consulta e usa a API JS do Office para inserir o resultado da consulta no documento.

### <a name="working-with-the-ooxml-sdk"></a>Trabalhando com o SDK OOXML

O SDK OOXML é baseado em .NET. Se o aplicativo Web não for o .NET, você precisará procurar uma maneira alternativa de trabalhar com o OOXML.

Há uma versão JavaScript do SDK OOXML disponível no [Open XML SDK para JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).

Você pode colocar o código OOXML em uma função do Azure para separar o código .NET do restante do seu aplicativo Web. Em seguida, chame a função Azure (para gerar o documento Excel) do aplicativo Web. Para obter mais informações sobre as funções do Azure, consulte [Uma introdução às funções do Azure](/azure/azure-functions/functions-overview).

### <a name="use-single-sign-on"></a>Usar o login único

Para simplificar a autenticação, recomendamos que seu complemento implemente o login único. Para obter mais informações, consulte [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>Confira também

- [Bem-vindo ao SDK Open XML 2.5 para Office](/office/open-xml/open-xml-sdk)
- [Abrir automaticamente um painel de tarefas com um documento](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Persistir o estado e as configurações do suplemento](../develop/persisting-add-in-state-and-settings.md)
- [Criar um documento de planilha fornecendo um nome de arquivo](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)