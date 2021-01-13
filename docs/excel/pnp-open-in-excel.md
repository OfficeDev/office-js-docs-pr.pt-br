---
title: Abra o Excel na sua página da Web e insrir seu Complemento do Office
description: Abra o Excel na sua página da Web e insrir seu Complemento do Office.
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: a88cc647fc1dba8ab6e6ddc0b504aab96517026a
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839863"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Abra o Excel na sua página da Web e insrir seu Complemento do Office

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Imagem do botão do Excel em sua página da Web abrindo um novo documento do Excel com o seu complemento incorporado e abrindo automaticamente.":::

Estenda seu aplicativo Web SaaS para que os clientes possam abrir seus dados em uma página da Web diretamente para o Microsoft Excel. Um cenário comum é que os clientes trabalharão com dados em seu aplicativo Web. Em seguida, eles vão querer copiar os dados em um documento do Excel. Por exemplo, eles podem querer executar análises adicionais usando o Excel. Normalmente, o cliente é solicitado a exportar os dados para um arquivo, como um arquivo .csv, e, em seguida, importar esses dados para o Excel. Eles também precisam adicionar manualmente o seu Complemento do Office ao documento.

Reduza o número de etapas para um único clique em sua página da Web que gera e abre o documento do Excel. Você também pode inserir seu Complemento do Office dentro do documento e exibi-lo quando o documento for aberto. Isso garante que o cliente ainda tenha acesso aos recursos do aplicativo. Quando o documento é aberto, os dados que o cliente selecionou e o Seu Complemento do Office já estão disponíveis para que ele continue funcionando.

Este artigo mostra o código e as técnicas para implementar esse cenário em seu próprio aplicativo Web SaaS.

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>Criar um novo documento do Excel e incorporar um Complemento do Office

Primeiro, vamos aprender a criar um documento do Excel a partir de uma página da Web e inserir um complemento no documento. The [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the Script Lab [add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document. Embora o exemplo funcione com qualquer documento do Office, vamos nos concentrar apenas nas planilhas do Excel neste artigo. Use as etapas a seguir para criar e executar o exemplo.

1. Extraia o código de  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip exemplo de uma pasta em seu computador.
2. Para criar e executar o exemplo, siga as etapas na seção Para **usar o** projeto do leiame.
3. Quando você executar o exemplo, ele exibirá uma página da Web semelhante à captura de tela a seguir. Use a página da Web para criar um novo documento do Excel que contenha Script Lab quando ele for aberto.
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Screen shot of the web page that the embed script lab sample displays for selecting an Excel file and embedding the script lab add-in into it.":::

### <a name="how-the-sample-works"></a>Como funciona o exemplo

O código de exemplo usa o SDK OOXML para inserir o complemento Script Lab no documento do Excel que você escolher. As informações a seguir são retiradas da [ **seção Sobre o código**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) no arquivo leia-me.

O arquivo **Home.aspx.cs:**

- Fornece os manipuladores de eventos de botão e a manipulação básica da interface do usuário.
- Usa técnicas ASP.NET padrão para carregar e baixar o arquivo.
- Usa a extensão de nome de arquivo do arquivo carregado (xlsx, docx ou pptx) para determinar o tipo de arquivo. Isso precisa ser feito no início porque o SDK do Open XML geralmente tem APIs distintas para cada tipo de arquivo.
- Chama o **OOXMLHelper** para validar o arquivo e chama **AddInEmbedder** para inserir o Script Lab no arquivo e definir para abrir automaticamente.

O arquivo **AddInEmbedder.cs:**

- Fornece a lógica de negócios principal, que neste exemplo é um método que incorpora o Script Lab.
- Faz chamadas para o auxiliar OOXML com base no tipo do arquivo.

O arquivo **OOXMLHelper.cs:**

- Fornece toda a manipulação OOXML detalhada.
- Usa uma técnica padrão para validar o arquivo do Office, que é simplesmente para chamar o **método Document.Open** nele. Se o arquivo for inválido, o método lançará uma exceção.
- Contém principalmente o código que foi gerado pelas Ferramentas de Produtividade do SDK do Open XML 2.5 que estão disponíveis no link para o [SDK do Open XML 2.5.](/office/open-xml/open-xml-sdk)

O **método GenerateWebExtensionPart1Content** no arquivo **OOXMLHelper.cs** define a referência à ID do Script Lab no Microsoft AppSource:

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- O **valor de StoreType** é "OMEX", um alias do Microsoft AppSource.
- O **valor** da Loja é "en-US" encontrado na seção de cultura do Microsoft AppSource para o Script Lab.
- O **valor da Id** é a ID de ativo do Microsoft AppSource para o Script Lab.

Se você estiver configurando um complemento de um catálogo de compartilhamento de arquivos para abertura automática, usará valores diferentes:

O **valor de StoreType** é "FileSystem".

- O **valor** da Loja é a URL do compartilhamento de rede; por exemplo, " \\ \\ MyComputer \\ MySharedFolder". Essa deve ser a URL exata que aparece como o Endereço de Catálogo Confiável do compartilhamento na Central de Confiações do Office.
- O **valor da ID** é a ID do aplicativo no manifesto dos complementos.
> [!NOTE]
> Para obter mais informações sobre valores alternativos para esses atributos, consulte [Abrir automaticamente um painel de tarefas com um documento.](../develop/automatically-open-a-task-pane-with-a-document.md)

## <a name="use-the-fluent-ui"></a>Usar a interface do usuário do Fluent

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Ícones de interface do usuário do Fluent para Word, Excel e PowerPoint.":::

Uma prática prática é usar a interface do usuário do Fluent para ajudar os usuários na transição entre produtos da Microsoft. Você sempre deve usar um ícone do Office para indicar qual aplicativo do Office será lançado na sua página da Web. Vamos modificar o código de exemplo para usar o ícone do Excel para indicar que ele inicia o aplicativo Excel.

1. Abra o exemplo no Visual Studio.
1. Abra a **página Home.aspx.**
1. Encontre o seguinte código que é o botão de download no formulário:
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. Substitua o código do botão pela marca de imagem a seguir.
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. Pressione **F5** (ou **Depure > Iniciar Depuração).** Você verá o ícone aparecer quando a home page for carregada.

Para obter mais informações, consulte [Ícones de Marca do Office](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) no portal do desenvolvedor da interface do usuário do Fluent.  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Carregar o documento do Excel no Microsoft OneDrive

Recomendamos carregar novos documentos no OneDrive se o cliente usar o OneDrive. Isso torna mais fácil para eles encontrar e trabalhar com os documentos. Vamos criar um novo exemplo de código e ver como você pode usar o SDK do Microsoft Graph para carregar um novo documento do Excel no OneDrive.

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>Use um início rápido para criar um novo aplicativo Web do Microsoft Graph

1. Acesse e siga as etapas para criar e abrir um exemplo de código de início rápido que interage com os serviços do [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) Office 365.
1. In **step 1: Pick you language or platform**, choose ASP.NET **MVC**. Embora as etapas deste procedimento usem a ASP.NET MVC, as etapas seguem um padrão que se aplica a qualquer linguagem ou plataforma.
1. Na **etapa 2: Obter uma ID** e um segredo do aplicativo, escolha **Obter uma ID e um segredo do aplicativo.**
1. Entre em sua conta do Microsoft 365.  
1. Na página da Web Por favor, salve **o** segredo do aplicativo, salve o segredo do aplicativo em um local de arquivo onde você possa recuperar e usá-lo mais tarde.
1. Choose **Got it, take me back to the quick start**.
1. Na **etapa 2: Registro bem-sucedido!** Insira o segredo do aplicativo gerado.
1. In **step 3: Start coding**, choose **Download the SDK-based code sample**.
1. Extraia a pasta zip de download para uma pasta local.  
1. Abra o arquivo graph-tutorial.sln no Visual Studio 2019.
1. Crie e execute a solução e confirme se ela está funcionando corretamente. Você deve ser capaz de usar a página da Web de calendário para exibir seu calendário do Microsoft 365.

### <a name="upload-a-file-to-onedrive"></a>Carregar um arquivo no OneDrive

1. Abra a **solução graph-tutorial.sln** no Visual Studio 2019 e abra o arquivo **PrivateSettings.config** arquivo.
1. Adicione um novo escopo **Files.ReadWrite** à chave   **ida:AppScopes** para que ela se pareça com o seguinte código:
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
1. Abra o **HomeController.cs** arquivo.
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
1. Abra o **GraphHelper.cs** arquivo.
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
1. Pressione **F5** (ou **Depure > Iniciar Depuração).** O aplicativo Web será iniciar.
1. Escolha **Clique aqui para entrar** e entrar.
1. Escolha **Clique aqui para criar um novo arquivo no OneDrive.**
1. Abra uma nova guia do navegador e entre em sua conta do OneDrive. Você verá o arquivo test.txt na pasta raiz.

Agora que você aprendeu a carregar um arquivo no OneDrive, é possível reutilizar esse código para carregar qualquer documento do Excel criado.

## <a name="additional-considerations-for-your-solution"></a>Considerações adicionais para sua solução

A solução de todos é diferente em termos de tecnologias e abordagens. As considerações a seguir ajudarão você a planejar como modificar sua solução para abrir documentos e inserir seu Complemento do Office.

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Criar uma nova planilha do Excel na página da Web

O exemplo modifica um documento existente do Excel. Um cenário mais comum é que você criará uma nova planilha do Excel na página da Web. Você pode encontrar detalhes adicionais sobre como criar uma nova planilha em Criar um documento **de planilha** fornecendo um nome de arquivo. Este artigo mostra como criar o arquivo localmente, mas você também pode criar o arquivo em um fluxo usando uma sobrecarga no método SpreadsheetDocument.Create.

### <a name="read-custom-properties-when-your-add-in-starts"></a>Ler propriedades personalizadas quando o seu complemento for iniciado

O exemplo de código armazena uma ID de trecho no novo documento do Excel usando o SDK OOXML. O Script Lab lê a ID do trecho do documento do Excel e exibe esse código de trecho quando ele é aberto. Talvez seja necessário enviar propriedades personalizadas para seu próprio add-in (como uma cadeia de caracteres de consulta ou um token de autenticação temporária).) Consulte **o estado e as configurações persistentes** do complemento para obter detalhes completos sobre como ler as propriedades personalizadas quando o seu complemento é iniciado.

### <a name="initialize-the-excel-document-with-data"></a>Inicializar o documento do Excel com dados

Normalmente, quando o cliente abre um documento do Excel a partir do seu site, ele espera que o documento contenha alguns dados do site. Há algumas maneiras de gravar dados no documento.

- **Use o SDK OOXML para gravar os dados.** Você pode usar o SDK para gravar diretamente quaisquer dados no documento. Essa abordagem será útil se você quiser que os dados sejam disponibilizados assim que o documento for aberto.
- **Passe uma propriedade de consulta personalizada para o seu complemento do Office.** Ao gerar o documento, você incorpora uma propriedade personalizada para o complemento do Office que contém uma cadeia de caracteres de consulta que recupera todos os dados necessários. Quando o seu complemento é aberto, ele recupera a consulta, executa a consulta e usa a API JS do Office para inserir o resultado da consulta no documento.

### <a name="working-with-the-ooxml-sdk"></a>Trabalhando com o OOXML SDK

O SDK OOXML é baseado em .NET. Se seu aplicativo Web não .NET, você precisará procurar uma maneira alternativa de trabalhar com OOXML.

Há uma versão JavaScript do SDK OOXML disponível no [SDK do Open XML para JavaScript.](https://archive.codeplex.com/?p=openxmlsdkjs)

Você pode colocar o código OOXML em uma função do Azure para separar o código .NET do restante do seu aplicativo Web. Em seguida, chame a função do Azure (para gerar o documento do Excel) a partir do aplicativo Web. Para saber mais sobre as funções do Azure, confira [Uma introdução às funções do Azure.](/azure/azure-functions/functions-overview)

### <a name="use-single-sign-on"></a>Usar o single sign-on

Para simplificar a autenticação, recomendamos que seu complemento implemente o single sign-on. Para saber mais, confira [Habilitar o single sign-on para Os Complementos do Office](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>Confira também

- [Bem-vindo ao SDK 2.5 do Open XML para Office](/office/open-xml/open-xml-sdk)
- [Abrir automaticamente um painel de tarefas com um documento](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Persistir o estado e as configurações do suplemento](../develop/persisting-add-in-state-and-settings.md)
- [Criar um documento de planilha fornecendo um nome de arquivo](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)