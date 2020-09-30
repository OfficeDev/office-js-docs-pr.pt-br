---
title: Abrir o Excel a partir da sua página da Web e incorporar o suplemento do Office
description: Abra o Excel a partir da sua página da Web e insira seu suplemento do Office.
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: 00846ca5ca05e65fd75629f5aad0e4fb3d947ab1
ms.sourcegitcommit: 42202d7e2ac24dffa77cf937f5697a1cd79ee790
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/30/2020
ms.locfileid: "48308541"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Abrir o Excel a partir da sua página da Web e incorporar o suplemento do Office

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Imagem do botão do Excel na página da Web que está abrindo um novo documento do Excel com seu suplemento incorporado e de abertura automática.":::

Estenda o aplicativo Web SaaS para que seus clientes possam abrir seus dados de uma página da Web diretamente para o Microsoft Excel. Um cenário comum é que os clientes trabalhem com dados no seu aplicativo Web. Em seguida, eles deverão copiar os dados em um documento do Excel. Por exemplo, eles podem desejar executar análise adicional usando o Excel. Normalmente, o cliente precisa exportar os dados para um arquivo, como um arquivo. csv, e importá-los para o Excel. Eles também precisam adicionar manualmente o suplemento do Office ao documento.

Reduza o número de etapas para um único clique de botão na página da Web que gera e abre o documento do Excel. Você também pode inserir seu suplemento do Office dentro do documento e exibi-lo quando o documento é aberto. Isso garante que o cliente ainda tenha acesso aos recursos do aplicativo. Quando o documento é aberto, os dados selecionados pelo cliente e seu suplemento do Office já estão disponíveis para continuar trabalhando.

Este artigo mostra códigos e técnicas para implementar esse cenário em seu próprio aplicativo Web SaaS.

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>Criar um novo documento do Excel e incorporar um suplemento do Office

Primeiro, vamos aprender como criar um documento do Excel a partir de uma página da Web e incorporar um suplemento no documento. O [exemplo de código do suplemento embed do Office OOXML](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) mostra como incorporar o [suplemento do laboratório de script](https://appsource.microsoft.com/product/office/wa104380862) em um novo documento do Office. Embora o exemplo funcione com qualquer documento do Office, vamos nos concentrar em planilhas do Excel neste artigo. Use as etapas a seguir para criar e executar o exemplo.

1. Extraia o código de exemplo de  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip uma pasta no seu computador.
2. Para criar e executar o exemplo, siga as etapas na seção **para usar o projeto** do Leiame.
3. Quando você executar o exemplo, será exibida uma página da Web semelhante à captura de tela a seguir. Use a página da Web para criar um novo documento do Excel que contém o laboratório de script quando ele é aberto.
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Imagem do botão do Excel na página da Web que está abrindo um novo documento do Excel com seu suplemento incorporado e de abertura automática.":::

### <a name="how-the-sample-works"></a>Como funciona o exemplo

O código de exemplo usa o SDK do OOXML para inserir o suplemento de laboratório de script no documento do Excel que você escolher. As informações a seguir são obtidas da [seção **sobre o código** ](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) no arquivo Leiame.

O arquivo **Home.aspx.cs**:

- Fornece os manipuladores de eventos de botão e a manipulação básica da interface do usuário.
- Usa as técnicas de ASP.NET padrão para carregar e baixar o arquivo.
- Usa a extensão de nome de arquivo do arquivo carregado (xlsx, docx ou pptx) para determinar o tipo de arquivo. Isso precisa ser feito no início porque o SDK do Open XML geralmente tem APIs distintas para cada tipo de arquivo.
- Chamadas para o **OOXMLHelper** para validar o arquivo e as chamadas para o **AddInEmbedder** para inserir o laboratório de script no arquivo e definidas para abrir automaticamente.

O arquivo **AddInEmbedder.cs**:

- Fornece a principal lógica de negócios, que neste exemplo é um método que incorpora o script Lab.
- Faz chamadas para o auxiliar OOXML com base no tipo de arquivo.

O arquivo **OOXMLHelper.cs**:

- Fornece toda a manipulação detalhada do OOXML.
- Usa uma técnica padrão para validar o arquivo do Office, que é simplesmente chamar o método **Document. Open** . Se o arquivo for inválido, o método gera uma exceção.
- Contém principalmente o código que foi gerado pelas ferramentas de produtividade do SDK do Open XML 2,5 que estão disponíveis no link para o [Open xml 2,5 SDK](/office/open-xml/open-xml-sdk).

O método **GenerateWebExtensionPart1Content** no arquivo **OOXMLHelper.cs** define a referência à ID do laboratório de script no Microsoft AppSource:

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- O valor **storetype** é "OMEX", um alias para o Microsoft AppSource.
- O valor da **loja** é "en-US", encontrado na seção Microsoft AppSource Culture for script Lab.
- O valor de **ID** é a ID de ativo do Microsoft AppSource para o laboratório de scripts.

Se você estiver configurando um suplemento de um catálogo de compartilhamento de arquivos para abrir automaticamente, você usará valores diferentes:

O valor **storetype** é "FileSystem".

- O valor da **loja** é a URL do compartilhamento de rede; por exemplo, " \\ \\ \\ myMySharedFolder". Esta deve ser a URL exata que aparece como o endereço do catálogo confiável do compartilhamento na central de confiabilidade do Office.
- O valor de **ID** é a ID do aplicativo no manifesto de suplementos.
> [!NOTE]
> Para obter mais informações sobre valores alternativos para esses atributos, consulte [abrir automaticamente um painel de tarefas com um documento](../develop/automatically-open-a-task-pane-with-a-document.md).

## <a name="use-the-fluent-ui"></a>Usar a interface do usuário fluente

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Imagem do botão do Excel na página da Web que está abrindo um novo documento do Excel com seu suplemento incorporado e de abertura automática.":::

Uma prática recomendada é usar a interface do usuário fluente para ajudar os usuários a fazer a transição entre os produtos da Microsoft. Você deve sempre usar um ícone do Office para indicar qual aplicativo do Office será iniciado na sua página da Web. Vamos modificar o código de exemplo para usar o ícone do Excel para indicar que ele inicia o aplicativo Excel.

1. Abra o exemplo no Visual Studio.
1. Abra a página **Home. aspx** .
1. Localize o código a seguir que é o botão baixar no formulário:
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. Substitua o código do botão pela marca de imagem a seguir.
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. Pressione **F5** (ou **debug > iniciar a depuração**). Você verá o ícone aparecerá quando a home page for carregada.

Para obter mais informações, consulte [ícones da marca do Office](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) no portal do desenvolvedor da interface do usuário Fluent.  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Carregar o documento do Excel para o Microsoft OneDrive

Recomendamos carregar novos documentos para o OneDrive se o cliente usar o OneDrive. Isso facilita a localização e o trabalho com os documentos. Vamos criar um novo exemplo de código e ver como você pode usar o SDK do Microsoft Graph para carregar um novo documento do Excel para o OneDrive.

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>Usar um início rápido para criar um novo aplicativo Web do Microsoft Graph

1. Vá para [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) e siga as etapas para criar e abrir um exemplo de código de início rápido que interaja com os serviços do Office 365.
1. Na **etapa 1: escolha o idioma ou a plataforma**, escolha **ASP.NET MVC**. Embora as etapas neste procedimento usem a opção MVC ASP.NET, as etapas seguem um padrão que se aplica a qualquer idioma ou plataforma.
1. Na **etapa 2: obter uma ID de aplicativo e um segredo**, escolha **obter uma ID de aplicativo e segredo**.
1. Entre em sua conta do Microsoft 365.  
1. Na página **salvar seu segredo de aplicativo** , salve o segredo do aplicativo em um local de arquivo onde você possa recuperá-lo e usá-lo mais tarde.
1. Escolha se **o fez, volte para o início rápido**.
1. Na **etapa 2: registro bem-sucedido!** Insira o segredo do aplicativo gerado.
1. Na **etapa 3: iniciar a codificação**, escolha **baixar o exemplo de código baseado em SDK**.
1. Extraia a pasta zip de download em uma pasta local.  
1. Abra o arquivo Graph-tutorial. sln no Visual Studio 2019.
1. Criar e executar a solução e confirmar se ela está funcionando corretamente. Você deve ser capaz de usar a página da Web de calendário para exibir seu calendário do Microsoft 365.

### <a name="upload-a-file-to-onedrive"></a>Carregar um arquivo para o OneDrive

1. Abra a solução **Graph-tutorial. sln** no Visual Studio 2019 e abra o arquivo **PrivateSettings.config** .
1. Adicione um novo escopo **files. ReadWrite**   à chave **ida: AppScopes** para que se pareça com o seguinte código:
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. Abra o arquivo **index. cshtml** .
1. Insira o seguinte código ActionLink para criar um botão para carregar um arquivo para o OneDrive.
    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```
1. Abra o arquivo **HomeController.cs** .
1. Insira o código a seguir para lidar com a solicitação do link de ação.
    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```
1. Abra o arquivo **GraphHelper.cs** .
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
1. Pressione **F5** (ou **debug > iniciar a depuração**). O aplicativo Web será iniciado.
1. Escolha **clique aqui para entrar**e entrar.
1. Escolha **clique aqui para criar um novo arquivo no onedrive**.
1. Abra uma nova guia do navegador e entre em sua conta do OneDrive. Você verá o arquivo test.txt na pasta raiz.

Agora que você aprendeu como carregar um arquivo para o OneDrive, você pode reutilizar esse código para carregar qualquer documento do Excel que você criar.

## <a name="additional-considerations-for-your-solution"></a>Considerações adicionais para sua solução

A solução de todos é diferente em termos de tecnologias e abordagens. As considerações a seguir ajudarão você a planejar como modificar sua solução para abrir documentos e incorporar o suplemento do Office.

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Criar uma nova planilha do Excel a partir da página da Web

O exemplo modifica um documento existente do Excel. Um cenário mais comum é criar uma nova planilha do Excel a partir da sua página da Web. Você pode encontrar mais detalhes sobre como criar uma nova planilha em **criar um documento de planilha** fornecendo um nome de arquivo. Este artigo mostra como criar o arquivo localmente, mas você também pode criar o arquivo em um Stream usando uma sobrecarga no método SpreadsheetDocument. Create.

### <a name="read-custom-properties-when-your-add-in-starts"></a>Ler as propriedades personalizadas quando o suplemento for iniciado

O exemplo de código armazena uma ID de trecho no novo documento do Excel usando o SDK do OOXML. O script Lab lê o ID de trecho do documento do Excel e, em seguida, exibe esse código de trecho quando ele é aberto. Talvez você precise enviar propriedades personalizadas para seu próprio suplemento (como uma cadeia de caracteres de consulta ou um token de autenticação temporário). Confira o **estado e as configurações do suplemento persistentes** para obter detalhes completos sobre como ler as propriedades personalizadas quando o suplemento for iniciado.

### <a name="initialize-the-excel-document-with-data"></a>Inicializar o documento do Excel com dados

Normalmente, quando o cliente abre um documento do Excel do seu site da Web, ele espera que o documento contenha alguns dados do site. Há algumas maneiras de gravar dados no documento.

- **Use o SDK do OOXML para gravar os dados**. Você pode usar o SDK para gravar dados diretamente no documento. Essa abordagem é útil se você deseja que os dados estejam disponíveis para o momento em que o documento é aberto.
- **Passe uma propriedade de consulta personalizada para o suplemento do Office**. Ao gerar o documento, você incorpora uma propriedade personalizada para o suplemento do Office que contém uma cadeia de caracteres de consulta que recupera todos os dados necessários. Quando o suplemento é aberto, ele recupera a consulta, executa a consulta e usa a API do Office JS para inserir o resultado da consulta no documento.

### <a name="working-with-the-ooxml-sdk"></a>Trabalhar com o SDK do OOXML

O SDK do OOXML é baseado no .NET. Se o seu aplicativo Web não .NET, você precisará procurar por uma maneira alternativa de trabalhar com o OOXML.

Há uma versão JavaScript do SDK do OOXML disponível no [Open XML SDK para JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).

Você pode colocar o código OOXML em uma função do Azure para separar o código .NET do restante do seu aplicativo Web. Em seguida, chame a função do Azure (para gerar o documento do Excel) a partir do seu aplicativo Web. Para obter mais informações sobre as funções do Azure, consulte [introdução às funções do Azure](https://docs.microsoft.com/azure/azure-functions/functions-overview).

### <a name="use-single-sign-on"></a>Usar logon único

Para simplificar a autenticação, recomendamos que seu suplemento implemente o logon único. Para obter mais informações, consulte [habilitar o logon único para suplementos do Office](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>Confira também

- [Bem-vindo ao Open XML SDK 2,5 para Office](/office/open-xml/open-xml-sdk)
- [Abrir automaticamente um painel de tarefas com um documento](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Persistir o estado e as configurações do suplemento](../develop/persisting-add-in-state-and-settings.md)
- [Criar um documento de planilha fornecendo um nome de arquivo](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)
