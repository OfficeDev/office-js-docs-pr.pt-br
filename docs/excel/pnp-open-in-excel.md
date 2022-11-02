---
title: Abra o Excel na página da Web e insira seu Suplemento do Office
description: Abra o Excel na página da Web e insira seu Suplemento do Office.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 835518fb822602d6ca1af633f96d2be1e2697f44
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810341"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Abra o Excel na página da Web e insira seu Suplemento do Office

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Imagem do botão Excel em sua página da Web abrindo um novo documento do Excel com o suplemento inserido e a abertura automática.":::

Estenda seu aplicativo Web SaaS para que seus clientes possam abrir seus dados de uma página da Web diretamente para o Microsoft Excel. Um cenário comum é que os clientes trabalharão com dados em seu aplicativo Web. Em seguida, eles vão querer copiar os dados em um documento do Excel. Por exemplo, eles podem querer executar análises adicionais usando o Excel. Normalmente, o cliente é obrigado a exportar os dados para um arquivo, como um arquivo .csv e importar esses dados para o Excel. Eles também precisam adicionar manualmente seu Suplemento do Office ao documento.

Reduza o número de etapas para um único botão clique em sua página da Web que gera e abre o documento do Excel. Você também pode inserir seu Suplemento do Office dentro do documento e exibi-lo quando o documento for aberto. Isso garante que o cliente ainda tenha acesso aos recursos do aplicativo. Quando o documento é aberto, os dados que o cliente selecionou e seu Suplemento do Office já estão disponíveis para que eles continuem trabalhando.

Este artigo mostra código e técnicas para implementar esse cenário em seu próprio aplicativo Web SaaS.

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>Criar um novo documento do Excel e inserir um suplemento do Office

Primeiro, vamos aprender a criar um documento do Excel a partir de uma página da Web e inserir um suplemento no documento. O [exemplo de código de suplemento do Office OOXML](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) mostra como inserir o [suplemento Script Lab em](https://appsource.microsoft.com/product/office/wa104380862) um novo documento do Office. Embora o exemplo funcione com qualquer documento do Office, apenas nos concentraremos nas planilhas do Excel neste artigo. Use as etapas a seguir para compilar e executar o exemplo.

1. Extraia o código de exemplo de  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip em uma pasta em seu computador.
2. Para compilar e executar o exemplo, siga as etapas na seção **Para usar o projeto** do readme.
3. Quando você executar o exemplo, ele exibirá uma página da Web semelhante à captura de tela a seguir. Use a página da Web para criar um novo documento do Excel que contém Script Lab quando ele for aberto.
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Captura de tela da página da Web que o exemplo de laboratório de script inserido exibe para selecionar um arquivo do Excel e inserir o suplemento do laboratório de script nele.":::

### <a name="how-the-sample-works"></a>Como o exemplo funciona

O código de exemplo usa o SDK OOXML para inserir o suplemento Script Lab no documento do Excel escolhido. As informações a seguir são retiradas da [seção **Sobre o código**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) no arquivo readme.

O arquivo **Home.aspx.cs**:

- Fornece os manipuladores de eventos de botão e a manipulação básica da interface do usuário.
- Usa técnicas de ASP.NET padrão para carregar e baixar o arquivo.
- Usa a extensão de nome do arquivo carregado (xlsx, docx ou pptx) para determinar o tipo de arquivo. Isso precisa ser feito no início porque o SDK Open XML geralmente tem APIs distintas para cada tipo de arquivo.
- Chama o **OOXMLHelper** para validar o arquivo e chama para o **AddInEmbedder** para inserir Script Lab no arquivo e definir como aberto automaticamente.

O arquivo **AddInEmbedder.cs**:

- Fornece a lógica de negócios principal, que neste exemplo é um método que insira Script Lab.
- Faz chamadas para o auxiliar OOXML com base no tipo do arquivo.

O arquivo **OOXMLHelper.cs**:

- Fornece toda a manipulação detalhada do OOXML.
- Usa uma técnica padrão para validar o arquivo do Office, que é simplesmente chamar o método **Document.Open** nele. Se o arquivo for inválido, o método gerará uma exceção.
- Contém principalmente o código gerado pelas Ferramentas de Produtividade do SDK Open XML 2.5 que estão disponíveis no link para o [SDK Open XML 2.5](/office/open-xml/open-xml-sdk).

O método **GenerateWebExtensionPart1Content** no arquivo **OOXMLHelper.cs** define a referência à ID de Script Lab no Microsoft AppSource:

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- O valor **StoreType** é "OMEX", um alias para o Microsoft AppSource.
- O valor **da Loja** é "en-US" encontrado na seção cultura Microsoft AppSource para Script Lab.
- O valor **da ID** é a ID do ativo do Microsoft AppSource para Script Lab.

Se você estiver configurando um suplemento de um catálogo de compartilhamento de arquivos para abrir automaticamente, usará valores diferentes:

O valor **StoreType** é "FileSystem".

- O valor **store** é a URL do compartilhamento de rede; por exemplo, "\\\\MyComputer\\MySharedFolder". Essa deve ser a URL exata que aparece como o Endereço de Catálogo Confiável do compartilhamento no Office Trust Center.
- O valor **da ID** é a ID do aplicativo no manifesto de suplementos.
> [!NOTE]
> Para obter mais informações sobre valores alternativos para esses atributos, consulte [Abrir automaticamente um painel de tarefas com um documento](../develop/automatically-open-a-task-pane-with-a-document.md).

## <a name="use-the-fluent-ui"></a>Usar a interface do usuário fluente

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Ícones fluentes da interface do usuário para Word, Excel e PowerPoint.":::

Uma prática recomendada é usar a interface do usuário fluente para ajudar os usuários a fazer a transição entre os produtos da Microsoft. Você sempre deve usar um ícone do Office para indicar qual aplicativo do Office será iniciado a partir de sua página da Web. Vamos modificar o código de exemplo para usar o ícone do Excel para indicar que ele inicia o aplicativo excel.

1. Abra o exemplo no Visual Studio.
1. Abra a página **Home.aspx** .
1. Localize o código a seguir que é o botão de download no formulário.

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. Substitua o código do botão pela marca de imagem a seguir.

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. Pressione **F5** (ou **Depurar** > **Iniciar Depuração**). Você verá o ícone aparecer quando a home page for carregada.

Para obter mais informações, consulte [Ícones de Marca do Office](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) no portal do desenvolvedor de interface do usuário fluente.  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Carregar o documento do Excel no Microsoft OneDrive

Recomendamos carregar novos documentos no OneDrive se o cliente usar o OneDrive. Isso facilita a localização e o trabalho com os documentos. Vamos criar um novo exemplo de código e ver como você pode usar o SDK do Microsoft Graph para carregar um novo documento do Excel no OneDrive.

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>Usar um início rápido para criar um novo aplicativo Web do Microsoft Graph

1. [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) Acesse e siga as etapas para criar e abrir um exemplo de código de início rápido que interaja com os serviços do Office.
1. Na **etapa 1: Escolha seu idioma ou plataforma**, escolha **ASP.NET MVC**. Embora as etapas neste procedimento usem a opção ASP.NET MVC, as etapas seguem um padrão que se aplica a qualquer idioma ou plataforma.
1. Na **etapa 2: Obter uma ID do aplicativo e um segredo**, escolha **Obter uma ID do aplicativo e um segredo**.
1. Entre em sua conta do Microsoft 365.  
1. Na página Web **segredo do aplicativo** , salve o segredo do aplicativo em um local de arquivo em que você possa recuperá-lo e usá-lo posteriormente.
1. Escolha **Tenho, leve-me de volta para o início rápido**.
1. Na **etapa 2: Registro bem-sucedido!** Insira o segredo do aplicativo gerado.
1. Na **etapa 3: iniciar a codificação**, escolha **Baixar o exemplo de código baseado em SDK**.
1. Extraia a pasta zip de download em uma pasta local.  
1. Abra o arquivo graph-tutorial.sln no Visual Studio 2019.
1. Crie e execute a solução e confirme se ela está funcionando corretamente. Você deve poder usar a página da Web do calendário para exibir seu calendário do Microsoft 365.

### <a name="upload-a-file-to-onedrive"></a>Carregar um arquivo no OneDrive

1. Abra a solução **graph-tutorial.sln** no Visual Studio 2019 e abra o arquivo **PrivateSettings.config** .

1. Adicione um novo escopo **Files.ReadWrite** à chave **ida:AppScopes** para que ele se pareça com o código a seguir.

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. Abra o arquivo **Index.cshtml** .
1. Insira o código ActionLink a seguir para criar um botão para carregar um arquivo no OneDrive.

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. Abra o arquivo **HomeController.cs** .
1. Insira o código a seguir para manipular a solicitação no link de ação.

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
1. Insira o código a seguir para chamar o Microsoft API do Graph para criar um novo arquivo no OneDrive.

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

1. Pressione **F5** (ou **Depurar** > **Iniciar Depuração**). O aplicativo Web será iniciado.
1. Escolha **Clicar aqui para entrar** e entrar.
1. Escolha **Clicar aqui para criar um novo arquivo no OneDrive**.
1. Abra uma nova guia do navegador e entre em sua conta do OneDrive. Você verá o arquivo test.txt na pasta raiz.

Agora que você aprendeu a carregar um arquivo no OneDrive, você pode reutilizar esse código para carregar qualquer documento do Excel que você criar.

## <a name="additional-considerations-for-your-solution"></a>Considerações adicionais para sua solução

A solução de todos é diferente em termos de tecnologias e abordagens. As considerações a seguir ajudarão você a planejar como modificar sua solução para abrir documentos e inserir seu Suplemento do Office.

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Criar uma nova planilha do Excel na página da Web

O exemplo modifica um documento do Excel existente. Um cenário mais comum é que você criará uma nova planilha do Excel a partir de sua página da Web. Você pode encontrar detalhes adicionais sobre como criar uma nova planilha em **Criar um documento de planilha** fornecendo um nome de arquivo. Este artigo mostra como criar o arquivo localmente, mas você também pode criar o arquivo em um fluxo usando uma sobrecarga no método SpreadsheetDocument.Create.

### <a name="read-custom-properties-when-your-add-in-starts"></a>Ler propriedades personalizadas quando o suplemento for iniciado

O exemplo de código armazena uma ID de snippet no novo documento do Excel usando o SDK OOXML. Script Lab lê a ID do snippet do documento do Excel e exibe esse código de snippet quando ele é aberto. Talvez seja necessário enviar propriedades personalizadas para seu próprio suplemento (como uma cadeia de caracteres de consulta ou token de autenticação temporária).) Consulte **O estado e as configurações de suplemento persistentes** para obter detalhes completos sobre como ler propriedades personalizadas quando o suplemento for iniciado.

### <a name="initialize-the-excel-document-with-data"></a>Inicializar o documento do Excel com dados

Normalmente, quando o cliente abre um documento do Excel do seu site, ele espera que o documento contenha alguns dados do site. Há algumas maneiras de gravar dados no documento.

- **Use o SDK OOXML para gravar os dados**. Você pode usar o SDK para gravar diretamente todos os dados no documento. Essa abordagem será útil se você quiser que os dados estejam disponíveis no instante em que o documento for aberto.
- **Passe uma propriedade de consulta personalizada para seu Suplemento do Office**. Ao gerar o documento, você insira uma propriedade personalizada para o Suplemento do Office que contém uma cadeia de caracteres de consulta que recupera todos os dados necessários. Quando o suplemento é aberto, ele recupera a consulta, executa a consulta e usa a API JS do Office para inserir o resultado da consulta no documento.

### <a name="working-with-the-ooxml-sdk"></a>Trabalhando com o SDK do OOXML

O SDK OOXML é baseado em .NET. Se o aplicativo Web não fizer o .NET, você precisará procurar uma maneira alternativa de trabalhar com o OOXML.

Você pode colocar o código OOXML em uma função do Azure para separar o código .NET do restante do aplicativo Web. Em seguida, chame a função do Azure (para gerar o documento do Excel) do seu aplicativo Web. Para obter mais informações sobre funções do Azure, consulte [Uma introdução ao Azure Functions](/azure/azure-functions/functions-overview).

### <a name="use-single-sign-on"></a>Usar logon único

Para simplificar a autenticação, recomendamos que o suplemento implemente o logon único. Para obter mais informações, confira [Habilitar logon único para suplementos do Office](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>Confira também

- [Bem-vindo ao SDK do Open XML 2.5 para Office](/office/open-xml/open-xml-sdk)
- [Abrir automaticamente um painel de tarefas com um documento](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Persistir o estado e as configurações do suplemento](../develop/persisting-add-in-state-and-settings.md)
- [Criar um documento de planilha fornecendo um nome de arquivo](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)