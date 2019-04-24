---
title: Criar e depurar suplementos do Office no Visual Studio
description: Use o Visual Studio para criar e depurar suplementos do Office na área de trabalho do cliente Office para Windows
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 74a1430482b507d04f1be60683242fd9ae4a4393
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449903"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Criar e depurar suplementos do Office no Visual Studio

Este artigo descreve como usar o Visual Studio 2017 para criar um suplemento do Office para Excel, Word, PowerPoint ou Outlook e depurar suplemento na área de trabalho do cliente Office no Windows. Se você estiver usando outra versão do Visual Studio, os procedimentos poderão variar um pouco.

> [!NOTE]
> O Visual Studio não suporta a criação de suplementos do Office para o OneNote ou o Project, mas você pode usar o [Yeoman gerador de suplementos do Office](https://github.com/OfficeDev/generator-office) para criar esses tipos de suplementos.
> - Para começar a usar um suplemento do OneNote, confira o artigo [Crie seu primeiro suplemento do OneNote](../quickstarts/onenote-quickstart.md).
>
> - Para começar a usar um suplemento do Project, confira o artigo [Crie seu primeiro suplemento do Project](../quickstarts/project-quickstart.md).

## <a name="prerequisites"></a>Pré-requisitos

- [Visual Studio 2017](https://www.visualstudio.com/vs/) com a carga de trabalho de **desenvolvimento do Office/SharePoint** instalada

    > [!TIP]
    > Se você já instalou o Visual Studio 2017, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Office/SharePoint** seja instalada. Se essa carga de trabalho ainda não estiver instalada, use o instalador Visual Studio para [instalá-la](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).

- Office 2013 ou posterior

    > [!TIP]
    > Se você não tiver o Office, você pode participar do[programa de desenvolvedor do Office 365](https://developer.microsoft.com/office/dev-program) para obter uma assinatura do Office 365 ou [Inscreva-se para uma avaliação gratuita de 1 mês](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).

## <a name="create-the-add-in-project-in-visual-studio"></a>Criar um projeto de suplemento no Visual Studio

Inicie realizando estas três etapas e, em seguida, conclua as etapas na seção a seguir que corresponde ao tipo de suplemento que você está criando. 

1. Na barra de menus do Visual Studio, selecione **Arquivo** > **Novo**  >  **Projeto**.

2. Na lista de tipos de projeto em **Visual C#** ou no **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e escolha o tipo de projeto que você deseja criar. 

3. Dê um nome ao projeto e escolha **OK**.

### <a name="word-web-add-in-or-outlook-web-add-in"></a>Suplemento do Word na Web ou suplemento do Outlook na Web

Se você optou por criar um **Suplemento Word na Web** ou uma **Suplemento do Outlook Web**O Visual Studio cria uma solução e os dois projetos aparecem no**Explorador de soluções**. Em seguida, você pode [explorar a solução do Visual Studio](#explore-the-visual-studio-solution). 

### <a name="powerpoint-web-add-in"></a>Suplemento do PowerPoint Web 

Se você optou por criar um **suplemento PowerPoint Web**, a caixa de diálogo**Criar Suplemento do Office** aparece. 

- Para criar um suplemento no painel tarefas, selecione **Adicionar novas funcionalidades para o PowerPoint** e, em seguida, escolha o botão **Concluir** para criar a  solução no Visual Studio.

- Para criar um suplemento de conteúdo, selecione **Inserir conteúdo nos slides do PowerPoint** e, em seguida, escolha o botão **Concluir** para criar a solução no Visual Studio.

Em seguida, você pode [explorar a solução do Visual Studio](#explore-the-visual-studio-solution).

### <a name="excel-web-add-in"></a>Suplemento do Excel Web

Se você optou por criar um **suplemento Excel Web**, a caixa de diálogo**Criar Suplemento do Office** aparece. 

- Para criar um suplemento no painel tarefas, selecione **Adicionar novas funcionalidades para o Excel** e, em seguida, escolha o botão **Concluir** para criar a solução no Visual Studio.

- Para criar um suplemento de conteúdo, selecione **Inserir conteúdo em planilhas do Excel**, escolha o botão **Próximo**, selecione uma das seguintes opções e, em seguida, escolha o botão **Concluir** para criar a solução no Visual Studio:

    - **Suplemento Básico** – criar um projeto de suplemento de conteúdo com o código inicial mínimo

    - **Documento visualização no** – criar um projeto de suplemento de conteúdo com o código inicial para visualizar e vincular a dados  

### <a name="explore-the-visual-studio-solution"></a>Explorar a solução do Visual Studio

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

## <a name="modify-your-add-in-settings"></a>Modificar as configurações de suplemento

Para alterar as configurações no seu suplemento, edite o arquivo de manifesto XML no suplemento do projeto. No **Gerenciador de Soluções**, expanda o nó de projeto do suplemento, expanda a pasta que contém o manifesto XML e escolha o manifesto XML. Você pode apontar para qualquer elemento do arquivo para exibir uma dica de ferramenta que descreve a finalidade do elemento. Para saber mais sobre o arquivo de manifesto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).

## <a name="develop-the-contents-of-your-add-in"></a>Desenvolver o conteúdo do suplemento

Enquanto o projeto de suplemento permite modificar as configurações que descrevem o suplemento, o aplicativo Web fornece o conteúdo que aparece no suplemento. 

O projeto de aplicativo Web contém um arquivo HTML padrão e o arquivo JavaScript e arquivo CSS que você pode usar para começar. Alguns desses arquivos contêm referências a outras bibliotecas JavaScript, incluindo a API JavaScript para Office. Você pode desenvolver o suplemento para atualizar esses arquivos e/ou adicionar mais arquivos HTML e JavaScript. A tabela a seguir descreve os arquivos padrão que o projeto de aplicativo web contém quando a solução Visual Studio é criada.

|**Nome do arquivo**|**Descrição**|
|:-----|:-----|
|**Home.html**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.html**<br/>(Outlook)|The default HTML page of the add-in. Essa página é exibida como a primeira no suplemento quando ele é ativado em um documento, mensagem de email ou item de compromisso. Esse arquivo contém todas as referências de arquivo de que você precisa para começar. Você pode começar a desenvolver o suplemento, adicionando o código HTML para esse arquivo.|
|**Home.js**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.js**<br/>(Outlook)|O arquivo JavaScript associado a página **Home.html** (Excel, PowerPoint, Word) ou página **MessageRead.html** (Outlook). Esse arquivo deve conter qualquer código específico para o comportamento da página **Home.html** (Excel, PowerPoint, Word) ou página **MessageRead.html**(Outlook). Esse contém código de exemplo para você começar.|
|**Home.CSS**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.css**<br/>(Outlook)|Define os estilos padrão para aplicar ao suplemento. É recomendável usar a estrutura da interface do usuário do Office para design e estilos. Para saber mais, confira [Office UI Fabric em suplementos do Office](../design/office-ui-fabric.md).|

> [!NOTE]
> Não é necessário usar esses arquivos. Fique à vontade para adicionar outros arquivos ao projeto e usá-los em vez disso. Se desejar que outro arquivo HTML apareça como a página inicial do suplemento, abra o editor de manifesto e defina a propriedade **SourceLocation** para o nome do arquivo.

## <a name="debug-your-add-in"></a>Depurar o suplemento

Você pode usar o Visual Studio para depurar seu suplemento no cliente da área de trabalho do Office no Windows, conforme descrito nas seções a seguir:

- [Revise as propriedades de build e depuração](#review-the-build-and-debug-properties)
- [Usar um documento existente para depurar o suplemento](#use-an-existing-document-to-debug-the-add-in)
- [Iniciar o projeto](#start-the-project)
- [Depurar o código de um suplemento Excel, PowerPoint ou Word](#debug-the-code-for-an-excel-powerpoint-or-word-add-in)
- [Depurar o código de um suplemento do Outlook](#debug-the-code-for-an-outlook-add-in)

> [!NOTE]
> É possível usar o Visual Studio para depurar suplementos do Office no Office Online ou Office para Mac. Confira informações sobre a depuração nessas plataformas [Depurar Suplementos do Office no Office Online](../testing/debug-add-ins-in-office-online.md) ou [Depurar Suplementos do Office no iPad e no Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)

### <a name="review-the-build-and-debug-properties"></a>Examinar as propriedades de compilação e depuração

Antes de começar a depuração, examine as propriedades de cada projeto para confirmar se o Visual Studio abrirá o aplicativo do host desejado e se as propriedades de compilação e depuração propriedades estão configuradas adequadamente.

#### <a name="add-in-project-properties"></a>Propriedades do projeto de suplemento

Abrir a janela **Propriedades**para o projeto de suplemento revisar as propriedades do projeto:

1. No **Explorador de soluções** Escolha o projeto do suplemento (*não* o projeto do aplicativo Web).

2. Na barra de menus, escolha **Exibir** >  **Janela Propriedades**.

A tabela a seguir descreve as propriedades do projeto.

|**Property**|**Descrição**|
|:-----|:-----|
|**Iniciar Ação**|Especifica o modo de depuração do suplemento. Atualmente, só **cliente de área de trabalho do Office** modo tem suporte para projetos de suplementos do Office.|
|**Iniciar documento**<br/> (apenas suplementos Excel, PowerPoint e Word)|Especifica o documento a ser aberto quando você iniciar o projeto.|
|**Projeto da Web**|Especifica o nome do projeto Web associado ao suplemento.|
|**Email Address**<br/>(Apenas suplementos do Outlook)|Especifica o endereço de email da conta de usuário no Exchange Server ou no Exchange Online que você quer usar para testar o suplemento do Outlook.|
|**EWS Url**<br/>(Apenas suplementos do Outlook)|URL do serviço Web do Exchange (por exemplo: `https://www.contoso.com/ews/exchange.aspx`). |
|**OWA Url**<br/>(Apenas suplementos do Outlook)|URL do Outlook Web App (Por exemplo: `https://www.contoso.com/owa`).|
|**Usar autenticação multifator**<br/>(Apenas suplementos do Outlook)|Valor Booleano que indica se a autenticação multifator deve ser utilizada.|
|**Nome de Usuário**<br/>(Apenas suplementos do Outlook)|Especifica o nome da conta de usuário no Exchange Server ou no Exchange Online com a qual você deseja testar o suplemento do Outlook.|
|**Arquivo do projeto**|Especifica o nome do arquivo que contém informações de compilação, configuração e outras informações sobre o projeto.|
|**Pasta do projeto**|O local do arquivo do projeto.|

> [!NOTE]
> Para um suplemento do Outlook, você pode optar por especificar valores para uma ou mais das propriedades *Apenas suplemento Outlook* na janela**Propriedades** mas isso não é necessário.

#### <a name="web-application-project-properties"></a>Propriedades do projeto de aplicativo Web

Abrir a janela **Propriedades**para o projeto de aplicativo Web para revisar as propriedades do projeto:

1. No **Explorador de soluções** Escolha o projeto o projeto do aplicativo Web.

2. Na barra de menus, escolha **Exibir** >  **Janela Propriedades**.

A tabela a seguir descreve as propriedades do projeto de aplicativo web que são mais relevantes para projetos de suplementos do Office.

|**Property**|**Descrição**|
|:-----|:-----|
|**SSL habilitado**|Especifica se o SSL está habilitado no site. Essa propriedade deve ser definida como **Verdadeira** para projetos de suplementos do Office.|
|**URL SSL**|Especifica a URL HTTPS segura para o site. Somente leitura.|
|**URL**|Especifica a URL HTTP para o site. Somente leitura.|
|**Arquivo do projeto**|Especifica o nome do arquivo que contém informações de compilação, configuração e outras informações sobre o projeto.|
|**Pasta do projeto**|Especifica o local do arquivo do projeto. Somente leitura. O arquivo de manifesto do Visual Studio gerado no tempo de execução é escrito para a pasta `bin\Debug\OfficeAppManifests` neste local.|

### <a name="use-an-existing-document-to-debug-the-add-in"></a>Usar um documento existente para depurar o suplemento

Se você tiver um documento que contém os dados de teste deseja usar ao depurar seu suplemento o Excel, PowerPoint ou Word, o Visual Studio pode ser configurado para abrir esse documento quando você iniciar o projeto. Para especificar um documento existente a ser usado durante a depuração do complemento, execute as etapas a seguir.

1. No **Explorador de soluções** Escolha o projeto do suplemento (*não* o projeto do aplicativo Web).

2. Na barra de menus, escolha **Projeto** > **Adicionar Item Existente**.

3. Na caixa de diálogo **Adicionar Item Existente**, localize e selecione o documento que você deseja adicionar.

4. Escolha o botão **Adicionar** para adicionar o documento ao projeto.

5. No **Explorador de soluções** Escolha o projeto do suplemento (*não* o projeto do aplicativo Web).

6. Na barra de menus, escolha **Exibir**,  > **Janela Propriedades**.

7. Na janela **Propriedades**, escolha a lista **Iniciar Documento** e selecione o documento que você adicionou ao projeto. O projeto agora está configurado para iniciar o suplemento nesse documento.

### <a name="start-the-project"></a>Iniciar o projeto

Iniciar o projeto escolhendo **Depurar** > **Iniciar Depuração** na barra do menu. O Visual Studio compilará automaticamente a inicie o Office para hospedar o suplemento.

> [!NOTE]
> Quando você inicia um projeto de um suplemento do Outlook, você será solicitado a inserir as credenciais de logon. Se você for solicitado a fazer logon repetidamente ou se receber um erro informando que você não está autorizado, a Autenticação Básica pode estar desabilitada para contas em seu locatário do Office 365. Nesse caso, tente usar uma conta da Microsoft. Você também pode precisar definir a propriedade "Usar autenticação multifator" como Verdadeiro na caixa de diálogo de propriedades do projeto de suplemento do Outlook na Web.

Quando o Visual Studio compila o projeto ele executa as seguintes tarefas:

1. Cria uma cópia do arquivo de manifesto XML e a adiciona ao diretório `_ProjectName_\bin\Debug\OfficeAppManifests`. O aplicativo host consome esta cópia quando você inicia o Visual Studio e depura o suplemento.

2. Cria um conjunto de entradas de registro no computador que habilitam o suplemento a aparecer no aplicativo host.

3. Compila o projeto de aplicativo Web e o implanta no servidor Web IIS local(https://localhost).

4. Se este for o primeiro projeto de suplemento implantado no servidor Web do IIS local, talvez seja solicitado que você instale um Certificado Autoassinado no repositório de Certificado Raiz Confiável do usuário atual. Isso é necessário para que o IIS Express exiba o conteúdo do seu suplemento corretamente.


> [!NOTE]
> A versão mais recente do Office pode usar um controle da Web mais recente para exibir o conteúdo do suplemento ao ser executado no Windows 10. Se este for o caso, o Visual Studio pode solicitar que você adicione uma isenção de loopback de rede local. Isso é necessário para que o controle da Web, no aplicativo host do Office, possa acessar o site implantado no servidor Web do IIS local. Você também pode alterar essa configuração a qualquer momento no Visual Studio, em **Ferramentas** > **Opções** > **Ferramentas do Office (Web)** > **Depuração do Suplemento da Web**.


Depois, o Visual Studio faz o seguinte:

1. Modifica o elemento [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) do arquivo de manifesto XML, substituindo o token `~remoteAppUrl` pelo endereço totalmente qualificado da página inicial (por exemplo,`https://localhost:44302/Home.html` ).

2. Inicia o projeto de aplicativo Web no IIS Express.

3. Abre o aplicativo host.

O Visual Studio não mostra erros de validação na janela **OUTPUT** ao compilar o projeto. O Visual Studio relata erros e avisos na janela **ERRORLIST** à medida que eles ocorrem. O Visual Studio também relata erros de validação mostrando sublinhados ondulados (conhecidos como rabiscos) de cores diferentes no editor de código e texto. Essas marcas o notificam de problemas que o Visual Studio detectou no código. Para saber mais, confira [Editor de código e texto](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). Para saber mais sobre como habilitar ou desabilitar a validação, confira: [Opções, Editor de texto, JavaScript, IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2017).

Para examinar as regras de validação do arquivo de manifesto XML no projeto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).

### <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a>Depurar o código de um suplemento Excel, PowerPoint ou Word

Se o suplemento não está visível no documento que é exibido no aplicativo host (Excel, PowerPoint ou Word) após você[iniciar o projeto](#start-the-project), inicie manualmente o suplemento no aplicativo do host. Por exemplo, inicie o suplemento do painel de tarefas, escolhendo o **Mostrar painel de tarefas** botão na faixa de opções da guia **Home**. Depois do suplemento ser exibido no Excel, PowerPoint ou Word, você pode depurar seu código fazendo o seguinte:

1. No Excel, PowerPoint ou Word, escolha o **Inserir** pressione tab e, em seguida, escolha a seta para baixo à direita de **Meus suplementos**.

    ![Inserir faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)

2. Na lista de suplementos disponíveis, localize a seção**suplementos do desenvolvedor** e selecione o seu suplemento para registrar.

3. No Visual Studio, defina pontos de interrupção no seu código.

4. No Excel, PowerPoint ou Word, interaja com o suplemento.

5. Como os pontos de interrupção são atingidos no Visual Studio, percorra o código conforme necessário.

Você pode alterar o código e examinar os efeitos das alterações no suplemento sem ter que fechar o aplicativo host e reiniciar o projeto. Depois de salvar o código, simplesmente recarregue o suplemento no aplicativo do host. Por exemplo, recarregue um suplemento do painel tarefas escolhendo o canto superior direito do painel de tarefas para ativar o [menu personalidade](../design/task-pane-add-ins.md#personality-menu) e, em seguida, escolha **Recarregar**.

### <a name="debug-the-code-for-an-outlook-add-in"></a>Depurar o código de um suplemento do Outlook

Após você [iniciar o projeto](#start-the-project) e o Visual Studio iniciar o Outlook para hospedar o suplemento, abra um item de compromisso ou uma mensagem de email. 

O Outlook ativa o suplemento para o item, contanto que os critérios de ativação sejam atendidos. A barra de suplementos aparece na parte superior da janela Inspetor ou Painel de Leitura, e o suplemento do Outlook aparece como um botão na barra de suplementos. Se o suplemento tiver um comando de suplemento, aparecerá um botão na faixa de opções, na guia padrão ou em uma guia personalizada especificada, e o suplemento não aparecerá na barra de suplementos.

Para exibir o suplemento do Outlook, escolha o botão do suplemento do Outlook. Depois do suplemento ser exibido no Outlook, você pode depurar seu código fazendo o seguinte:

1. No Visual Studio, defina pontos de interrupção no seu código.

2. No Outlook, interagir com o suplemento.

3. Como os pontos de interrupção são atingidos no Visual Studio, percorra o código conforme necessário.

Você pode alterar o código e examinar os efeitos das alterações no suplemento sem ter que fechar o Outlook e reiniciar o projeto. Após salvar as mudanças ao código, abra o menu de atalho do suplemento (no Outlook) e escolha **Recarregar**.

## <a name="next-steps"></a>Próximas etapas

Depois do suplemento funcionar conforme desejado, veja [Implantar e publicar o suplemento Office](../publish/publish.md) para saber mais como você pode distribuir o suplemento para os usuários.
