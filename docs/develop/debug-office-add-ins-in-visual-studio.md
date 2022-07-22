---
title: Depurar suplementos do Office no Visual Studio
description: Use o Visual Studio para depurar suplementos do Office no cliente da área de trabalho do Office no Windows.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 09693f81c069aba97740265fa88bf117a937c742
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958710"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>Depurar Suplementos do Office no Visual Studio

Este artigo descreve como depurar o código do lado do cliente em Suplementos do Office criados com um dos modelos de projeto do Suplemento do Office no Visual Studio 2022.  Para obter informações sobre como depurar código do lado do servidor em Suplementos do Office, consulte Visão geral da [depuração de suplementos do Office –](../testing/debug-add-ins-overview.md#server-side-or-client-side) lado do servidor ou do cliente?.

> [!NOTE]
> Você não pode usar o Visual Studio para depurar suplementos no Office no Mac. Para obter informações sobre depuração em um Mac, consulte [Depurar Suplementos do Office em um Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md).

## <a name="review-the-build-and-debug-properties"></a>Examinar as propriedades de compilação e depuração

Antes de iniciar a depuração, examine as propriedades de cada projeto para confirmar que o Visual Studio abrirá o aplicativo do Office desejado e que outras propriedades de compilação e depuração estão definidas adequadamente.

### <a name="add-in-project-properties"></a>Propriedades do projeto de suplemento

Abra a **janela Propriedades** do projeto de suplemento para examinar as propriedades do projeto.

1. No **Explorador de soluções** Escolha o projeto do suplemento (*não* o projeto do aplicativo Web).

2. Na barra de menus, escolha **Exibir**,  > **Janela Propriedades**.

A tabela a seguir descreve as propriedades do projeto.

|Propriedade|Descrição|
|:-----|:-----|
|**Iniciar Ação**|Especifica o modo de depuração do suplemento. Isso deve ser definido como **Microsoft Edge** para um suplemento do Outlook. Para todos os outros aplicativos do Office, ele deve ser definido como **Cliente de Área de Trabalho do Office**.|
|**Iniciar documento**<br/> (apenas suplementos Excel, PowerPoint e Word)|Especifica o documento a ser aberto quando você iniciar o projeto. Em um novo projeto, isso é definido como **[Nova** Pasta de Trabalho do Excel], [Novo Documento **do Word]** ou **[Nova Apresentação do PowerPoint]**. Para especificar um documento específico, siga as etapas em [Usar um documento existente para depurar o suplemento](#use-an-existing-document-to-debug-the-add-in).|
|**Projeto da Web**|Especifica o nome do projeto Web associado ao suplemento.|
|**Email Address**<br/>(Apenas suplementos do Outlook)|Especifica o endereço de email da conta de usuário no Exchange Server ou no Exchange Online que você quer usar para testar o suplemento do Outlook. Se deixado em branco, você será solicitado a fornecer o endereço de email ao iniciar a depuração.|
|**EWS Url**<br/>(Apenas suplementos do Outlook)|Especifica a URL dos Serviços Web do Exchange (por exemplo: `https://www.contoso.com/ews/exchange.aspx`). Essa propriedade pode ser deixada em branco.|
|**OWA Url**<br/>(Apenas suplementos do Outlook)|Especifica a URL Outlook na Web dados (por exemplo: `https://www.contoso.com/owa`). Essa propriedade pode ser deixada em branco.|
|**Usar autenticação multifator**<br/>(Apenas suplementos do Outlook)|Especifica o valor booliano que indica se a autenticação multifator deve ser usada. O padrão é **false**, mas a propriedade não tem nenhum efeito prático. Se você normalmente precisar fornecer um segundo fator para fazer logon na conta de email, será solicitado quando iniciar a depuração. |
|**Nome de Usuário**<br/>(Apenas suplementos do Outlook)|Especifica o nome da conta de usuário no Exchange Server ou no Exchange Online com a qual você deseja testar o suplemento do Outlook. Essa propriedade pode ser deixada em branco.|
|**Arquivo do projeto**|Especifica o nome do arquivo que contém informações de compilação, configuração e outras informações sobre o projeto.|
|**Pasta do projeto**|Especifica o local do arquivo do projeto.|

> [!NOTE]
> Para um suplemento do Outlook, você pode optar por especificar valores para uma ou mais das propriedades *Apenas suplemento Outlook* na janela **Propriedades** mas isso não é necessário.

### <a name="web-application-project-properties"></a>Propriedades do projeto de aplicativo Web

Abra a **janela Propriedades** do projeto de aplicativo Web para examinar as propriedades do projeto.

1. Em **Gerenciador de Soluções**, escolha o projeto de aplicativo Web.

2. Na barra de menus, escolha **Exibir**,  > **Janela Propriedades**.

A tabela a seguir descreve as propriedades do projeto de aplicativo web que são mais relevantes para projetos de suplementos do Office.

|Propriedade|Descrição|
|:-----|:-----|
|**SSL habilitado**|Especifica se o SSL está habilitado no site. Essa propriedade deve ser definida como **Verdadeira** para projetos de suplementos do Office.|
|**URL SSL**|Especifica a URL HTTPS segura para o site. Somente leitura.|
|**URL**|Especifica a URL HTTP para o site. Somente leitura.|
|**Arquivo do projeto**|Especifica o nome do arquivo que contém informações de compilação, configuração e outras informações sobre o projeto.|
|**Pasta do projeto**|Especifica o local do arquivo do projeto. Somente leitura. O arquivo de manifesto do Visual Studio gerado no tempo de execução é escrito para a pasta `bin\Debug\OfficeAppManifests` neste local.|

## <a name="debug-an-excel-powerpoint-or-word-add-in-project"></a>Depurar um projeto de suplemento do Excel, PowerPoint ou Word

Esta seção descreve como iniciar e depurar um suplemento do Excel, PowerPoint ou Word.

### <a name="start-the-excel-powerpoint-or-word-add-in-project"></a>Iniciar o projeto de suplemento do Excel, PowerPoint ou Word

Inicie o projeto escolhendo **Depurar** > **Iniciar Depuração** na barra de menus ou pressione o botão F5. O Visual Studio criará automaticamente a solução e iniciará o aplicativo host do Office.

Quando o Visual Studio compila o projeto, ele executa as seguintes tarefas:

1. Cria uma cópia do arquivo de manifesto XML e a adiciona ao  `_ProjectName_\bin\Debug\OfficeAppManifests` diretório. O aplicativo do Office que hospeda o suplemento consome essa cópia quando você inicia o Visual Studio e depura o suplemento.

2. Cria um conjunto de entradas do Registro em seu computador Windows que permite que o suplemento apareça no aplicativo do Office.

3. Compila o projeto de aplicativo Web e, em seguida, o implanta no servidor Web do IIS local (`https://localhost`).

4. Se este for o primeiro projeto de suplemento implantado no servidor Web do IIS local, talvez seja solicitado que você instale um certificado Self-Signed para o repositório de Certificados Raiz Confiáveis do usuário atual. Isso é necessário para que o IIS Express exiba o conteúdo do seu suplemento corretamente.

> [!NOTE]
> Se o Office usar o controle edge Legacy webview (EdgeHTML) para executar suplementos em seu computador Windows, o Visual Studio poderá solicitar que você adicione uma isenção de loopback de rede local. Isso é necessário para que o controle de modo de exibição da Web possa acessar o site implantado no servidor Web do IIS local. Você também pode alterar essa configuração a qualquer momento no Visual Studio, em **Ferramentas** > **Opções** > **Ferramentas do Office (Web)** > **Depuração do Suplemento da Web**. Para descobrir qual controle de navegador é usado em seu computador Windows, consulte [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

Depois, o Visual Studio faz o seguinte:

1. Modifica o [elemento SourceLocation](/javascript/api/manifest/sourcelocation) do arquivo de manifesto XML ( `_ProjectName_\bin\Debug\OfficeAppManifests` que foi copiado para o diretório) `~remoteAppUrl` substituindo o token pelo endereço totalmente qualificado da página inicial (por exemplo, `https://localhost:44302/Home.html`).

2. Inicia o projeto de aplicativo Web no IIS Express.

3. Valida o manifesto. Para examinar as regras de validação do arquivo de manifesto XML no projeto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md). 

   > [!IMPORTANT]
   > Os arquivos XSD de manifesto do Office instalados pelo Visual Studio estão desatualizados. Se você receber erros de validação para o manifesto, sua primeira etapa de solução de problemas deverá ser substituir um ou mais desses arquivos com as versões mais recentes. Para obter instruções detalhadas, consulte [Erros de validação de esquema de manifesto em projetos do Visual Studio](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

4. Abre o aplicativo do Office e o sideload do suplemento.

### <a name="debug-the-excel-powerpoint-or-word-add-in"></a>Depurar o suplemento Excel, PowerPoint ou Word

1. Inicie o suplemento no aplicativo do Office. Por exemplo, se for um suplemento do painel de tarefas, ele terá adicionado um botão à faixa de opções Página  Inicial (por exemplo, um botão Mostrar Painel **de Tarefas**). Selecione o botão na faixa de opções. 

   > [!NOTE]
   > Se o suplemento não for sideload pelo Visual Studio, você poderá fazer sideload dele manualmente. No Excel, no PowerPoint ou no Word, escolha  a guia Inserir e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.
   >
   > ![Captura de tela mostrando a faixa de opções Inserir no Excel no Windows com a seta Meus Suplementos realçada.](../images/excel-cf-register-add-in-1b.png)
   >
   > Na lista de suplementos disponíveis, localize a seção **suplementos do desenvolvedor** e selecione o seu suplemento para registrar.

   > [!TIP]
   > O painel de tarefas pode aparecer em branco quando ele é aberto pela primeira vez. Nesse caso, ele deverá ser renderizado corretamente quando você iniciar as ferramentas de depuração em uma etapa posterior.

3. Abra o [menu de personalidade](../design/task-pane-add-ins.md#personality-menu) e escolha **Anexar um depurador**. Isso abrirá as ferramentas de depuração para o controle de modo de exibição da Web que o Office está usando para executar suplementos em seu computador Windows. Você pode definir pontos de interrupção e percorrer o código conforme descrito em um dos seguintes artigos:

    - [Depurar os suplementos usando as ferramentas de desenvolvedor para o Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md)
    - [Depurar suplementos usando ferramentas de desenvolvedor para Edge Legacy](../testing/debug-add-ins-using-devtools-edge-legacy.md)
    - [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](../testing/debug-add-ins-using-devtools-edge-chromium.md)

4. Para fazer alterações no código, primeiro interrompa a sessão de depuração no Visual Studio e feche o aplicativo do Office. Faça suas alterações e inicie uma nova sessão de depuração.

## <a name="debug-an-outlook-add-in-project"></a>Depurar um projeto de suplemento do Outlook

Esta seção descreve como iniciar e depurar um suplemento do Outlook.

### <a name="start-the-outlook-add-in-project"></a>Iniciar o projeto de suplemento do Outlook

Inicie o projeto escolhendo **Depurar** > **Iniciar Depuração** na barra de menus ou pressione o botão F5. O Visual Studio criará automaticamente a solução e iniciará a página do Outlook de sua locação do Microsoft 365.

Quando o Visual Studio compila o projeto, ele executa as tarefas a seguir.

1. Solicita credenciais de logon. Se você for solicitado a entrar repetidamente ou se receber um erro de que não está autorizado, a Autenticação Básica poderá ser desabilitada para contas em seu locatário do Microsoft 365. Nesse caso, tente usar uma conta da Microsoft. Você também pode tentar definir a propriedade **Usar autenticação multifator** como **True** no painel de propriedades do projeto do Suplemento web do Outlook. Consulte [as propriedades do projeto de suplemento](#add-in-project-properties).

1. Cria uma cópia do arquivo de manifesto XML e a adiciona ao `_ProjectName_\bin\Debug\OfficeAppManifests` diretório. O Outlook consome essa cópia quando você inicia o Visual Studio e depura o suplemento.

2. Compila o projeto de aplicativo Web e, em seguida, o implanta no servidor Web do IIS local (`https://localhost`).

3. Se este for o primeiro projeto de suplemento implantado no servidor Web do IIS local, talvez seja solicitado que você instale um certificado Self-Signed para o repositório de Certificados Raiz Confiáveis do usuário atual. Isso é necessário para que o IIS Express exiba o conteúdo do seu suplemento corretamente.

> [!NOTE]
> Se o Office usar o controle edge Legacy webview (EdgeHTML) para executar suplementos em seu computador Windows, o Visual Studio poderá solicitar que você adicione uma isenção de loopback de rede local. Isso é necessário para que o controle de modo de exibição da Web possa acessar o site implantado no servidor Web do IIS local. Você também pode alterar essa configuração a qualquer momento no Visual Studio, em **Ferramentas** > **Opções** > **Ferramentas do Office (Web)** > **Depuração do Suplemento da Web**. Para descobrir qual controle de navegador é usado em seu computador Windows, consulte [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

Depois, o Visual Studio faz o seguinte:

1. Modifica o [elemento SourceLocation](/javascript/api/manifest/sourcelocation) do arquivo de manifesto XML ( `_ProjectName_\bin\Debug\OfficeAppManifests` que foi copiado para o diretório) `~remoteAppUrl` substituindo o token pelo endereço totalmente qualificado da página inicial (por exemplo, `https://localhost:44302/Home.html`).

2. Inicia o projeto de aplicativo Web no IIS Express.

3. Valida o manifesto. Para examinar as regras de validação do arquivo de manifesto XML no projeto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md). 

   > [!IMPORTANT]
   > Os arquivos XSD de manifesto do Office instalados pelo Visual Studio estão desatualizados. Se você receber erros de validação para o manifesto, sua primeira etapa de solução de problemas deverá ser substituir um ou mais desses arquivos com as versões mais recentes. Para obter instruções detalhadas, consulte [Erros de validação de esquema de manifesto em projetos do Visual Studio](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

4. Abre a página do Outlook da locação do Microsoft 365 no Microsoft Edge.

### <a name="debug-the-outlook-add-in"></a>Depurar o suplemento do Outlook

1. Na página do Outlook, selecione uma mensagem de email ou item de compromisso para abri-lo em sua própria janela. 

2. Pressione F12 para abrir a ferramenta de depuração do Edge.

3. Depois que a ferramenta estiver aberta, inicie o suplemento. Por exemplo, na barra de ferramentas na parte superior de uma mensagem, selecione o  botão Mais aplicativos e, em seguida, selecione o suplemento no texto explicativo que é aberto.

   ![Captura de tela mostrando o botão Mais aplicativos e o texto explicativo que ele abre com o nome e o ícone do suplemento visíveis junto com outros ícones de aplicativo.](../images/outlook-more-apps-button.png)

4. Use as instruções em um dos artigos a seguir para definir pontos de interrupção e percorrer o código. Cada um deles tem um link para diretrizes mais detalhadas.

   - [Depurar suplementos usando ferramentas de desenvolvedor para Edge Legacy](../testing/debug-add-ins-using-devtools-edge-legacy.md)
   - [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](../testing/debug-add-ins-using-devtools-edge-chromium.md)

   > [!TIP]
   > Para depurar `Office.initialize` `Office.onReady` o código que é executado na função ou em uma função executada quando o suplemento é aberto, defina seus pontos de interrupção e feche e reabra o suplemento. Para obter mais informações sobre essas funções, consulte [Inicializar seu Suplemento do Office](../develop/initialize-add-in.md).

5. Para fazer alterações no código, primeiro interrompa a sessão de depuração no Visual Studio e feche as páginas do Outlook. Faça suas alterações e inicie uma nova sessão de depuração.

## <a name="use-an-existing-document-to-debug-the-add-in"></a>Usar um documento existente para depurar o suplemento

Se você tiver um documento que contém os dados de teste deseja usar ao depurar seu suplemento o Excel, PowerPoint ou Word, o Visual Studio pode ser configurado para abrir esse documento quando você iniciar o projeto. Para especificar um documento existente a ser usado durante a depuração do complemento, execute as etapas a seguir.

1. No **Explorador de soluções** Escolha o projeto do suplemento (*não* o projeto do aplicativo Web).

2. Na barra de menus, escolha **Projeto** > **Adicionar Item Existente**.

3. Na caixa de diálogo **Adicionar Item Existente**, localize e selecione o documento que você deseja adicionar.

4. Escolha o botão **Adicionar** para adicionar o documento ao projeto.

5. No **Explorador de soluções** Escolha o projeto do suplemento (*não* o projeto do aplicativo Web).

6. Na barra de menus, escolha **Exibir**,  > **Janela Propriedades**.

7. Na janela **Propriedades**, escolha a lista **Iniciar Documento** e selecione o documento que você adicionou ao projeto. O projeto agora está configurado para iniciar o suplemento nesse documento.

## <a name="next-steps"></a>Próximas etapas

Depois do suplemento funcionar conforme desejado, veja [Implantar e publicar o suplemento Office](../publish/publish.md) para saber mais como você pode distribuir o suplemento para os usuários.
