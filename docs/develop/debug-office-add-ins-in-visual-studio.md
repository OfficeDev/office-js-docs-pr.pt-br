---
title: Depurar suplementos do Office no Visual Studio
description: Use o Visual Studio para depurar suplementos do Office na área de trabalho do cliente Office no Windows
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 44a3d56a276d70e24a3b466e16dd24d264f6555d
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773815"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>Depurar suplementos do Office no Visual Studio

Este artigo descreve como usar o Visual Studio 2019 para depurar um suplemento do Office na área de trabalho do cliente Office no Windows. Se você estiver usando outra versão do Visual Studio, os procedimentos poderão variar um pouco.

> [!NOTE]
> Você não pode usar o Visual Studio para depurar suplementos do Office na Web ou Mac. Para obter informações sobre a depuração nestas plataformas, confira [Depurar suplementos do Office no Office na Web](../testing/debug-add-ins-in-office-online.md) ou [Depurar suplementos do Office no Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md).

## <a name="enable-debugging-for-add-in-commands-and-ui-less-code"></a>Habilitar a depuração para comandos de suplemento e código sem interface de usuário

Quando o Visual Studio depura o Office no Windows, o suplemento é hospedado na instância do navegador do Microsoft Internet Explorer ou do Microsoft Edge. Para determinar qual navegador está sendo usado em seu computador de desenvolvimento, confira [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).
> [!NOTE]
> A variável de ambiente JS_Debug não é mais necessária no procedimento a seguir. Para obter mais informações, confira [Comportamentos de depuração em suplementos Web do Office](https://developercommunity.visualstudio.com/content/problem/740413/office-development-inconsistent-script-debugging-b.html), no fórum de suporte da Comunidade de Desenvolvedores da Microsoft.

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

## <a name="review-the-build-and-debug-properties"></a>Examinar as propriedades de compilação e depuração

Antes de iniciar Office depuração, revise as propriedades de cada projeto para confirmar se o Visual Studio abrirá o aplicativo Office desejado e que outras propriedades de compuração e depuração serão definidas adequadamente.

### <a name="add-in-project-properties"></a>Propriedades do projeto de suplemento

Abra a **janela Propriedades** do projeto do complemento para revisar as propriedades do projeto.

1. No **Explorador de soluções** Escolha o projeto do suplemento (*não* o projeto do aplicativo Web).

2. Na barra de menus, escolha **Exibir**,  > **Janela Propriedades**.

A tabela a seguir descreve as propriedades do projeto.

|Propriedade|Descrição|
|:-----|:-----|
|**Iniciar Ação**|Especifica o modo de depuração do suplemento. Atualmente, só **cliente de área de trabalho do Office** modo tem suporte para projetos de suplementos do Office.|
|**Iniciar documento**<br/> (apenas suplementos Excel, PowerPoint e Word)|Especifica o documento a ser aberto quando você iniciar o projeto.|
|**Projeto da Web**|Especifica o nome do projeto Web associado ao suplemento.|
|**Email Address**<br/>(Apenas suplementos do Outlook)|Especifica o endereço de email da conta de usuário no Exchange Server ou no Exchange Online que você quer usar para testar o suplemento do Outlook.|
|**EWS Url**<br/>(Apenas suplementos do Outlook)|URL do serviço Web do Exchange (por exemplo: `https://www.contoso.com/ews/exchange.aspx`). |
|**OWA Url**<br/>(Apenas suplementos do Outlook)|Outlook na URL da Web (por exemplo: `https://www.contoso.com/owa`).|
|**Usar autenticação multifator**<br/>(Apenas suplementos do Outlook)|Valor Booleano que indica se a autenticação multifator deve ser utilizada.|
|**Nome de Usuário**<br/>(Apenas suplementos do Outlook)|Especifica o nome da conta de usuário no Exchange Server ou no Exchange Online com a qual você deseja testar o suplemento do Outlook.|
|**Arquivo do projeto**|Especifica o nome do arquivo que contém informações de compilação, configuração e outras informações sobre o projeto.|
|**Pasta do projeto**|O local do arquivo do projeto.|

> [!NOTE]
> Para um suplemento do Outlook, você pode optar por especificar valores para uma ou mais das propriedades *Apenas suplemento Outlook* na janela **Propriedades** mas isso não é necessário.

### <a name="web-application-project-properties"></a>Propriedades do projeto de aplicativo Web

Abra a **janela Propriedades** do projeto do aplicativo Web para revisar as propriedades do projeto.

1. No **Explorador de Soluções,** escolha o projeto do aplicativo Web.

2. Na barra de menus, escolha **Exibir**,  > **Janela Propriedades**.

A tabela a seguir descreve as propriedades do projeto de aplicativo web que são mais relevantes para projetos de suplementos do Office.

|Propriedade|Descrição|
|:-----|:-----|
|**SSL habilitado**|Especifica se o SSL está habilitado no site. Essa propriedade deve ser definida como **Verdadeira** para projetos de suplementos do Office.|
|**URL SSL**|Especifica a URL HTTPS segura para o site. Somente leitura.|
|**URL**|Especifica a URL HTTP para o site. Somente leitura.|
|**Arquivo do projeto**|Especifica o nome do arquivo que contém informações de compilação, configuração e outras informações sobre o projeto.|
|**Pasta do projeto**|Especifica o local do arquivo do projeto. Somente leitura. O arquivo de manifesto do Visual Studio gerado no tempo de execução é escrito para a pasta `bin\Debug\OfficeAppManifests` neste local.|

## <a name="use-an-existing-document-to-debug-the-add-in"></a>Usar um documento existente para depurar o suplemento

Se você tiver um documento que contém os dados de teste deseja usar ao depurar seu suplemento o Excel, PowerPoint ou Word, o Visual Studio pode ser configurado para abrir esse documento quando você iniciar o projeto. Para especificar um documento existente a ser usado durante a depuração do complemento, execute as etapas a seguir.

1. No **Explorador de soluções** Escolha o projeto do suplemento (*não* o projeto do aplicativo Web).

2. Na barra de menus, escolha **Projeto** > **Adicionar Item Existente**.

3. Na caixa de diálogo **Adicionar Item Existente**, localize e selecione o documento que você deseja adicionar.

4. Escolha o botão **Adicionar** para adicionar o documento ao projeto.

5. No **Explorador de soluções** Escolha o projeto do suplemento (*não* o projeto do aplicativo Web).

6. Na barra de menus, escolha **Exibir**,  > **Janela Propriedades**.

7. Na janela **Propriedades**, escolha a lista **Iniciar Documento** e selecione o documento que você adicionou ao projeto. O projeto agora está configurado para iniciar o suplemento nesse documento.

## <a name="start-the-project"></a>Iniciar o projeto

Iniciar o projeto escolhendo **Depurar** > **Iniciar Depuração** na barra do menu. O Visual Studio compilará automaticamente a inicie o Office para hospedar o suplemento.

> [!NOTE]
> Quando você inicia um projeto de um suplemento do Outlook, você será solicitado a inserir as credenciais de logon. Se você for solicitado a entrar repetidamente ou se receber um erro não autorizado, a Auth Básica poderá ser desabilitada para contas em seu locatário Microsoft 365 locatário. Nesse caso, tente usar uma conta da Microsoft. Você também pode precisar definir a propriedade "Usar autenticação multifator" como Verdadeiro na caixa de diálogo de propriedades do projeto de suplemento do Outlook na Web.

Quando Visual Studio cria o projeto, ele executa as seguintes tarefas.

1. Cria uma cópia do arquivo de manifesto XML e a adiciona ao diretório `_ProjectName_\bin\Debug\OfficeAppManifests`. O Office aplicativo que hospeda o seu add-in consome essa cópia quando você Visual Studio e depura o complemento.

2. Cria um conjunto de entradas do Registro em seu computador que permitem que o add-in apareça no Office aplicativo.

3. Compila o projeto de aplicativo Web e o implanta no servidor Web IIS local(https://localhost).

4. Se este for o primeiro projeto de suplemento implantado no servidor Web do IIS local, talvez seja solicitado que você instale um Certificado Autoassinado no repositório de Certificado Raiz Confiável do usuário atual. Isso é necessário para que o IIS Express exiba o conteúdo do seu suplemento corretamente.

> [!NOTE]
> A versão mais recente do Office pode usar um controle da Web mais recente para exibir o conteúdo do suplemento ao ser executado no Windows 10. Se este for o caso, o Visual Studio pode solicitar que você adicione uma isenção de loopback de rede local. Isso é necessário para que o controle web, no aplicativo cliente Office, possa acessar o site implantado no servidor Web local do IIS. Você também pode alterar essa configuração a qualquer momento no Visual Studio, em **Ferramentas** > **Opções** > **Ferramentas do Office (Web)** > **Depuração do Suplemento da Web**.

Depois, o Visual Studio faz o seguinte:

1. Modifica o elemento [SourceLocation](../reference/manifest/sourcelocation.md) do arquivo de manifesto XML, substituindo o token `~remoteAppUrl` pelo endereço totalmente qualificado da página inicial (por exemplo,`https://localhost:44302/Home.html` ).

2. Inicia o projeto de aplicativo Web no IIS Express.

3. Abre o Office aplicativo.

Visual Studio não mostra erros de validação na janela **SAÍDA** ao criar o projeto. O Visual Studio relata erros e avisos na janela **ERRORLIST** à medida que eles ocorrem. O Visual Studio também relata erros de validação mostrando sublinhados ondulados (conhecidos como rabiscos) de cores diferentes no editor de código e texto. Essas marcas o notificam de problemas que o Visual Studio detectou no código. Para saber mais sobre como habilitar ou desabilitar a validação, confira: [Opções, Editor de texto, JavaScript, IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2019&preserve-view=true).

Para examinar as regras de validação do arquivo de manifesto XML no projeto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).

## <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a>Depurar o código de um suplemento Excel, PowerPoint ou Word

Se o seu complemento não estiver visível no documento exibido no aplicativo Office (Excel, PowerPoint ou Word) depois de iniciar o [projeto,](#start-the-project)iniciar manualmente o complemento no aplicativo Office. Por exemplo, inicie o suplemento do painel de tarefas, escolhendo o **Mostrar painel de tarefas** botão na faixa de opções da guia **Home**. Depois do suplemento ser exibido no Excel, PowerPoint ou Word, você pode depurar seu código fazendo o seguinte:

1. No Excel, PowerPoint ou Word, escolha o **Inserir** pressione tab e, em seguida, escolha a seta para baixo à direita de **Meus suplementos**.

    ![Captura de tela mostrando Insert ribbon in Excel on Windows com a seta Meus Complementos realçada.](../images/excel-cf-register-add-in-1b.png)

2. Na lista de suplementos disponíveis, localize a seção **suplementos do desenvolvedor** e selecione o seu suplemento para registrar.

3. No Visual Studio, defina pontos de interrupção no seu código.

4. No Excel, PowerPoint ou Word, interaja com o suplemento.

5. Como os pontos de interrupção são atingidos no Visual Studio, percorra o código conforme necessário.

Você pode alterar seu código e revisar os efeitos dessas alterações no seu complemento sem precisar fechar o aplicativo Office e reiniciar o projeto. Depois de salvar as alterações no código, basta recarregar o complemento no Office aplicativo. Por exemplo, recarregue um suplemento do painel tarefas escolhendo o canto superior direito do painel de tarefas para ativar o [menu personalidade](../design/task-pane-add-ins.md#personality-menu) e, em seguida, escolha **Recarregar**.

## <a name="debug-the-code-for-an-outlook-add-in"></a>Depurar o código de um suplemento do Outlook

Após você [iniciar o projeto](#start-the-project) e o Visual Studio iniciar o Outlook para hospedar o suplemento, abra um item de compromisso ou uma mensagem de email.

O Outlook ativa o suplemento para o item, contanto que os critérios de ativação sejam atendidos. A barra de suplementos aparece na parte superior da janela Inspetor ou Painel de Leitura, e o suplemento do Outlook aparece como um botão na barra de suplementos. Se o suplemento tiver um comando de suplemento, aparecerá um botão na faixa de opções, na guia padrão ou em uma guia personalizada especificada, e o suplemento não aparecerá na barra de suplementos.

Para exibir o suplemento do Outlook, escolha o botão do suplemento do Outlook. Depois do suplemento ser exibido no Outlook, você pode depurar seu código fazendo o seguinte:

1. No Visual Studio, defina pontos de interrupção no seu código.

2. No Outlook, interagir com o suplemento.

3. Como os pontos de interrupção são atingidos no Visual Studio, percorra o código conforme necessário.

Você pode alterar o código e examinar os efeitos das alterações no suplemento sem ter que fechar o Outlook e reiniciar o projeto. Após salvar as mudanças ao código, abra o menu de atalho do suplemento (no Outlook) e escolha **Recarregar**.

## <a name="next-steps"></a>Próximas etapas

Depois do suplemento funcionar conforme desejado, veja [Implantar e publicar o suplemento Office](../publish/publish.md) para saber mais como você pode distribuir o suplemento para os usuários.
