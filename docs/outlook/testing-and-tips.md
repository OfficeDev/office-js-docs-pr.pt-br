---
title: Implante e instale suplementos do Outlook para teste
description: Crie um arquivo de manifesto, implante o arquivo de interface do usuário suplemento em um servidor web, instale o suplemento na caixa de correio e teste o suplemento.
ms.date: 03/18/2020
localization_priority: Priority
ms.openlocfilehash: 76688ad3e1eca2dda832a94c3a9ae815e37678bc
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890974"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a><span data-ttu-id="6ef2d-103">Implante e instale suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="6ef2d-103">Deploy and install Outlook add-ins for testing</span></span>

<span data-ttu-id="6ef2d-104">Como parte do processo de desenvolvimento de um suplemento do Outlook, você provavelmente já se pegou fazendo a iteração da implantação e da instalação do suplemento para teste, que envolve as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="6ef2d-104">As part of the process of developing an Outlook add-in, you will probably find yourself iteratively deploying and installing the add-in for testing, which involves the following steps:</span></span>

1. <span data-ttu-id="6ef2d-105">Criação de um arquivo de manifesto que descreve o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-105">Creating a manifest file that describes the add-in.</span></span>
1. <span data-ttu-id="6ef2d-106">Implantação dos arquivos da interface do usuário em um servidor Web.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-106">Deploying the add-in UI file(s) to a web server.</span></span>
1. <span data-ttu-id="6ef2d-107">Instalação do suplemento em sua caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-107">Installing the add-in in your mailbox.</span></span>
1. <span data-ttu-id="6ef2d-108">Teste do suplemento, fazendo as alterações apropriadas na interface de usuário ou nos arquivos de manifesto e repetindo as etapas 2 e 3 para testar as alterações.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-108">Testing the add-in, making appropriate changes to the UI or manifest files, and repeating steps 2 and 3 to test the changes.</span></span>

> [!NOTE]
> <span data-ttu-id="6ef2d-109">[Os painéis personalizados foram preteridos](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/); portanto, certifique-se de estar usando um [ponto de extensão de suplemento com suporte](outlook-add-ins-overview.md#extension-points).</span><span class="sxs-lookup"><span data-stu-id="6ef2d-109">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using [a supported add-in extension point](outlook-add-ins-overview.md#extension-points).</span></span>

## <a name="create-a-manifest-file-for-the-add-in"></a><span data-ttu-id="6ef2d-110">Criar um arquivo de manifesto para o suplemento</span><span class="sxs-lookup"><span data-stu-id="6ef2d-110">Create a manifest file for the add-in</span></span>

<span data-ttu-id="6ef2d-p101">Cada suplemento é descrito por um manifesto XML, um documento que fornece as informações do servidor sobre o suplemento, fornece informações sobre o suplemento para o usuário e identifica o local da interface do arquivo HTML de interface do usuário do suplemento. Você pode armazenar o manifesto em uma pasta ou servidor local, desde que o local possa ser acessado pelo servidor Exchange da caixa de correio que você está testando. Vamos pressupor que você armazena seu manifesto em uma pasta local. Para obter informações sobre como criar um arquivo de manifesto, confira [Manifestos de suplementos do Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="6ef2d-p101">Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="deploy-an-add-in-to-a-web-server"></a><span data-ttu-id="6ef2d-115">Implantar um suplemento em um servidor Web</span><span class="sxs-lookup"><span data-stu-id="6ef2d-115">Deploy an add-in to a web server</span></span>

<span data-ttu-id="6ef2d-p102">Você pode usar HTML e JavaScript para criar o suplemento. Os arquivos de origem resultantes são armazenados em um servidor Web que pode ser acessado pelo servidor Exchange que hospeda o suplemento. Depois de implantar inicialmente os arquivos de origem para o suplemento, você pode atualizar a interface do usuário e o comportamento dele substituindo os arquivos HTML ou JavaScript armazenados no servidor Web por uma nova versão do arquivo HTML.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-p102">You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.</span></span>

## <a name="install-the-add-in"></a><span data-ttu-id="6ef2d-119">Instalar o suplemento</span><span class="sxs-lookup"><span data-stu-id="6ef2d-119">Install the add-in</span></span>

<span data-ttu-id="6ef2d-120">Depois de preparar o arquivo de manifesto do suplemento e implantar a interface de usuário do suplemento em um servidor Web que possa ser acessado, é possível realizar o sideload do suplemento para uma caixa de correio em um servidor Exchange usando um cliente do Outlook ou instalar o suplemento executando cmdlets remotos do Windows PowerShell.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-120">After preparing the add-in manifest file and deploying the add-in UI to a web server that can be accessed, you can sideload the add-in for a mailbox on an Exchange server by using an Outlook client, or install the add-in by running remote Windows PowerShell cmdlets.</span></span>

### <a name="sideload-the-add-in"></a><span data-ttu-id="6ef2d-121">Realizar o sideload do suplemento</span><span class="sxs-lookup"><span data-stu-id="6ef2d-121">Sideload the add-in</span></span>

<span data-ttu-id="6ef2d-p103">Você pode instalar um suplemento se sua caixa de correio está no Exchange Online, no Exchange 2013 ou em uma versão posterior. Os suplementos de sideload exigem no mínimo a função **Meus Aplicativos Personalizados** do seu Exchange Server. Para testar seu suplemento ou instalar suplementos em geral especificando uma URL ou um nome de arquivo de manifesto do suplemento, é preciso solicitar que o administrador do Exchange forneça as permissões necessárias.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-p103">You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.</span></span>

<span data-ttu-id="6ef2d-p104">O administrador do Exchange pode executar o cmdlet do PowerShell a seguir para atribuir as permissões necessárias a um único usuário. Neste exemplo, `wendyri` é o alias de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-p104">The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.</span></span>

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

<span data-ttu-id="6ef2d-127">Se necessário, o administrador pode executar o cmdlet a seguir para atribuir permissões necessárias semelhantes a vários usuários:</span><span class="sxs-lookup"><span data-stu-id="6ef2d-127">If necessary, the administrator can run the following cmdlet to assign multiple users the similar necessary permissions:</span></span>

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

<span data-ttu-id="6ef2d-128">Para saber mais sobre a função Meus Suplementos Personalizados, confira [Função Meus Suplementos Personalizados](/exchange/my-custom-apps-role-exchange-2013-help).</span><span class="sxs-lookup"><span data-stu-id="6ef2d-128">For more information about the My Custom Apps role, see [My Custom Apps role](/exchange/my-custom-apps-role-exchange-2013-help).</span></span>

<span data-ttu-id="6ef2d-129">O uso do Office 365 ou do Visual Studio para desenvolver suplementos atribui a você a função de administrador da organização, o que permite que você instale suplementos por arquivo ou URL no EAC, ou por cmdlets do Powershell.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-129">Using Office 365 or Visual Studio to develop add-ins assigns you the organization administrator role which allows you to install add-ins by file or URL in the EAC, or by Powershell cmdlets.</span></span>

### <a name="install-an-add-in-by-using-remote-powershell"></a><span data-ttu-id="6ef2d-130">Instalar um suplemento usando o PowerShell remoto</span><span class="sxs-lookup"><span data-stu-id="6ef2d-130">Install an add-in by using remote PowerShell</span></span>

<span data-ttu-id="6ef2d-131">Depois de criar uma sessão remota do Windows PowerShell em seu servidor Exchange, você pode instalar um suplemento do Outlook usando o cmdlet `New-App` com o comando do PowerShell a seguir.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-131">After you create a remote Windows PowerShell session on your Exchange server, you can install an Outlook add-in by using the `New-App` cmdlet with the following PowerShell command.</span></span>

```powershell
New-App -URL:"http://<fully-qualified URL">
```

<span data-ttu-id="6ef2d-132">A URL totalmente qualificada é o local do arquivo de manifesto do suplemento que você preparou para seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-132">The fully qualified URL is the location of the add-in manifest file that you prepared for your add-in.</span></span>

<span data-ttu-id="6ef2d-133">Você pode usar os seguintes cmdlets do PowerShell adicionais para gerenciar os suplementos de uma caixa de correio:</span><span class="sxs-lookup"><span data-stu-id="6ef2d-133">You can use the following additional PowerShell cmdlets to manage the add-ins for a mailbox:</span></span>

-  <span data-ttu-id="6ef2d-134">`Get-App` – Lista os suplementos que estão habilitados para uma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-134">`Get-App` - Lists the add-ins that are enabled for a mailbox.</span></span>
-  <span data-ttu-id="6ef2d-135">`Set-App` – Habilita ou desabilita um suplemento em uma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-135">`Set-App` - Enables or disables a add-in on a mailbox.</span></span>
-  <span data-ttu-id="6ef2d-136">`Remove-App` – Remove um suplemento instalado anteriormente de um servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-136">`Remove-App` - Removes a previously installed add-in from an Exchange server.</span></span>

## <a name="client-versions"></a><span data-ttu-id="6ef2d-137">Versões de cliente</span><span class="sxs-lookup"><span data-stu-id="6ef2d-137">Client versions</span></span>

<span data-ttu-id="6ef2d-138">A decisão de quais versões de cliente do Outlook testar depende dos seus requisitos de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-138">Deciding what versions of the Outlook client to test depends on your development requirements.</span></span>

- <span data-ttu-id="6ef2d-p105">Se você estiver desenvolvendo um suplemento para uso privado ou apenas para membros da sua organização, é importante testar as versões do Outlook usadas pela sua empresa. Lembre-se que alguns usuários podem usar o Outlook na Web, portanto testar as versões para o navegador padrão da sua empresa também é importante.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-p105">If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.</span></span>

- <span data-ttu-id="6ef2d-p106">Se você estiver desenvolvendo um suplemento no [AppSource](https://appsource.microsoft.com), teste as versões exigidas conforme especificado nas [Políticas de certificação do mercado comercial 1120.3](/legal/marketplace/certification-policies#11203-functionality). Isso inclui:</span><span class="sxs-lookup"><span data-stu-id="6ef2d-p106">If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:</span></span>
    - <span data-ttu-id="6ef2d-143">A versão mais recente do Outlook no Windows e a versão anterior à mais recente.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-143">The latest version of Outlook on Windows and the version prior to the latest.</span></span>
    - <span data-ttu-id="6ef2d-144">A versão mais recente do Outlook no Mac.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-144">The latest version of Outlook on Mac.</span></span>
    - <span data-ttu-id="6ef2d-145">A versão mais recente do Outlook no iOS e Android (se o suplemento for [compatível com mobilidade](add-mobile-support.md)).</span><span class="sxs-lookup"><span data-stu-id="6ef2d-145">The latest version of Outlook on iOS and Android (if your add-in [supports mobile form factor](add-mobile-support.md)).</span></span>
    - <span data-ttu-id="6ef2d-146">As versões do navegador especificadas na política de validação do mercado comercial 1120.3</span><span class="sxs-lookup"><span data-stu-id="6ef2d-146">The browser versions specified in the Commercial marketplace validation policy 1120.3.</span></span>

> [!NOTE]
> <span data-ttu-id="6ef2d-147">Se seu suplemento não for compatível com um dos clientes acima devido a uma [solicitação de um conjunto de exigências da API](apis.md) que o cliente não dá suporte, esse cliente será removido da lista de clientes exigidos.</span><span class="sxs-lookup"><span data-stu-id="6ef2d-147">If your add-in does not support one of the above clients due to [requesting an API requirement set](apis.md) that the client does not support, that client would be removed from the list of required clients.</span></span>

## <a name="see-also"></a><span data-ttu-id="6ef2d-148">Confira também</span><span class="sxs-lookup"><span data-stu-id="6ef2d-148">See also</span></span>

- [<span data-ttu-id="6ef2d-149">Solucionar erros de usuários com Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6ef2d-149">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
