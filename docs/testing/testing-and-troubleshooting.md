---
title: Solucionar erros de usuários com suplementos do Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 84f18543c7bafac905805095c89f8e19a855ea76
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449810"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="211d1-102">Solucionar erros de usuários com suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="211d1-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="211d1-p101">Às vezes, seus usuários podem encontrar problemas com suplementos do Office desenvolvidos por você. Por exemplo, um suplemento falha ao carregar ou está inacessível. Use as informações neste artigo para ajudar a resolver problemas comuns que os usuários têm com o seu suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="211d1-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="211d1-106">Também é possível usar o [Fiddler](https://www.telerik.com/fiddler) para identificar e depurar problemas com os suplementos.</span><span class="sxs-lookup"><span data-stu-id="211d1-106">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

<span data-ttu-id="211d1-107">Depois de resolver o problema do usuário, é possível [responder diretamente às avaliações dos clientes no AppSource](/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="211d1-107">After you resolve the user's issue, you can [respond directly to customer reviews in AppSource](/office/dev/store/create-effective-office-store-listings).</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="211d1-108">Erros comuns e etapas de solução de problemas</span><span class="sxs-lookup"><span data-stu-id="211d1-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="211d1-109">A tabela a seguir lista as mensagens de erro comuns que os usuários podem receber e as etapas que os usuários podem seguir para resolver os erros.</span><span class="sxs-lookup"><span data-stu-id="211d1-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="211d1-110">**Mensagem de erro**</span><span class="sxs-lookup"><span data-stu-id="211d1-110">**Error message**</span></span>|<span data-ttu-id="211d1-111">**Resolução**</span><span class="sxs-lookup"><span data-stu-id="211d1-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="211d1-112">Erro do aplicativo: catálogo não pôde ser alcançado</span><span class="sxs-lookup"><span data-stu-id="211d1-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="211d1-p102">Verifique as configurações do firewall. "Catálogo" refere-se ao AppSource. Essa mensagem indica que o usuário não consegue acessar o AppSource.</span><span class="sxs-lookup"><span data-stu-id="211d1-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="211d1-p103">ERRO DO APLICATIVO: este aplicativo não pôde ser iniciado. Feche essa caixa de diálogo para ignorar o problema ou clique em "Reiniciar"para tentar novamente.</span><span class="sxs-lookup"><span data-stu-id="211d1-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="211d1-117">Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="211d1-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="211d1-118">Erro: objeto não dá suporte à propriedade ou ao método 'defineProperty'</span><span class="sxs-lookup"><span data-stu-id="211d1-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="211d1-p104">Confirme se o Internet Explorer não está sendo executado no modo de compatibilidade. Vá para Ferramentas >  **Configurações do Modo de Exibição de Compatibilidade**.</span><span class="sxs-lookup"><span data-stu-id="211d1-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="211d1-p105">Não foi possível carregar o aplicativo porque não há suporte para sua versão do navegador. Clique aqui para obter uma lista de versões do navegador compatíveis.</span><span class="sxs-lookup"><span data-stu-id="211d1-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="211d1-p106">Verifique se o navegador dá suporte a armazenamento local HTML5 ou redefina as configurações do Internet Explorer. Para saber mais sobre os navegadores compatíveis, confira [Requisitos para a execução de Suplementos do Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="211d1-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|


## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="211d1-125">O suplemento do Outlook não funciona corretamente</span><span class="sxs-lookup"><span data-stu-id="211d1-125">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="211d1-126">Se um suplemento do Outlook executado no Windows não está funcionando corretamente, tente ativar a depuração de scripts no Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="211d1-126">If an Outlook add-in running on Windows is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="211d1-127">Vá para Ferramentas > **Opções da Internet** > **Avançado**.</span><span class="sxs-lookup"><span data-stu-id="211d1-127">Go to Tools >  **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="211d1-128">Em **Navegação**, desmarque **Desabilitar depuração de scripts (Internet Explorer)** e **Desabilitar depuração de scripts (Outros)**.</span><span class="sxs-lookup"><span data-stu-id="211d1-128">Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="211d1-p107">Recomendamos que você desmarque essas configurações somente para solucionar o problema. Se você deixar desmarcado, receberá prompts durante a navegação. Depois que o problema for resolvido, marque **Desabilitar depuração de scripts (Internet Explorer)** e **Desabilitar depuração de scripts (Outros)** novamente.</span><span class="sxs-lookup"><span data-stu-id="211d1-p107">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="211d1-132">O suplemento não é ativado no Office 2013</span><span class="sxs-lookup"><span data-stu-id="211d1-132">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="211d1-133">Se o suplemento não for ativado quando o usuário executar as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="211d1-133">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="211d1-134">Entrar com a conta da Microsoft no Office 2013.</span><span class="sxs-lookup"><span data-stu-id="211d1-134">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="211d1-135">Habilitar a verificação de duas etapas para a conta da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="211d1-135">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="211d1-136">Verificar a identidade ao ser solicitado quando tentar inserir um suplemento.</span><span class="sxs-lookup"><span data-stu-id="211d1-136">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="211d1-137">Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="211d1-137">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="211d1-138">Não é possível carregar o suplemento no painel de tarefas ou outros problemas relacionados ao manifesto do suplemento</span><span class="sxs-lookup"><span data-stu-id="211d1-138">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="211d1-139">Confira [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md) para depurar problemas do manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="211d1-139">See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="211d1-140">Não é possível exibir a caixa de diálogo do suplemento</span><span class="sxs-lookup"><span data-stu-id="211d1-140">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="211d1-p108">Quando o usuário usa um suplemento do Office, ele é solicitado a permitir a exibição de uma caixa de diálogo. O usuário escolhe **Permitir** e, em seguida, recebe a seguinte mensagem de erro:</span><span class="sxs-lookup"><span data-stu-id="211d1-p108">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="211d1-p109">"As configurações de segurança do navegador nos impedem de criar uma caixa de diálogo. Tente outro navegador ou configure o navegador para que a 'URL' e o domínio mostrado na barra de endereço estejam na mesma zona de segurança".</span><span class="sxs-lookup"><span data-stu-id="211d1-p109">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Captura de tela da mensagem de erro na caixa de diálogo](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="211d1-146">**Navegadores afetados**</span><span class="sxs-lookup"><span data-stu-id="211d1-146">**Affected browsers**</span></span>|<span data-ttu-id="211d1-147">**Plataformas afetadas**</span><span class="sxs-lookup"><span data-stu-id="211d1-147">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="211d1-148">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="211d1-148">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="211d1-149">Office Online</span><span class="sxs-lookup"><span data-stu-id="211d1-149">Office Online</span></span>|

<span data-ttu-id="211d1-p110">Para resolver o problema, os administradores ou usuários finais podem adicionar o domínio do suplemento à lista de sites confiáveis no Internet Explorer. Use o mesmo procedimento se estiver trabalhando com o navegador Internet Explorer ou Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="211d1-p110">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="211d1-152">Caso não confie no suplemento, não adicione a respectiva URL à lista de sites confiáveis.</span><span class="sxs-lookup"><span data-stu-id="211d1-152">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="211d1-153">Para adicionar uma URL à lista de sites confiáveis:</span><span class="sxs-lookup"><span data-stu-id="211d1-153">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="211d1-154">No Internet Explorer, escolha o botão Ferramentas e vá para **Opções da Internet** > **Segurança**.</span><span class="sxs-lookup"><span data-stu-id="211d1-154">In Internet Explorer, choose the Tools button, and go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="211d1-155">Escolha a zona de **Sites confiáveis** e escolha **Sites**.</span><span class="sxs-lookup"><span data-stu-id="211d1-155">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="211d1-156">Insira a URL exibida na mensagem de erro e escolha **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="211d1-156">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="211d1-p111">Tente usar o suplemento novamente. Se o problema persistir, verifique as configurações de outras zonas de segurança e confira se o domínio do suplemento está na mesma zona que a URL exibida na barra de endereços do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="211d1-p111">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="211d1-p112">Esse problema ocorre quando a API da caixa de diálogo é usada no modo pop-up. Para evitar esse problema, use o sinalizador [displayInFrame](/javascript/api/office/office.ui). Isso requer que a página tenha suporte para exibição dentro de um iframe. O exemplo a seguir mostra como usar o sinalizador.</span><span class="sxs-lookup"><span data-stu-id="211d1-p112">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js

Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="211d1-163">Alterações nos comandos de suplemento, incluindo botões da faixa de opções e itens de menu, não entram em vigor</span><span class="sxs-lookup"><span data-stu-id="211d1-163">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>
<span data-ttu-id="211d1-p113">Às vezes, as alterações nos comandos de suplemento, como o ícone de um botão da faixa de opções ou o texto de um item de menu, não parecem entrar em vigor. Limpe o cache do Office das versões antigas.</span><span class="sxs-lookup"><span data-stu-id="211d1-p113">Sometimes changes to add-in commands such as the icon for a ribbon button or the text of a menu item do not seem to take effect. Clear the Office cache of the old versions.</span></span>

#### <a name="for-windows"></a><span data-ttu-id="211d1-166">No Windows:</span><span class="sxs-lookup"><span data-stu-id="211d1-166">For Windows:</span></span>
<span data-ttu-id="211d1-167">Exclua o conteúdo da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="211d1-167">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="211d1-168">No Mac:</span><span class="sxs-lookup"><span data-stu-id="211d1-168">For Mac:</span></span>
<span data-ttu-id="211d1-169">Exclua o conteúdo da pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="211d1-169">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="211d1-170">No iOS:</span><span class="sxs-lookup"><span data-stu-id="211d1-170">For iOS:</span></span>
<span data-ttu-id="211d1-p114">Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="211d1-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="211d1-173">Confira também</span><span class="sxs-lookup"><span data-stu-id="211d1-173">See also</span></span>

- [<span data-ttu-id="211d1-174">Depurar suplementos no Office Online</span><span class="sxs-lookup"><span data-stu-id="211d1-174">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="211d1-175">Realizar sideload de um suplemento do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="211d1-175">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="211d1-176">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="211d1-176">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="211d1-177">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="211d1-177">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
    
