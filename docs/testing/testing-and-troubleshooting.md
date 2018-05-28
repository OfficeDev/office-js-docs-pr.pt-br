---
title: Solucionar erros de usu?rios com suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 375b3819d423362c7d5e124700a0bea2dcf6e9e0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="29ddc-102">Solucionar erros de usu?rios com suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="29ddc-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="29ddc-p101">?s vezes, seus usu?rios podem encontrar problemas com suplementos do Office desenvolvidos por voc?. Por exemplo, um suplemento falha ao carregar ou est? inacess?vel. Use as informa??es neste artigo para ajudar a resolver problemas comuns que os usu?rios t?m com o seu suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="29ddc-106">Tamb?m ? poss?vel usar o [Fiddler](http://www.telerik.com/fiddler) para identificar e depurar problemas com os suplementos.</span><span class="sxs-lookup"><span data-stu-id="29ddc-106">You can also use [Fiddler](http://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

<span data-ttu-id="29ddc-107">Depois de resolver o problema do usu?rio, ? poss?vel [responder diretamente ?s avalia??es dos clientes no AppSource](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="29ddc-107">After you resolve the user's issue, you can [respond directly to customer reviews in AppSource](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings).</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="29ddc-108">Erros comuns e etapas de solu??o de problemas</span><span class="sxs-lookup"><span data-stu-id="29ddc-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="29ddc-109">A tabela a seguir lista as mensagens de erro comuns que os usu?rios podem receber e as etapas que os usu?rios podem seguir para resolver os erros.</span><span class="sxs-lookup"><span data-stu-id="29ddc-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="29ddc-110">**Mensagem de erro**</span><span class="sxs-lookup"><span data-stu-id="29ddc-110">**Error message**</span></span>|<span data-ttu-id="29ddc-111">**Resolu??o**</span><span class="sxs-lookup"><span data-stu-id="29ddc-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="29ddc-112">Erro do aplicativo: cat?logo n?o p?de ser alcan?ado</span><span class="sxs-lookup"><span data-stu-id="29ddc-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="29ddc-p102">Verifique as configura??es do firewall. "Cat?logo" refere-se ao AppSource. Essa mensagem indica que o usu?rio n?o consegue acessar o AppSource.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="29ddc-p103">ERRO DO APLICATIVO: este aplicativo n?o p?de ser iniciado. Feche essa caixa de di?logo para ignorar o problema ou clique em "Reiniciar"para tentar novamente.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="29ddc-117">Verifique se as atualiza??es mais recentes do Office foram instaladas ou baixe a [atualiza??o do Office 2013](https://support.microsoft.com/en-us/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="29ddc-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/en-us/kb/2986156/).</span></span>|
|<span data-ttu-id="29ddc-118">Erro: objeto n?o d? suporte ? propriedade ou ao m?todo 'defineProperty'</span><span class="sxs-lookup"><span data-stu-id="29ddc-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="29ddc-p104">Confirme se o Internet Explorer n?o est? sendo executado no modo de compatibilidade. V? para Ferramentas >  **Configura??es do Modo de Exibi??o de Compatibilidade**.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="29ddc-p105">N?o foi poss?vel carregar o aplicativo porque n?o h? suporte para sua vers?o do navegador. Clique aqui para obter uma lista de vers?es do navegador compat?veis.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="29ddc-p106">Verifique se o navegador d? suporte a armazenamento local HTML5 ou redefina as configura??es do Internet Explorer. Para saber mais sobre os navegadores compat?veis, confira [Requisitos para a execu??o de Suplementos do Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="29ddc-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|


## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="29ddc-125">O suplemento do Outlook n?o funciona corretamente</span><span class="sxs-lookup"><span data-stu-id="29ddc-125">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="29ddc-126">Se um suplemento do Outlook executado no Windows n?o est? funcionando corretamente, tente ativar a depura??o de scripts no Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="29ddc-126">If an Outlook add-in running on Windows is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="29ddc-127">V? para Ferramentas > **Op??es da Internet** > **Avan?ado**.</span><span class="sxs-lookup"><span data-stu-id="29ddc-127">Go to Tools >  **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="29ddc-128">Em **Navega??o**, desmarque **Desabilitar depura??o de scripts (Internet Explorer)** e **Desabilitar depura??o de scripts (Outros)**.</span><span class="sxs-lookup"><span data-stu-id="29ddc-128">Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="29ddc-p107">Recomendamos que voc? desmarque essas configura??es somente para solucionar o problema. Se voc? deixar desmarcado, receber? prompts durante a navega??o. Depois que o problema for resolvido, marque **Desabilitar depura??o de scripts (Internet Explorer)** e **Desabilitar depura??o de scripts (Outros)** novamente.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p107">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="29ddc-132">O suplemento n?o ? ativado no Office 2013</span><span class="sxs-lookup"><span data-stu-id="29ddc-132">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="29ddc-133">Se o suplemento n?o for ativado quando o usu?rio executar as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="29ddc-133">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="29ddc-134">Entrar com a conta da Microsoft no Office 2013.</span><span class="sxs-lookup"><span data-stu-id="29ddc-134">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="29ddc-135">Habilitar a verifica??o de duas etapas para a conta da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="29ddc-135">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="29ddc-136">Verificar a identidade ao ser solicitado quando tentar inserir um suplemento.</span><span class="sxs-lookup"><span data-stu-id="29ddc-136">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="29ddc-137">Verifique se as atualiza??es mais recentes do Office foram instaladas ou baixe a [atualiza??o do Office 2013](https://support.microsoft.com/en-us/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="29ddc-137">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/en-us/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="29ddc-138">N?o ? poss?vel carregar o suplemento no painel de tarefas ou outros problemas relacionados ao manifesto do suplemento</span><span class="sxs-lookup"><span data-stu-id="29ddc-138">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="29ddc-139">Confira [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md) para depurar problemas do manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="29ddc-139">See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="29ddc-140">N?o ? poss?vel exibir a caixa de di?logo do suplemento</span><span class="sxs-lookup"><span data-stu-id="29ddc-140">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="29ddc-p108">Quando o usu?rio usa um suplemento do Office, ele ? solicitado a permitir a exibi??o de uma caixa de di?logo. O usu?rio escolhe **Permitir** e, em seguida, recebe a seguinte mensagem de erro:</span><span class="sxs-lookup"><span data-stu-id="29ddc-p108">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="29ddc-p109">"As configura??es de seguran?a do navegador nos impedem de criar uma caixa de di?logo. Tente outro navegador ou configure o navegador para que a 'URL' e o dom?nio mostrado na barra de endere?o estejam na mesma zona de seguran?a".</span><span class="sxs-lookup"><span data-stu-id="29ddc-p109">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Captura de tela da mensagem de erro na caixa de di?logo](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="29ddc-146">**Navegadores afetados**</span><span class="sxs-lookup"><span data-stu-id="29ddc-146">**Affected browsers**</span></span>|<span data-ttu-id="29ddc-147">**Plataformas afetadas**</span><span class="sxs-lookup"><span data-stu-id="29ddc-147">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="29ddc-148">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="29ddc-148">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="29ddc-149">Office Online</span><span class="sxs-lookup"><span data-stu-id="29ddc-149">Office Online</span></span>|

<span data-ttu-id="29ddc-p110">Para resolver o problema, os administradores ou usu?rios finais podem adicionar o dom?nio do suplemento ? lista de sites confi?veis no Internet Explorer. Use o mesmo procedimento se estiver trabalhando com o navegador Internet Explorer ou Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p110">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="29ddc-152">Caso n?o confie no suplemento, n?o adicione a respectiva URL ? lista de sites confi?veis.</span><span class="sxs-lookup"><span data-stu-id="29ddc-152">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="29ddc-153">Para adicionar uma URL ? lista de sites confi?veis:</span><span class="sxs-lookup"><span data-stu-id="29ddc-153">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="29ddc-154">No Internet Explorer, escolha o bot?o Ferramentas e v? para **Op??es da Internet** > **Seguran?a**.</span><span class="sxs-lookup"><span data-stu-id="29ddc-154">In Internet Explorer, choose the Tools button, and go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="29ddc-155">Escolha a zona de **Sites confi?veis** e escolha **Sites**.</span><span class="sxs-lookup"><span data-stu-id="29ddc-155">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="29ddc-156">Insira a URL exibida na mensagem de erro e escolha **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="29ddc-156">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="29ddc-p111">Tente usar o suplemento novamente. Se o problema persistir, verifique as configura??es de outras zonas de seguran?a e confira se o dom?nio do suplemento est? na mesma zona que a URL exibida na barra de endere?os do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p111">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="29ddc-p112">Esse problema ocorre quando a API da caixa de di?logo ? usada no modo pop-up. Para evitar esse problema, use o sinalizador [displayInFrame](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync). Isso requer que a p?gina tenha suporte para exibi??o dentro de um iframe. O exemplo a seguir mostra como usar o sinalizador.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p112">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js

Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="29ddc-163">Altera??es nos comandos de suplemento, incluindo bot?es da faixa de op??es e itens de menu, n?o entram em vigor</span><span class="sxs-lookup"><span data-stu-id="29ddc-163">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>
<span data-ttu-id="29ddc-p113">?s vezes, as altera??es nos comandos de suplemento, como o ?cone de um bot?o da faixa de op??es ou o texto de um item de menu, n?o parecem entrar em vigor. Limpe o cache do Office das vers?es antigas.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p113">Sometimes changes to add-in commands such as the icon for a ribbon button or the text of a menu item do not seem to take effect. Clear the Office cache of the old versions.</span></span>

#### <a name="for-windows"></a><span data-ttu-id="29ddc-166">No Windows:</span><span class="sxs-lookup"><span data-stu-id="29ddc-166">For Windows:</span></span>
<span data-ttu-id="29ddc-167">Exclua o conte?do da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="29ddc-167">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="29ddc-168">No Mac:</span><span class="sxs-lookup"><span data-stu-id="29ddc-168">For Mac:</span></span>
<span data-ttu-id="29ddc-169">Exclua o conte?do da pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="29ddc-169">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="29ddc-170">No iOS:</span><span class="sxs-lookup"><span data-stu-id="29ddc-170">For iOS:</span></span>
<span data-ttu-id="29ddc-p114">Chame `window.location.reload(true)` usando o JavaScript no suplemento para for?ar um recarregamento. Outra alternativa ? reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="29ddc-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="29ddc-173">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="29ddc-173">See also</span></span>

- [<span data-ttu-id="29ddc-174">Depurar suplementos no Office Online</span><span class="sxs-lookup"><span data-stu-id="29ddc-174">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="29ddc-175">Realizar sideload de um suplemento do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="29ddc-175">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="29ddc-176">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="29ddc-176">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="29ddc-177">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="29ddc-177">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
    
