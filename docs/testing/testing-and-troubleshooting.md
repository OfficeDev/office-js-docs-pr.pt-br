---
title: Solucionar erros de usuários com suplementos do Office
description: Saiba como solucionar erros de usuários em suplementos do Office.
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 51f5ec406a09b18ece24b74dc22718e7fd422e38
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159182"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="8174a-103">Solucionar erros de usuários com suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8174a-103">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="8174a-p101">Às vezes, seus usuários podem encontrar problemas com suplementos do Office desenvolvidos por você. Por exemplo, um suplemento falha ao carregar ou está inacessível. Use as informações neste artigo para ajudar a resolver problemas comuns que os usuários têm com o seu suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="8174a-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="8174a-107">Também é possível usar o [Fiddler](https://www.telerik.com/fiddler) para identificar e depurar problemas com os suplementos.</span><span class="sxs-lookup"><span data-stu-id="8174a-107">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="8174a-108">Erros comuns e etapas de solução de problemas</span><span class="sxs-lookup"><span data-stu-id="8174a-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="8174a-109">A tabela a seguir lista as mensagens de erro comuns que os usuários podem receber e as etapas que os usuários podem seguir para resolver os erros.</span><span class="sxs-lookup"><span data-stu-id="8174a-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="8174a-110">**Mensagem de erro**</span><span class="sxs-lookup"><span data-stu-id="8174a-110">**Error message**</span></span>|<span data-ttu-id="8174a-111">**Resolução**</span><span class="sxs-lookup"><span data-stu-id="8174a-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="8174a-112">Erro do aplicativo: catálogo não pôde ser alcançado</span><span class="sxs-lookup"><span data-stu-id="8174a-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="8174a-p102">Verifique as configurações do firewall. "Catálogo" refere-se ao AppSource. Essa mensagem indica que o usuário não consegue acessar o AppSource.</span><span class="sxs-lookup"><span data-stu-id="8174a-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="8174a-p103">ERRO DO APLICATIVO: este aplicativo não pôde ser iniciado. Feche essa caixa de diálogo para ignorar o problema ou clique em "Reiniciar"para tentar novamente.</span><span class="sxs-lookup"><span data-stu-id="8174a-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="8174a-117">Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="8174a-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="8174a-118">Erro: objeto não dá suporte à propriedade ou ao método 'defineProperty'</span><span class="sxs-lookup"><span data-stu-id="8174a-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="8174a-119">Confirme se o Internet Explorer não está sendo executado no modo de compatibilidade.</span><span class="sxs-lookup"><span data-stu-id="8174a-119">Confirm that Internet Explorer is not running in Compatibility Mode.</span></span> <span data-ttu-id="8174a-120">Vá para ferramentas > **configurações do modo de exibição de compatibilidade**.</span><span class="sxs-lookup"><span data-stu-id="8174a-120">Go to Tools > **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="8174a-p105">Não foi possível carregar o aplicativo porque não há suporte para sua versão do navegador. Clique aqui para obter uma lista de versões do navegador compatíveis.</span><span class="sxs-lookup"><span data-stu-id="8174a-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="8174a-p106">Verifique se o navegador dá suporte a armazenamento local HTML5 ou redefina as configurações do Internet Explorer. Para saber mais sobre os navegadores compatíveis, confira [Requisitos para a execução de Suplementos do Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="8174a-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="8174a-125">Ao instalar um suplemento, você verá “erro ao carregar suplemento” na barra de status</span><span class="sxs-lookup"><span data-stu-id="8174a-125">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="8174a-126">Feche o Office.</span><span class="sxs-lookup"><span data-stu-id="8174a-126">Close Office.</span></span>
2. <span data-ttu-id="8174a-127">Verifique se o manifesto é valido</span><span class="sxs-lookup"><span data-stu-id="8174a-127">Verify that the manifest is valid</span></span>
3. <span data-ttu-id="8174a-128">Reinicie o suplemento</span><span class="sxs-lookup"><span data-stu-id="8174a-128">Restart the add-in</span></span>
4. <span data-ttu-id="8174a-129">Instale o suplemento novamente.</span><span class="sxs-lookup"><span data-stu-id="8174a-129">Install the add-in again.</span></span>

<span data-ttu-id="8174a-130">Você também pode enviar comentários: se estiver usando o Excel no Windows ou Mac, envie comentários à equipe de extensibilidade do Office diretamente do Excel.</span><span class="sxs-lookup"><span data-stu-id="8174a-130">You can also give us feedback: if using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="8174a-131">Para fazer isso, selecione **Arquivo** | **Comentário** | **Enviar um Rosto Triste**.</span><span class="sxs-lookup"><span data-stu-id="8174a-131">To do this, select **File** | **Feedback** | **Send a Frown**.</span></span> <span data-ttu-id="8174a-132">Enviando um rosto triste, você fornece os logs necessários para entendermos o problema.</span><span class="sxs-lookup"><span data-stu-id="8174a-132">Sending a frown provides the necessary logs to understand the issue.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="8174a-133">O suplemento do Outlook não funciona corretamente</span><span class="sxs-lookup"><span data-stu-id="8174a-133">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="8174a-134">Se um suplemento do Outlook executado no Windows e [usando o Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) não está funcionando corretamente, tente ativar a depuração de scripts no Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="8174a-134">If an Outlook add-in running on Windows and [using Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="8174a-135">Vá para ferramentas > **Opções da Internet**  >  **avançadas**.</span><span class="sxs-lookup"><span data-stu-id="8174a-135">Go to Tools > **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="8174a-136">Em **navegação**, desmarque **Desabilitar depuração de scripts (Internet Explorer)** e **Desabilitar depuração de scripts (outros)**.</span><span class="sxs-lookup"><span data-stu-id="8174a-136">Under **Browsing**, uncheck **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="8174a-137">Recomendamos que você desmarque essas configurações somente para solucionar o problema.</span><span class="sxs-lookup"><span data-stu-id="8174a-137">We recommend that you uncheck these settings only to troubleshoot the issue.</span></span> <span data-ttu-id="8174a-138">Se você deixar desmarcado, receberá prompts durante a navegação.</span><span class="sxs-lookup"><span data-stu-id="8174a-138">If you leave them unchecked, you will get prompts when you browse.</span></span> <span data-ttu-id="8174a-139">Depois que o problema for resolvido, marque **Desabilitar depuração de scripts (Internet Explorer)** e **Desabilitar depuração de scripts (outros)** novamente.</span><span class="sxs-lookup"><span data-stu-id="8174a-139">After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="8174a-140">O suplemento não é ativado no Office 2013</span><span class="sxs-lookup"><span data-stu-id="8174a-140">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="8174a-141">Se o suplemento não for ativado quando o usuário executar as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="8174a-141">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="8174a-142">Entrar com a conta da Microsoft no Office 2013.</span><span class="sxs-lookup"><span data-stu-id="8174a-142">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="8174a-143">Habilitar a verificação de duas etapas para a conta da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="8174a-143">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="8174a-144">Verificar a identidade ao ser solicitado quando tentar inserir um suplemento.</span><span class="sxs-lookup"><span data-stu-id="8174a-144">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="8174a-145">Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="8174a-145">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="8174a-146">Não é possível carregar o suplemento no painel de tarefas ou outros problemas relacionados ao manifesto do suplemento</span><span class="sxs-lookup"><span data-stu-id="8174a-146">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="8174a-147">Confira [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md) e [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md) para depurar problemas de manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="8174a-147">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="8174a-148">Não é possível exibir a caixa de diálogo do suplemento</span><span class="sxs-lookup"><span data-stu-id="8174a-148">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="8174a-p109">Quando o usuário usa um suplemento do Office, ele é solicitado a permitir a exibição de uma caixa de diálogo. O usuário escolhe **Permitir** e, em seguida, recebe a seguinte mensagem de erro:</span><span class="sxs-lookup"><span data-stu-id="8174a-p109">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="8174a-p110">"As configurações de segurança do navegador nos impedem de criar uma caixa de diálogo. Tente outro navegador ou configure o navegador para que a 'URL' e o domínio mostrado na barra de endereço estejam na mesma zona de segurança".</span><span class="sxs-lookup"><span data-stu-id="8174a-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Captura de tela da mensagem de erro na caixa de diálogo](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="8174a-154">**Navegadores afetados**</span><span class="sxs-lookup"><span data-stu-id="8174a-154">**Affected browsers**</span></span>|<span data-ttu-id="8174a-155">**Plataformas afetadas**</span><span class="sxs-lookup"><span data-stu-id="8174a-155">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="8174a-156">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="8174a-156">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="8174a-157">Office na Web</span><span class="sxs-lookup"><span data-stu-id="8174a-157">Office on the web</span></span>|

<span data-ttu-id="8174a-p111">Para resolver o problema, os administradores ou usuários finais podem adicionar o domínio do suplemento à lista de sites confiáveis no Internet Explorer. Use o mesmo procedimento se estiver trabalhando com o navegador Internet Explorer ou Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="8174a-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8174a-160">Caso não confie no suplemento, não adicione a respectiva URL à lista de sites confiáveis.</span><span class="sxs-lookup"><span data-stu-id="8174a-160">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="8174a-161">Para adicionar uma URL à lista de sites confiáveis:</span><span class="sxs-lookup"><span data-stu-id="8174a-161">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="8174a-162">No **Painel de Controle**, abra **Opções da Internet** > **Security**.</span><span class="sxs-lookup"><span data-stu-id="8174a-162">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="8174a-163">Escolha a zona de **Sites confiáveis** e escolha **Sites**.</span><span class="sxs-lookup"><span data-stu-id="8174a-163">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="8174a-164">Insira a URL exibida na mensagem de erro e escolha **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="8174a-164">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="8174a-p112">Tente usar o suplemento novamente. Se o problema persistir, verifique as configurações de outras zonas de segurança e confira se o domínio do suplemento está na mesma zona que a URL exibida na barra de endereços do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="8174a-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="8174a-p113">Esse problema ocorre quando a API da caixa de diálogo é usada no modo pop-up. Para evitar esse problema, use o sinalizador [displayInFrame](/javascript/api/office/office.ui). Isso requer que a página tenha suporte para exibição dentro de um iframe. O exemplo a seguir mostra como usar o sinalizador.</span><span class="sxs-lookup"><span data-stu-id="8174a-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="8174a-171">Alterações nos comandos de suplemento, incluindo botões da faixa de opções e itens de menu, não entram em vigor</span><span class="sxs-lookup"><span data-stu-id="8174a-171">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="8174a-172">Se alterações feitas no manifesto, como nomes de arquivo de ícones de botão da faixa de opções ou texto de comandos de suplemento, não parecerem entrar em vigor, experimente limpar o cache do Office no computador.</span><span class="sxs-lookup"><span data-stu-id="8174a-172">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="8174a-173">Para Windows:</span><span class="sxs-lookup"><span data-stu-id="8174a-173">For Windows:</span></span>
<span data-ttu-id="8174a-174">Exclua os conteúdos da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="8174a-174">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="8174a-175">Para Mac:</span><span class="sxs-lookup"><span data-stu-id="8174a-175">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="8174a-176">No iOS:</span><span class="sxs-lookup"><span data-stu-id="8174a-176">For iOS:</span></span>
<span data-ttu-id="8174a-p114">Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="8174a-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="8174a-179">Alterações em arquivos estáticos, como JavaScript, HTML e CSS, não entram em vigor</span><span class="sxs-lookup"><span data-stu-id="8174a-179">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="8174a-180">O navegador pode estar armazenando esses arquivos em cache.</span><span class="sxs-lookup"><span data-stu-id="8174a-180">The browser may be caching these files.</span></span> <span data-ttu-id="8174a-181">Para evitar isso, desative o cache do lado do cliente ao desenvolver.</span><span class="sxs-lookup"><span data-stu-id="8174a-181">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="8174a-182">Os detalhes dependerão do tipo de servidor que você estiver usando.</span><span class="sxs-lookup"><span data-stu-id="8174a-182">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="8174a-183">Na maioria dos casos, envolve adicionar determinados cabeçalhos às respostas HTTP.</span><span class="sxs-lookup"><span data-stu-id="8174a-183">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="8174a-184">Sugerimos o seguinte conjunto:</span><span class="sxs-lookup"><span data-stu-id="8174a-184">We suggest the following set:</span></span>

- <span data-ttu-id="8174a-185">Controle de cache: "privado, sem cache, sem armazenamento"</span><span class="sxs-lookup"><span data-stu-id="8174a-185">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="8174a-186">Pragma: "sem cache"</span><span class="sxs-lookup"><span data-stu-id="8174a-186">Pragma: "no-cache"</span></span>
- <span data-ttu-id="8174a-187">Expira: "-1"</span><span class="sxs-lookup"><span data-stu-id="8174a-187">Expires: "-1"</span></span>

<span data-ttu-id="8174a-188">Para um exemplo de como fazer isso em um servidor Node.JS Express, confira [este arquivo app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span><span class="sxs-lookup"><span data-stu-id="8174a-188">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="8174a-189">Para um exemplo em um projeto ASP.NET, confira [este arquivo cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span><span class="sxs-lookup"><span data-stu-id="8174a-189">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="8174a-190">Se o seu suplemento estiver hospedado no Servidor de Informações da Internet (IIS), você também poderá adicionar o seguinte ao web.config.</span><span class="sxs-lookup"><span data-stu-id="8174a-190">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="8174a-191">Se essas etapas não parecerem funcionar a princípio, talvez seja necessário limpar o cache do navegador.</span><span class="sxs-lookup"><span data-stu-id="8174a-191">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="8174a-192">Faça isso através da interface do usuário do navegador.</span><span class="sxs-lookup"><span data-stu-id="8174a-192">Do this through the UI of the browser.</span></span> <span data-ttu-id="8174a-193">Às vezes, o cache do Microsoft Edge não é limpo com êxito quando você tenta limpá-lo na interface do usuário do Edge.</span><span class="sxs-lookup"><span data-stu-id="8174a-193">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="8174a-194">Se isso acontecer, execute o seguinte comando em um prompt de comando do Windows.</span><span class="sxs-lookup"><span data-stu-id="8174a-194">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="see-also"></a><span data-ttu-id="8174a-195">Confira também</span><span class="sxs-lookup"><span data-stu-id="8174a-195">See also</span></span>

- [<span data-ttu-id="8174a-196">Depurar suplementos no Office na Web</span><span class="sxs-lookup"><span data-stu-id="8174a-196">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="8174a-197">Realizar sideload de um suplemento do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="8174a-197">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="8174a-198">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="8174a-198">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="8174a-199">Extensão do depurador de suplementos do Microsoft Office para o Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="8174a-199">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="8174a-200">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="8174a-200">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="8174a-201">Depurar seu suplemento com o log do tempo de execução</span><span class="sxs-lookup"><span data-stu-id="8174a-201">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
