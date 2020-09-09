---
title: Solucionar erros de desenvolvimento com suplementos do Office
description: Saiba como solucionar erros de desenvolvimento em suplementos do Office.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5801146165446352ec806f6f832e9976f96467ac
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409378"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a><span data-ttu-id="22ac2-103">Solucionar erros de desenvolvimento com suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="22ac2-103">Troubleshoot development errors with Office Add-ins</span></span>

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="22ac2-104">Não é possível carregar o suplemento no painel de tarefas ou outros problemas relacionados ao manifesto do suplemento</span><span class="sxs-lookup"><span data-stu-id="22ac2-104">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="22ac2-105">Confira [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md) e [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md) para depurar problemas de manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="22ac2-105">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="22ac2-106">Alterações nos comandos de suplemento, incluindo botões da faixa de opções e itens de menu, não entram em vigor</span><span class="sxs-lookup"><span data-stu-id="22ac2-106">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="22ac2-107">Se alterações feitas no manifesto, como nomes de arquivo de ícones de botão da faixa de opções ou texto de comandos de suplemento, não parecerem entrar em vigor, experimente limpar o cache do Office no computador.</span><span class="sxs-lookup"><span data-stu-id="22ac2-107">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="22ac2-108">Para Windows:</span><span class="sxs-lookup"><span data-stu-id="22ac2-108">For Windows:</span></span>

<span data-ttu-id="22ac2-109">Exclua o conteúdo da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` e exclua o conteúdo da pasta `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` , se ela existir.</span><span class="sxs-lookup"><span data-stu-id="22ac2-109">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and delete the contents of the folder `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\`, if it exists.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="22ac2-110">Para Mac:</span><span class="sxs-lookup"><span data-stu-id="22ac2-110">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="22ac2-111">No iOS:</span><span class="sxs-lookup"><span data-stu-id="22ac2-111">For iOS:</span></span>
<span data-ttu-id="22ac2-p101">Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="22ac2-p101">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="22ac2-114">Alterações em arquivos estáticos, como JavaScript, HTML e CSS, não entram em vigor</span><span class="sxs-lookup"><span data-stu-id="22ac2-114">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="22ac2-115">O navegador pode estar armazenando esses arquivos em cache.</span><span class="sxs-lookup"><span data-stu-id="22ac2-115">The browser may be caching these files.</span></span> <span data-ttu-id="22ac2-116">Para evitar isso, desative o cache do lado do cliente ao desenvolver.</span><span class="sxs-lookup"><span data-stu-id="22ac2-116">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="22ac2-117">Os detalhes dependerão do tipo de servidor que você estiver usando.</span><span class="sxs-lookup"><span data-stu-id="22ac2-117">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="22ac2-118">Na maioria dos casos, envolve adicionar determinados cabeçalhos às respostas HTTP.</span><span class="sxs-lookup"><span data-stu-id="22ac2-118">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="22ac2-119">Sugerimos o seguinte conjunto:</span><span class="sxs-lookup"><span data-stu-id="22ac2-119">We suggest the following set:</span></span>

- <span data-ttu-id="22ac2-120">Controle de cache: "privado, sem cache, sem armazenamento"</span><span class="sxs-lookup"><span data-stu-id="22ac2-120">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="22ac2-121">Pragma: "sem cache"</span><span class="sxs-lookup"><span data-stu-id="22ac2-121">Pragma: "no-cache"</span></span>
- <span data-ttu-id="22ac2-122">Expira: "-1"</span><span class="sxs-lookup"><span data-stu-id="22ac2-122">Expires: "-1"</span></span>

<span data-ttu-id="22ac2-123">Para um exemplo de como fazer isso em um servidor Node.JS Express, confira [este arquivo app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span><span class="sxs-lookup"><span data-stu-id="22ac2-123">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="22ac2-124">Para um exemplo em um projeto ASP.NET, confira [este arquivo cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span><span class="sxs-lookup"><span data-stu-id="22ac2-124">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="22ac2-125">Se o seu suplemento estiver hospedado no Servidor de Informações da Internet (IIS), você também poderá adicionar o seguinte ao web.config.</span><span class="sxs-lookup"><span data-stu-id="22ac2-125">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="22ac2-126">Se essas etapas não parecerem funcionar a princípio, talvez seja necessário limpar o cache do navegador.</span><span class="sxs-lookup"><span data-stu-id="22ac2-126">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="22ac2-127">Faça isso através da interface do usuário do navegador.</span><span class="sxs-lookup"><span data-stu-id="22ac2-127">Do this through the UI of the browser.</span></span> <span data-ttu-id="22ac2-128">Às vezes, o cache do Microsoft Edge não é limpo com êxito quando você tenta limpá-lo na interface do usuário do Edge.</span><span class="sxs-lookup"><span data-stu-id="22ac2-128">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="22ac2-129">Se isso acontecer, execute o seguinte comando em um prompt de comando do Windows.</span><span class="sxs-lookup"><span data-stu-id="22ac2-129">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a><span data-ttu-id="22ac2-130">As alterações feitas nos valores de propriedade não acontecem e não há mensagem de erro</span><span class="sxs-lookup"><span data-stu-id="22ac2-130">Changes made to property values don't happen and there is no error message</span></span>

<span data-ttu-id="22ac2-131">Verifique a documentação de referência da propriedade para ver se ela é somente leitura.</span><span class="sxs-lookup"><span data-stu-id="22ac2-131">Check the reference documentation for the property to see if it is read only.</span></span> <span data-ttu-id="22ac2-132">Além disso, as [definições do TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura.</span><span class="sxs-lookup"><span data-stu-id="22ac2-132">Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="22ac2-133">Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro.</span><span class="sxs-lookup"><span data-stu-id="22ac2-133">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="22ac2-134">O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id). Consulte também [algumas propriedades não podem ser definidas diretamente](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span><span class="sxs-lookup"><span data-stu-id="22ac2-134">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a><span data-ttu-id="22ac2-135">O suplemento não funciona na borda, mas funciona em outros navegadores</span><span class="sxs-lookup"><span data-stu-id="22ac2-135">Add-in doesn't work on Edge but it works on other browsers</span></span>

<span data-ttu-id="22ac2-136">Consulte [solução de problemas do Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span><span class="sxs-lookup"><span data-stu-id="22ac2-136">See [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span>

## <a name="excel-add-in-throws-errors-but-not-consistently"></a><span data-ttu-id="22ac2-137">O suplemento do Excel gera erros, mas não consistentemente</span><span class="sxs-lookup"><span data-stu-id="22ac2-137">Excel add-in throws errors, but not consistently</span></span>

<span data-ttu-id="22ac2-138">Confira [solucionar problemas de suplementos do Excel](../excel/excel-add-ins-troubleshooting.md) para possíveis causas.</span><span class="sxs-lookup"><span data-stu-id="22ac2-138">See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.</span></span>

## <a name="see-also"></a><span data-ttu-id="22ac2-139">Confira também</span><span class="sxs-lookup"><span data-stu-id="22ac2-139">See also</span></span>

- [<span data-ttu-id="22ac2-140">Depurar suplementos no Office na Web</span><span class="sxs-lookup"><span data-stu-id="22ac2-140">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="22ac2-141">Realizar sideload de um suplemento do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="22ac2-141">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="22ac2-142">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="22ac2-142">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="22ac2-143">Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="22ac2-143">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="22ac2-144">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="22ac2-144">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="22ac2-145">Depurar seu suplemento com o log do tempo de execução</span><span class="sxs-lookup"><span data-stu-id="22ac2-145">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="22ac2-146">Solucionar erros de usuários com Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="22ac2-146">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
