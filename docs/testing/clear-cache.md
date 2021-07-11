---
title: Limpar o cache do Office
description: Saiba como limpar o cache do Office em seu computador.
ms.date: 05/22/2020
localization_priority: Priority
ms.openlocfilehash: 3ead54f3479fdc912f916705fe84645f16392784
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350194"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="96118-103">Limpar o cache do Office</span><span class="sxs-lookup"><span data-stu-id="96118-103">Clear the Office cache</span></span>

<span data-ttu-id="96118-104">Você pode remover um suplemento em que foi feito sideload no Windows, Mac ou iOS limpando o cache do Office em seu computador.</span><span class="sxs-lookup"><span data-stu-id="96118-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span>

<span data-ttu-id="96118-p101">Além disso, se você fizer alterações no manifesto do suplemento (por exemplo, atualizar nomes de arquivos de ícones ou texto de comandos de suplemento), deverá limpar o cache do Office e, em seguida, fazer o sideload do suplemento novamente usando o manifesto atualizado. Isso permitirá que o Office renderize o suplemento conforme descrito pelo manifesto atualizado.</span><span class="sxs-lookup"><span data-stu-id="96118-p101">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest. Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="96118-107">Limpar o cache do Office no Windows</span><span class="sxs-lookup"><span data-stu-id="96118-107">Clear the Office cache on Windows</span></span>

<span data-ttu-id="96118-108">Para remover todos os suplementos carregados do Excel, Word e PowerPoint, exclua o conteúdo da pasta:</span><span class="sxs-lookup"><span data-stu-id="96118-108">To remove all sideloaded add-ins from Excel, Word, and PowerPoint, delete the contents of the folder:</span></span>

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

<span data-ttu-id="96118-109">Se a pasta a seguir existir, exclua seu conteúdo também.</span><span class="sxs-lookup"><span data-stu-id="96118-109">If the following folder exists, delete its contents too.</span></span>

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

<span data-ttu-id="96118-110">Para remover um suplemento sideload do Outlook, use as etapas descritas em [Suplementos Sideload do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md) para localizar o suplemento na seção **Suplementos Personalizados** da caixa de diálogo que lista seus suplementos instalados. Escolha as reticências (`...`) para o suplemento e, em seguida, escolha **Remover** para remover esse suplemento específico.</span><span class="sxs-lookup"><span data-stu-id="96118-110">To remove a sideloaded add-in from Outlook, use the steps outlined in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in and then choose **Remove** to remove that specific add-in.</span></span> <span data-ttu-id="96118-111">Se a remoção do suplemento não funcionar, exclua o conteúdo da pasta `Wef` conforme observado anteriormente para Excel, Word e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="96118-111">If this add-in removal doesn't work, then delete the contents of the `Wef` folder as noted previously for Excel, Word, and PowerPoint.</span></span>

<span data-ttu-id="96118-112">Além disso, para limpar o cache do Office no Windows 10 quando o suplemento estiver sendo executado no Microsoft Edge, você pode usar o Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="96118-112">Additionally, to clear the Office cache on Windows 10 when the add-in is running in Microsoft Edge, you can use the Microsoft Edge DevTools.</span></span>

> [!TIP]
> <span data-ttu-id="96118-113">Se você deseja apenas que o suplemento sideloaded reflita as alterações recentes em seus arquivos de origem HTML ou JavaScript, não deve ser necessário limpar o cache.</span><span class="sxs-lookup"><span data-stu-id="96118-113">If you only want the sideloaded add-in to reflect recent changes to its HTML or JavaScript source files, you shouldn't need to clear the cache.</span></span> <span data-ttu-id="96118-114">Em vez disso, coloque o foco no painel de tarefas do suplemento (clicando em qualquer lugar no painel de tarefas) e, em seguida, pressione **F5** para recarregar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="96118-114">Instead, just put focus in the add-in's task pane (by clicking anywhere within the task pane) and then press **F5** to reload the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="96118-115">Para limpar o cache do Office usando as etapas a seguir, seu suplemento deve ter um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="96118-115">To clear the Office cache using the following steps, your add-in must have a task pane.</span></span> <span data-ttu-id="96118-116">Se o seu suplemento for um suplemento sem interface de usuário, por exemplo, um que use o recurso [em envio](../outlook/outlook-on-send-addins.md), você precisará adicionar um painel de tarefas ao seu suplemento que use o mesmo domínio para [SourceLocation](../reference/manifest/sourcelocation.md), antes de poder usar as etapas a seguir para limpar o cache.</span><span class="sxs-lookup"><span data-stu-id="96118-116">If your add-in is a UI-less add-in -- for example, one that uses the [on-send](../outlook/outlook-on-send-addins.md) feature -- you'll need to add a task pane to your add-in that uses the same domain for [SourceLocation](../reference/manifest/sourcelocation.md), before you can use the following steps to clear the cache.</span></span>

1. <span data-ttu-id="96118-117">Instalar o [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).</span><span class="sxs-lookup"><span data-stu-id="96118-117">Install the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).</span></span>

2. <span data-ttu-id="96118-118">Abra seu suplemento no cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="96118-118">Open your add-in in the Office client.</span></span>

3. <span data-ttu-id="96118-119">Execute o Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="96118-119">Run the Microsoft Edge DevTools.</span></span>

4. <span data-ttu-id="96118-120">No Microsoft Edge DevTools, abra a guia **Local**. Seu suplemento será listado por nome.</span><span class="sxs-lookup"><span data-stu-id="96118-120">In the Microsoft Edge DevTools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

5. <span data-ttu-id="96118-121">Selecione o nome do suplemento para anexar o depurador ao seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="96118-121">Select the add-in name to attach the debugger to your add-in.</span></span> <span data-ttu-id="96118-122">Uma nova janela do Microsoft Edge DevTools será aberta quando o depurador for anexado ao seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="96118-122">A new Microsoft Edge DevTools window will open when the debugger attaches to your add-in.</span></span>

6. <span data-ttu-id="96118-123">Na guia **Network** da nova janela, selecione o botão **Limpar cache**.</span><span class="sxs-lookup"><span data-stu-id="96118-123">On the **Network** tab of the new window, select the **Clear cache** button.</span></span>

    ![Captura de tela do Microsoft Edge DevTools com o botão Limpar cache realçado.](../images/edge-devtools-clear-cache.png)

7. <span data-ttu-id="96118-125">Se concluir essas etapas não produzir o resultado desejado, você também pode selecionar o botão **Sempre atualizar do servidor**.</span><span class="sxs-lookup"><span data-stu-id="96118-125">If completing these steps doesn't produce the desired result, you can also select the **Always refresh from server** button.</span></span>

    ![Captura de tela do Microsoft Edge DevTools com o botão sempre atualizar do servidor realçado.](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="96118-127">Limpar o cache do Office no Mac</span><span class="sxs-lookup"><span data-stu-id="96118-127">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="96118-128">Limpar o cache do Office no iOS</span><span class="sxs-lookup"><span data-stu-id="96118-128">Clear the Office cache on iOS</span></span>

<span data-ttu-id="96118-129">Para limpar o cache do Office no iOS, chame `window.location.reload(true)` a partir do JavaScript no suplemento para forçar um recarregamento.</span><span class="sxs-lookup"><span data-stu-id="96118-129">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="96118-130">Uma outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="96118-130">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="96118-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="96118-131">See also</span></span>

- [<span data-ttu-id="96118-132">Depurar suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="96118-132">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [<span data-ttu-id="96118-133">Depurar seu suplemento com o log do tempo de execução</span><span class="sxs-lookup"><span data-stu-id="96118-133">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="96118-134">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="96118-134">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="96118-135">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="96118-135">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="96118-136">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="96118-136">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
