---
title: Limpar o cache do Office
description: Saiba como limpar o cache do Office em seu computador.
ms.date: 05/22/2020
localization_priority: Normal
ms.openlocfilehash: fb08cd3838cf9c1d12df2fc8d37453776a05c768
ms.sourcegitcommit: d88b3dcfe13ba06f821b55db2de46aed152a378d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/22/2020
ms.locfileid: "44347648"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="9d973-103">Limpar o cache do Office</span><span class="sxs-lookup"><span data-stu-id="9d973-103">Clear the Office cache</span></span>

<span data-ttu-id="9d973-104">Você pode remover um suplemento em que foi feito sideload no Windows, Mac ou iOS limpando o cache do Office em seu computador.</span><span class="sxs-lookup"><span data-stu-id="9d973-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span>

<span data-ttu-id="9d973-105">Além disso, se você fizer alterações no manifesto do seu suplemento (por exemplo, atualizar nomes de arquivos de ícones ou texto de comandos de suplemento), você deve limpar o cache do Office e, em seguida, fazer o sideload novamente usando o manifesto atualizado.</span><span class="sxs-lookup"><span data-stu-id="9d973-105">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="9d973-106">Isso permitirá que o Office processe o suplemento conforme descrito no manifesto atualizado.</span><span class="sxs-lookup"><span data-stu-id="9d973-106">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="9d973-107">Limpar o cache do Office no Windows</span><span class="sxs-lookup"><span data-stu-id="9d973-107">Clear the Office cache on Windows</span></span>

<span data-ttu-id="9d973-108">Para remover todos os suplementos do suplementos foi feito do Excel, Word e PowerPoint, exclua o conteúdo da pasta:</span><span class="sxs-lookup"><span data-stu-id="9d973-108">To remove all sideloaded add-ins from Excel, Word, and PowerPoint, delete the contents of the folder:</span></span>

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

<span data-ttu-id="9d973-109">Para remover um suplemento do suplementos foi feito do Outlook, use as etapas descritas em suplementos do [Sideload do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md) para localizar o suplemento na seção **suplementos personalizados** da caixa de diálogo que lista seus suplementos instalados. escolha as reticências ( `...` ) para o suplemento e, em seguida, escolha **remover** para remover o suplemento específico do.</span><span class="sxs-lookup"><span data-stu-id="9d973-109">To remove a sideloaded add-in from Outlook, use the steps outlined in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in and then choose **Remove** to remove that specific add-in.</span></span> <span data-ttu-id="9d973-110">Se essa remoção de suplemento não funcionar, exclua o conteúdo da `Wef` pasta conforme indicado anteriormente para Excel, Word e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="9d973-110">If this add-in removal doesn't work, then delete the contents of the `Wef` folder as noted previously for Excel, Word, and PowerPoint.</span></span>

<span data-ttu-id="9d973-111">Além disso, para limpar o cache do Office no Windows 10 quando o suplemento estiver sendo executado no Microsoft Edge, você pode usar o Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="9d973-111">Additionally, to clear the Office cache on Windows 10 when the add-in is running in Microsoft Edge, you can use the Microsoft Edge DevTools.</span></span>

> [!TIP]
> <span data-ttu-id="9d973-112">Se você quer apenas que o suplemento sideloaded reflita alterações recentes em seus arquivos de origem HTML ou JavaScript, não deverá ser necessário usar as etapas a seguir para limpar o cache.</span><span class="sxs-lookup"><span data-stu-id="9d973-112">If you're just wanting the sideloaded add-in to reflect recent changes to its HTML or JavaScript source files, you shouldn't need to use the following steps to clear the cache.</span></span> <span data-ttu-id="9d973-113">Em vez disso, coloque o foco no painel de tarefas do suplemento (clicando em qualquer lugar no painel de tarefas) e, em seguida, pressione **F5** para recarregar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="9d973-113">Instead, just put focus in the add-in's task pane (by clicking anywhere within the task pane) and then press **F5** to reload the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9d973-114">Para limpar o cache do Office usando as etapas a seguir, seu suplemento deve ter um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="9d973-114">To clear the Office cache using the following steps, your add-in must have a task pane.</span></span> <span data-ttu-id="9d973-115">Se o seu suplemento for um suplemento sem interface de usuário, por exemplo, um que use o recurso [em envio](../outlook/outlook-on-send-addins.md), você precisará adicionar um painel de tarefas ao seu suplemento que use o mesmo domínio para [SourceLocation](../reference/manifest/sourcelocation.md), antes de poder usar as etapas a seguir para limpar o cache.</span><span class="sxs-lookup"><span data-stu-id="9d973-115">If your add-in is a UI-less add-in -- for example, one that uses the [on-send](../outlook/outlook-on-send-addins.md) feature -- you'll need to add a task pane to your add-in that uses the same domain for [SourceLocation](../reference/manifest/sourcelocation.md), before you can use the following steps to clear the cache.</span></span>

1. <span data-ttu-id="9d973-116">Instalar o [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).</span><span class="sxs-lookup"><span data-stu-id="9d973-116">Install the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).</span></span>

2. <span data-ttu-id="9d973-117">Abra seu suplemento no cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="9d973-117">Open your add-in in the Office client.</span></span>

3. <span data-ttu-id="9d973-118">Execute o Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="9d973-118">Run the Microsoft Edge DevTools.</span></span>

4. <span data-ttu-id="9d973-119">No Microsoft Edge DevTools, abra a guia **Local**. Seu suplemento será listado por nome.</span><span class="sxs-lookup"><span data-stu-id="9d973-119">In the Microsoft Edge DevTools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

5. <span data-ttu-id="9d973-120">Selecione o nome do suplemento para anexar o depurador ao seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="9d973-120">Select the add-in name to attach the debugger to your add-in.</span></span> <span data-ttu-id="9d973-121">Uma nova janela do Microsoft Edge DevTools será aberta quando o depurador for anexado ao seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="9d973-121">A new Microsoft Edge DevTools window will open when the debugger attaches to your add-in.</span></span>

6. <span data-ttu-id="9d973-122">Na guia **Network** da nova janela, selecione o botão **Limpar cache**.</span><span class="sxs-lookup"><span data-stu-id="9d973-122">On the **Network** tab of the new window, select the **Clear cache** button.</span></span>

    ![Captura de tela do Microsoft Edge DevTools com o botão Limpar cache realçado](../images/edge-devtools-clear-cache.png)

7. <span data-ttu-id="9d973-124">Se concluir essas etapas não produzir o resultado desejado, você também pode selecionar o botão **Sempre atualizar do servidor**.</span><span class="sxs-lookup"><span data-stu-id="9d973-124">If completing these steps doesn't produce the desired result, you can also select the **Always refresh from server** button.</span></span>

    ![Captura de tela do Microsoft Edge DevTools com o botão sempre atualizar do servidor realçado](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="9d973-126">Limpar o cache do Office no Mac</span><span class="sxs-lookup"><span data-stu-id="9d973-126">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="9d973-127">Limpar o cache do Office no iOS</span><span class="sxs-lookup"><span data-stu-id="9d973-127">Clear the Office cache on iOS</span></span>

<span data-ttu-id="9d973-128">Para limpar o cache do Office no iOS, chame `window.location.reload(true)` a partir do JavaScript no suplemento para forçar um recarregamento.</span><span class="sxs-lookup"><span data-stu-id="9d973-128">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="9d973-129">Uma outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="9d973-129">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="9d973-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="9d973-130">See also</span></span>

- [<span data-ttu-id="9d973-131">Depurar suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="9d973-131">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [<span data-ttu-id="9d973-132">Depurar seu suplemento com o log do tempo de execução</span><span class="sxs-lookup"><span data-stu-id="9d973-132">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="9d973-133">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="9d973-133">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="9d973-134">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="9d973-134">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="9d973-135">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="9d973-135">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
