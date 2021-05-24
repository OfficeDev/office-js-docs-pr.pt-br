---
title: Depurar seu Outlook de eventos (visualização)
description: Saiba como depurar seu Outlook que implementa a ativação baseada em eventos.
ms.topic: article
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: d7621a7407db3b8e773d1534beb6c881f7b48558
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555250"
---
# <a name="debug-your-event-based-outlook-add-in-preview"></a><span data-ttu-id="3a4e1-103">Depurar seu Outlook de eventos (visualização)</span><span class="sxs-lookup"><span data-stu-id="3a4e1-103">Debug your event-based Outlook add-in (preview)</span></span>

<span data-ttu-id="3a4e1-104">Este artigo fornece orientações de depuração à medida que você implementa a ativação baseada em [eventos](autolaunch.md) no seu complemento.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-104">This article provides debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in.</span></span> <span data-ttu-id="3a4e1-105">O recurso de ativação baseada em evento está atualmente na visualização.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-105">The event-based activation feature is currently in preview.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3a4e1-106">Esse recurso de depuração só é suportado para visualização no Outlook no Windows com uma assinatura Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-106">This debugging capability is only supported for preview in Outlook on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="3a4e1-107">Para obter mais informações, consulte a [seção Visualização de depuração do recurso de ativação](#preview-debugging-for-the-event-based-activation-feature) baseada em evento neste artigo.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-107">For more information, see the [Preview debugging for the event-based activation feature](#preview-debugging-for-the-event-based-activation-feature) section in this article.</span></span>

<span data-ttu-id="3a4e1-108">Neste artigo, abordamos os principais estágios para habilitar a depuração.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-108">In this article, we discuss the key stages to enable debugging.</span></span>

- [<span data-ttu-id="3a4e1-109">Marcar o complemento para depuração</span><span class="sxs-lookup"><span data-stu-id="3a4e1-109">Mark the add-in for debugging</span></span>](#mark-your-add-in-for-debugging)
- [<span data-ttu-id="3a4e1-110">Configurar Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="3a4e1-110">Configure Visual Studio Code</span></span>](#configure-visual-studio-code)
- [<span data-ttu-id="3a4e1-111">Anexar Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="3a4e1-111">Attach Visual Studio Code</span></span>](#attach-visual-studio-code)
- [<span data-ttu-id="3a4e1-112">Depuração</span><span class="sxs-lookup"><span data-stu-id="3a4e1-112">Debug</span></span>](#debug)

<span data-ttu-id="3a4e1-113">Você tem várias opções para criar seu projeto de complemento.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-113">You have several options for creating your add-in project.</span></span> <span data-ttu-id="3a4e1-114">Dependendo da opção que você está usando, as etapas podem variar.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-114">Depending on the option you're using, the steps may vary.</span></span> <span data-ttu-id="3a4e1-115">Nesse caso, se você usou o gerador Yeoman para os complementos do Office para criar seu projeto de complemento (por exemplo, fazendo o passo a passo de ativação baseada em eventos [),](autolaunch.md)siga as etapas **yo office,** caso contrário, siga as **outras** etapas.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-115">Where this is the case, if you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by doing the [event-based activation walkthrough](autolaunch.md)), then follow the **yo office** steps, otherwise follow the **Other** steps.</span></span> <span data-ttu-id="3a4e1-116">Visual Studio Code deve ser pelo menos a versão 1.56.1.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-116">Visual Studio Code should be at least version 1.56.1.</span></span>

## <a name="preview-debugging-for-the-event-based-activation-feature"></a><span data-ttu-id="3a4e1-117">Visualização de depuração para o recurso de ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="3a4e1-117">Preview debugging for the event-based activation feature</span></span>

<span data-ttu-id="3a4e1-118">Convidamos você a experimentar o recurso de depuração para o recurso de ativação baseada em evento!</span><span class="sxs-lookup"><span data-stu-id="3a4e1-118">We invite you to try out the debugging capability for the event-based activation feature!</span></span> <span data-ttu-id="3a4e1-119">Deixe-nos saber seus cenários e como podemos melhorar nos dando feedback por meio GitHub (consulte a seção **Comentários** no final desta página).</span><span class="sxs-lookup"><span data-stu-id="3a4e1-119">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="3a4e1-120">Para visualizar esse recurso para Outlook no Windows, o build mínimo necessário é 16.0.13729.20000.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-120">To preview this capability for Outlook on Windows, the minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="3a4e1-121">Para acessar as Office beta, participe do programa [Office Insider](https://insider.office.com).</span><span class="sxs-lookup"><span data-stu-id="3a4e1-121">For access to Office beta builds, join the [Office Insider program](https://insider.office.com).</span></span>

## <a name="mark-your-add-in-for-debugging"></a><span data-ttu-id="3a4e1-122">Marcar seu complemento para depuração</span><span class="sxs-lookup"><span data-stu-id="3a4e1-122">Mark your add-in for debugging</span></span>

1. <span data-ttu-id="3a4e1-123">De definir a chave do Registro `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .</span><span class="sxs-lookup"><span data-stu-id="3a4e1-123">Set the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span> <span data-ttu-id="3a4e1-124">`[Add-in ID]` é **a ID** no manifesto do complemento.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-124">`[Add-in ID]` is the **Id** in the add-in manifest.</span></span>

    <span data-ttu-id="3a4e1-125">**yo office**: em uma janela de linha de comando, navegue até a raiz da pasta do seu complemento e execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-125">**yo office**: In a command line window, navigate to the root of your add-in folder then run the following command.</span></span>

    ```command&nbsp;line
    npm start
    ```

    <span data-ttu-id="3a4e1-126">Além de criar o código e iniciar o servidor local, esse comando deve definir a chave do Registro para esse `UseDirectDebugger` complemento como `1` .</span><span class="sxs-lookup"><span data-stu-id="3a4e1-126">In addition to building the code and starting the local server, this command should set the `UseDirectDebugger` registry key for this add-in to `1`.</span></span>

    <span data-ttu-id="3a4e1-127">**Outros**: Adicione a `UseDirectDebugger` chave do Registro em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` .</span><span class="sxs-lookup"><span data-stu-id="3a4e1-127">**Other**: Add the `UseDirectDebugger` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`.</span></span> <span data-ttu-id="3a4e1-128">Substitua `[Add-in ID]` pela **ID do** manifesto do complemento.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-128">Replace `[Add-in ID]` with the **Id** from the add-in manifest.</span></span> <span data-ttu-id="3a4e1-129">De definir a chave do Registro como `1` .</span><span class="sxs-lookup"><span data-stu-id="3a4e1-129">Set the registry key to `1`.</span></span>

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. <span data-ttu-id="3a4e1-130">Inicie Outlook desktop (ou reinicie Outlook se ele já estiver aberto).</span><span class="sxs-lookup"><span data-stu-id="3a4e1-130">Start Outlook desktop (or restart Outlook if it's already open).</span></span>
1. <span data-ttu-id="3a4e1-131">Componha uma nova mensagem ou compromisso.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-131">Compose a new message or appointment.</span></span> <span data-ttu-id="3a4e1-132">Você deve ver a caixa de diálogo a seguir.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-132">You should see the following dialog.</span></span> <span data-ttu-id="3a4e1-133">Não *interaja* com a caixa de diálogo ainda.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-133">Do *not* interact with the dialog yet.</span></span>

    ![Captura de tela da caixa de diálogo de manipulador baseado em evento de depuração](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a><span data-ttu-id="3a4e1-135">Configurar Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="3a4e1-135">Configure Visual Studio Code</span></span>

### <a name="yo-office"></a><span data-ttu-id="3a4e1-136">yo office</span><span class="sxs-lookup"><span data-stu-id="3a4e1-136">yo office</span></span>

1. <span data-ttu-id="3a4e1-137">Na janela de linha de comando, abra Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-137">Back in the command line window, open Visual Studio Code.</span></span>

    ```command&nbsp;line
    code .
    ```

1. <span data-ttu-id="3a4e1-138">Em Visual Studio Code, abra o arquivo **./.vscode/launch.jse** adicione o seguinte trecho à sua lista de configurações.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-138">In Visual Studio Code, open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="3a4e1-139">Salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-139">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a><span data-ttu-id="3a4e1-140">Outros</span><span class="sxs-lookup"><span data-stu-id="3a4e1-140">Other</span></span>

1. <span data-ttu-id="3a4e1-141">Crie uma nova pasta chamada **Depuração** (talvez na pasta **Desktop).**</span><span class="sxs-lookup"><span data-stu-id="3a4e1-141">Create a new folder called **Debugging** (perhaps in your **Desktop** folder).</span></span>
1. <span data-ttu-id="3a4e1-142">Abra Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-142">Open Visual Studio Code.</span></span>
1. <span data-ttu-id="3a4e1-143">Vá para **Pasta**  >  **Aberta do Arquivo,** navegue até a pasta que você acabou de criar e escolha **Selecionar Pasta**.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-143">Go to **File** > **Open Folder**, navigate to the folder you just created, then choose **Select Folder**.</span></span>
1. <span data-ttu-id="3a4e1-144">Na Barra de Atividades, selecione o item **Depurar** (Ctrl+Shift+D).</span><span class="sxs-lookup"><span data-stu-id="3a4e1-144">On the Activity Bar, select the **Debug** item (Ctrl+Shift+D).</span></span>

    ![Captura de tela do ícone de depuração na Barra de Atividades](../images/vs-code-debug.png)

1. <span data-ttu-id="3a4e1-146">Selecione o **link criar um launch.jsno** arquivo.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-146">Select the **create a launch.json file** link.</span></span>

    ![Captura de tela do link para criar um arquivo launch.jsno Visual Studio Code](../images/vs-code-create-launch.json.png)

1. <span data-ttu-id="3a4e1-148">Na lista **suspenso Selecionar Ambiente,** selecione **Borda: Iniciar** para criar um launch.jsno arquivo.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-148">In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.</span></span>
1. <span data-ttu-id="3a4e1-149">Adicione o trecho a seguir à sua lista de configurações.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-149">Add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="3a4e1-150">Salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-150">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a><span data-ttu-id="3a4e1-151">Anexar Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="3a4e1-151">Attach Visual Studio Code</span></span>

1. <span data-ttu-id="3a4e1-152">Para encontrar o nome dobundle.js **do** bundle.js, abra a seguinte pasta no Windows Explorer e pesquise a **ID** do seu complemento (encontrada no manifesto).</span><span class="sxs-lookup"><span data-stu-id="3a4e1-152">To find the add-in's **bundle.js**, open the following folder in Windows Explorer and search for your add-in's **Id** (found in the manifest).</span></span>

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    <span data-ttu-id="3a4e1-153">Abra a pasta prefixada com essa ID e copie seu caminho completo.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-153">Open the folder prefixed with this ID and copy its full path.</span></span> <span data-ttu-id="3a4e1-154">Em Visual Studio Code, abra **bundle.js** dessa pasta.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-154">In Visual Studio Code, open **bundle.js** from that folder.</span></span> <span data-ttu-id="3a4e1-155">O padrão do caminho do arquivo deve ser o seguinte:</span><span class="sxs-lookup"><span data-stu-id="3a4e1-155">The pattern of the file path should be as follows:</span></span>

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. <span data-ttu-id="3a4e1-156">Coloque pontos de interrupção bundle.js onde você deseja que o depurador pare.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-156">Place breakpoints in bundle.js where you want the debugger to stop.</span></span>
1. <span data-ttu-id="3a4e1-157">No menu **suspenso DEPURar,** selecione o nome **Depuração Direta** e, em seguida, selecione **Executar**.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-157">In the **DEBUG** dropdown, select the name **Direct Debugging**, then select **Run**.</span></span>

    ![Captura de tela da seleção de Depuração Direta de opções de configuração no menu suspenso Visual Studio Code Depuração](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a><span data-ttu-id="3a4e1-159">Depuração</span><span class="sxs-lookup"><span data-stu-id="3a4e1-159">Debug</span></span>

1. <span data-ttu-id="3a4e1-160">Depois de confirmar se o depurador está anexado, retorne ao Outlook e, na caixa de diálogo Manipulador baseado em Evento de **Depuração,** escolha **OK** .</span><span class="sxs-lookup"><span data-stu-id="3a4e1-160">After confirming that the debugger is attached, return to Outlook, and in the **Debug Event-based handler** dialog, choose **OK** .</span></span>

1. <span data-ttu-id="3a4e1-161">Agora você pode atingir seus pontos de interrupção Visual Studio Code, permitindo que você depure seu código de ativação baseado em evento.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-161">You can now hit your breakpoints in Visual Studio Code, enabling you to debug your event-based activation code.</span></span>

## <a name="stop-debugging"></a><span data-ttu-id="3a4e1-162">Parar a depuração</span><span class="sxs-lookup"><span data-stu-id="3a4e1-162">Stop debugging</span></span>

<span data-ttu-id="3a4e1-163">Para interromper a depuração para o restante da sessão Outlook da área de trabalho atual, na caixa de diálogo Manipulador baseado em Evento de **Depuração,** escolha **Cancelar**.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-163">To stop debugging for the rest of the current Outlook desktop session, in the **Debug Event-based handler** dialog, choose **Cancel**.</span></span> <span data-ttu-id="3a4e1-164">Para habilitar novamente a depuração, reinicie Outlook área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="3a4e1-164">To re-enable debugging, restart Outlook desktop.</span></span>

<span data-ttu-id="3a4e1-165">Para impedir que a caixa **de** diálogo de manipulador baseada em Evento de depuração seja exibida e pare a depuração para sessões Outlook posteriores, exclua a chave do Registro associada ou desfaça seu valor como `0` : `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .</span><span class="sxs-lookup"><span data-stu-id="3a4e1-165">To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent Outlook sessions, delete the associated registry key or set its value to `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span>

## <a name="see-also"></a><span data-ttu-id="3a4e1-166">Confira também</span><span class="sxs-lookup"><span data-stu-id="3a4e1-166">See also</span></span>

- [<span data-ttu-id="3a4e1-167">Configurar seu Outlook para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="3a4e1-167">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
- [<span data-ttu-id="3a4e1-168">Depurar seu suplemento com o log do tempo de execução</span><span class="sxs-lookup"><span data-stu-id="3a4e1-168">Debug your add-in with runtime logging</span></span>](../testing/runtime-logging.md#runtime-logging-on-windows)
