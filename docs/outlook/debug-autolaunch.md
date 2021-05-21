---
title: Depurar seu complemento Outlook baseado em eventos (pré-visualização)
description: Aprenda a depurar seu Outlook complemento que implementa ativação baseada em eventos.
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
# <a name="debug-your-event-based-outlook-add-in-preview"></a><span data-ttu-id="f1c9b-103">Depurar seu complemento Outlook baseado em eventos (pré-visualização)</span><span class="sxs-lookup"><span data-stu-id="f1c9b-103">Debug your event-based Outlook add-in (preview)</span></span>

<span data-ttu-id="f1c9b-104">Este artigo fornece orientação de depuração à medida que você implementa [a ativação baseada](autolaunch.md) em eventos em seu complemento.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-104">This article provides debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in.</span></span> <span data-ttu-id="f1c9b-105">O recurso de ativação baseado em eventos está atualmente em pré-visualização.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-105">The event-based activation feature is currently in preview.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f1c9b-106">Esse recurso de depuração só é suportado para visualização em Outlook em Windows com uma assinatura Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-106">This debugging capability is only supported for preview in Outlook on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="f1c9b-107">Para obter mais informações, consulte a [depuração do Preview para a](#preview-debugging-for-the-event-based-activation-feature) seção de recursos de ativação baseada em eventos neste artigo.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-107">For more information, see the [Preview debugging for the event-based activation feature](#preview-debugging-for-the-event-based-activation-feature) section in this article.</span></span>

<span data-ttu-id="f1c9b-108">Neste artigo, discutimos as etapas-chave para permitir a depuração.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-108">In this article, we discuss the key stages to enable debugging.</span></span>

- [<span data-ttu-id="f1c9b-109">Marque o complemento para depuração</span><span class="sxs-lookup"><span data-stu-id="f1c9b-109">Mark the add-in for debugging</span></span>](#mark-your-add-in-for-debugging)
- [<span data-ttu-id="f1c9b-110">Configure Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="f1c9b-110">Configure Visual Studio Code</span></span>](#configure-visual-studio-code)
- [<span data-ttu-id="f1c9b-111">Anexar Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="f1c9b-111">Attach Visual Studio Code</span></span>](#attach-visual-studio-code)
- [<span data-ttu-id="f1c9b-112">depurar</span><span class="sxs-lookup"><span data-stu-id="f1c9b-112">Debug</span></span>](#debug)

<span data-ttu-id="f1c9b-113">Você tem várias opções para criar seu projeto de complemento.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-113">You have several options for creating your add-in project.</span></span> <span data-ttu-id="f1c9b-114">Dependendo da opção que você está usando, as etapas podem variar.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-114">Depending on the option you're using, the steps may vary.</span></span> <span data-ttu-id="f1c9b-115">Onde este é o caso, se você usou o gerador Yeoman para Office Add-ins para criar seu projeto de complementação (por exemplo, fazendo o [passo a passo de ativação baseado](autolaunch.md)em eventos ), então siga as etapas do **escritório,** siga as **outras** etapas.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-115">Where this is the case, if you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by doing the [event-based activation walkthrough](autolaunch.md)), then follow the **yo office** steps, otherwise follow the **Other** steps.</span></span> <span data-ttu-id="f1c9b-116">Visual Studio Code deve ser pelo menos a versão 1.56.1.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-116">Visual Studio Code should be at least version 1.56.1.</span></span>

## <a name="preview-debugging-for-the-event-based-activation-feature"></a><span data-ttu-id="f1c9b-117">Depuração de visualização para o recurso de ativação baseado em eventos</span><span class="sxs-lookup"><span data-stu-id="f1c9b-117">Preview debugging for the event-based activation feature</span></span>

<span data-ttu-id="f1c9b-118">Convidamos você a experimentar o recurso de depuração para o recurso de ativação baseado em eventos!</span><span class="sxs-lookup"><span data-stu-id="f1c9b-118">We invite you to try out the debugging capability for the event-based activation feature!</span></span> <span data-ttu-id="f1c9b-119">Deixe-nos saber seus cenários e como podemos melhorar nos dando feedback através de GitHub (veja a seção **Feedback** no final desta página).</span><span class="sxs-lookup"><span data-stu-id="f1c9b-119">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="f1c9b-120">Para visualizar essa capacidade para Outlook em Windows, a compilação mínima necessária é de 16.0.13729.20000.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-120">To preview this capability for Outlook on Windows, the minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="f1c9b-121">Para ter acesso a Office compilações beta, participe do [programa Office Insider](https://insider.office.com).</span><span class="sxs-lookup"><span data-stu-id="f1c9b-121">For access to Office beta builds, join the [Office Insider program](https://insider.office.com).</span></span>

## <a name="mark-your-add-in-for-debugging"></a><span data-ttu-id="f1c9b-122">Marque seu complemento para depuração</span><span class="sxs-lookup"><span data-stu-id="f1c9b-122">Mark your add-in for debugging</span></span>

1. <span data-ttu-id="f1c9b-123">Defina a chave de registro `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .</span><span class="sxs-lookup"><span data-stu-id="f1c9b-123">Set the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span> <span data-ttu-id="f1c9b-124">`[Add-in ID]` é o **ID** no manifesto add-in.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-124">`[Add-in ID]` is the **Id** in the add-in manifest.</span></span>

    <span data-ttu-id="f1c9b-125">**yo office**: Em uma janela de linha de comando, navegue até a raiz da pasta de complementação e execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-125">**yo office**: In a command line window, navigate to the root of your add-in folder then run the following command.</span></span>

    ```command&nbsp;line
    npm start
    ```

    <span data-ttu-id="f1c9b-126">Além de construir o código e iniciar o servidor local, este comando deve definir a `UseDirectDebugger` chave de registro para este complemento. `1`</span><span class="sxs-lookup"><span data-stu-id="f1c9b-126">In addition to building the code and starting the local server, this command should set the `UseDirectDebugger` registry key for this add-in to `1`.</span></span>

    <span data-ttu-id="f1c9b-127">**Outros:** Adicione a `UseDirectDebugger` chave de registro em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` .</span><span class="sxs-lookup"><span data-stu-id="f1c9b-127">**Other**: Add the `UseDirectDebugger` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`.</span></span> <span data-ttu-id="f1c9b-128">Substitua `[Add-in ID]` pelo **ID** do manifesto de complemento.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-128">Replace `[Add-in ID]` with the **Id** from the add-in manifest.</span></span> <span data-ttu-id="f1c9b-129">Defina a chave de registro para `1` .</span><span class="sxs-lookup"><span data-stu-id="f1c9b-129">Set the registry key to `1`.</span></span>

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. <span data-ttu-id="f1c9b-130">Inicie Outlook desktop (ou reinicie Outlook se já estiver aberto).</span><span class="sxs-lookup"><span data-stu-id="f1c9b-130">Start Outlook desktop (or restart Outlook if it's already open).</span></span>
1. <span data-ttu-id="f1c9b-131">Componha uma nova mensagem ou nomeação.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-131">Compose a new message or appointment.</span></span> <span data-ttu-id="f1c9b-132">Você deve ver o seguinte diálogo.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-132">You should see the following dialog.</span></span> <span data-ttu-id="f1c9b-133">*Ainda não* interaja com o diálogo.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-133">Do *not* interact with the dialog yet.</span></span>

    ![Captura de tela da caixa de diálogo do manipulador baseado em eventos Debug](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a><span data-ttu-id="f1c9b-135">Configure Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="f1c9b-135">Configure Visual Studio Code</span></span>

### <a name="yo-office"></a><span data-ttu-id="f1c9b-136">yo escritório</span><span class="sxs-lookup"><span data-stu-id="f1c9b-136">yo office</span></span>

1. <span data-ttu-id="f1c9b-137">De volta à janela da linha de comando, abra Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-137">Back in the command line window, open Visual Studio Code.</span></span>

    ```command&nbsp;line
    code .
    ```

1. <span data-ttu-id="f1c9b-138">Em Visual Studio Code, abra o arquivo **./.vscode/launch.js** e adicione o trecho a seguir à sua lista de configurações.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-138">In Visual Studio Code, open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="f1c9b-139">Salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-139">Save your changes.</span></span>

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

### <a name="other"></a><span data-ttu-id="f1c9b-140">Outros</span><span class="sxs-lookup"><span data-stu-id="f1c9b-140">Other</span></span>

1. <span data-ttu-id="f1c9b-141">Crie uma nova pasta chamada **Depuração** (talvez na pasta **Desktop).**</span><span class="sxs-lookup"><span data-stu-id="f1c9b-141">Create a new folder called **Debugging** (perhaps in your **Desktop** folder).</span></span>
1. <span data-ttu-id="f1c9b-142">Abra Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-142">Open Visual Studio Code.</span></span>
1. <span data-ttu-id="f1c9b-143">Vá para  >  **File Open Folder**, navegue até a pasta que você acabou de criar e escolha Selecionar **pasta**.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-143">Go to **File** > **Open Folder**, navigate to the folder you just created, then choose **Select Folder**.</span></span>
1. <span data-ttu-id="f1c9b-144">Na Barra de Atividades, selecione o item **Depuração** (Ctrl+Shift+D).</span><span class="sxs-lookup"><span data-stu-id="f1c9b-144">On the Activity Bar, select the **Debug** item (Ctrl+Shift+D).</span></span>

    ![Captura de tela do ícone Debug na Barra de Atividades](../images/vs-code-debug.png)

1. <span data-ttu-id="f1c9b-146">Selecione **a criação de um launch.jsno** link de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-146">Select the **create a launch.json file** link.</span></span>

    ![Captura de tela do link para criar um launch.jsno arquivo em Visual Studio Code](../images/vs-code-create-launch.json.png)

1. <span data-ttu-id="f1c9b-148">Na lista suspensa **do Ambiente Select,** selecione **Borda: Inicie** para criar uma launch.jsno arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-148">In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.</span></span>
1. <span data-ttu-id="f1c9b-149">Adicione o trecho a seguir à sua lista de configurações.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-149">Add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="f1c9b-150">Salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-150">Save your changes.</span></span>

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

## <a name="attach-visual-studio-code"></a><span data-ttu-id="f1c9b-151">Anexar Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="f1c9b-151">Attach Visual Studio Code</span></span>

1. <span data-ttu-id="f1c9b-152">Para encontrar o **bundle.js** do complemento, abra a seguinte pasta no Windows Explorer e pesquise o ID do seu **complemento** (encontrado no manifesto).</span><span class="sxs-lookup"><span data-stu-id="f1c9b-152">To find the add-in's **bundle.js**, open the following folder in Windows Explorer and search for your add-in's **Id** (found in the manifest).</span></span>

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    <span data-ttu-id="f1c9b-153">Abra a pasta prefixada com este ID e copie seu caminho completo.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-153">Open the folder prefixed with this ID and copy its full path.</span></span> <span data-ttu-id="f1c9b-154">Em Visual Studio Code, abra **bundle.js** dessa pasta.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-154">In Visual Studio Code, open **bundle.js** from that folder.</span></span> <span data-ttu-id="f1c9b-155">O padrão do caminho do arquivo deve ser o seguinte:</span><span class="sxs-lookup"><span data-stu-id="f1c9b-155">The pattern of the file path should be as follows:</span></span>

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. <span data-ttu-id="f1c9b-156">Coloque pontos de interrupção em bundle.js onde você quer que o depurador pare.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-156">Place breakpoints in bundle.js where you want the debugger to stop.</span></span>
1. <span data-ttu-id="f1c9b-157">Na **lista suspensa do DEBUG,** selecione o nome **Depuração Direta**, e selecione **Executar**.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-157">In the **DEBUG** dropdown, select the name **Direct Debugging**, then select **Run**.</span></span>

    ![Captura de tela de seleção de depuração direta das opções de configuração no Visual Studio Code Debug Dropdown](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a><span data-ttu-id="f1c9b-159">depurar</span><span class="sxs-lookup"><span data-stu-id="f1c9b-159">Debug</span></span>

1. <span data-ttu-id="f1c9b-160">Depois de confirmar que o depurador está conectado, retorne ao Outlook e na caixa de diálogo manipulador baseado em **Evento Debug,** escolha **OK** .</span><span class="sxs-lookup"><span data-stu-id="f1c9b-160">After confirming that the debugger is attached, return to Outlook, and in the **Debug Event-based handler** dialog, choose **OK** .</span></span>

1. <span data-ttu-id="f1c9b-161">Agora você pode acertar seus pontos de interrupção em Visual Studio Code, permitindo que você depure seu código de ativação baseado em eventos.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-161">You can now hit your breakpoints in Visual Studio Code, enabling you to debug your event-based activation code.</span></span>

## <a name="stop-debugging"></a><span data-ttu-id="f1c9b-162">Pare de depurar</span><span class="sxs-lookup"><span data-stu-id="f1c9b-162">Stop debugging</span></span>

<span data-ttu-id="f1c9b-163">Para parar de depurar o resto da sessão de desktop Outlook atual, na caixa de diálogo do manipulador baseado em **eventos Debug,** escolha **Cancelar**.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-163">To stop debugging for the rest of the current Outlook desktop session, in the **Debug Event-based handler** dialog, choose **Cancel**.</span></span> <span data-ttu-id="f1c9b-164">Para reo enable depurar, reinicie Outlook desktop.</span><span class="sxs-lookup"><span data-stu-id="f1c9b-164">To re-enable debugging, restart Outlook desktop.</span></span>

<span data-ttu-id="f1c9b-165">Para evitar que a caixa de diálogo do **manipulador baseado em Eventos de depuração** apareça e pare de depurar sessões de Outlook subsequentes, exclua a tecla de registro associada ou defina seu valor `0` para: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .</span><span class="sxs-lookup"><span data-stu-id="f1c9b-165">To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent Outlook sessions, delete the associated registry key or set its value to `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span>

## <a name="see-also"></a><span data-ttu-id="f1c9b-166">Confira também</span><span class="sxs-lookup"><span data-stu-id="f1c9b-166">See also</span></span>

- [<span data-ttu-id="f1c9b-167">Configure seu Outlook complemento para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="f1c9b-167">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
- [<span data-ttu-id="f1c9b-168">Depurar seu suplemento com o log do tempo de execução</span><span class="sxs-lookup"><span data-stu-id="f1c9b-168">Debug your add-in with runtime logging</span></span>](../testing/runtime-logging.md#runtime-logging-on-windows)