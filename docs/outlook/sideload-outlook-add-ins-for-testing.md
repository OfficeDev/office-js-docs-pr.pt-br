---
title: Realizar sideload de suplementos do Outlook para teste
description: Use o sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 47eb5da19f858b6e30339acc59da24a818fc0959
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077026"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="312d2-103">Realizar sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="312d2-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="312d2-104">Você pode usar sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="312d2-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-automatically"></a><span data-ttu-id="312d2-105">Sideload automaticamente</span><span class="sxs-lookup"><span data-stu-id="312d2-105">Sideload automatically</span></span>

<span data-ttu-id="312d2-106">Se você criou seu Outlook de usuário usando o gerador Yeoman para Office de [complementos,](https://github.com/OfficeDev/generator-office)o sideload será melhor feito através da linha de comando.</span><span class="sxs-lookup"><span data-stu-id="312d2-106">If you created your Outlook add-in using [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), sideloading is best done through the command line.</span></span> <span data-ttu-id="312d2-107">Isso aproveitará nossas ferramentas e sideload em todos os dispositivos com suporte em um comando.</span><span class="sxs-lookup"><span data-stu-id="312d2-107">This will take advantage of our tooling and sideload across all of your supported devices in one command.</span></span>

1. <span data-ttu-id="312d2-108">Usando a linha de comando, navegue até o diretório raiz do seu projeto de complemento gerado pelo Yeoman.</span><span class="sxs-lookup"><span data-stu-id="312d2-108">Using the command line, navigate to the root directory of your Yeoman generated add-in project.</span></span> <span data-ttu-id="312d2-109">Execute o comando `npm start`.</span><span class="sxs-lookup"><span data-stu-id="312d2-109">Run the command `npm start`.</span></span>

1. <span data-ttu-id="312d2-110">Seu Outlook de usuário será automaticamente sideload para Outlook no computador da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="312d2-110">Your Outlook add-in will automatically sideload to Outlook on your desktop computer.</span></span> <span data-ttu-id="312d2-111">Você verá uma caixa de diálogo aparecer, informando que há uma tentativa de sideload do add-in, listando o nome e o local do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="312d2-111">You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file.</span></span> <span data-ttu-id="312d2-112">Selecione **OK**, que registrará o manifesto.</span><span class="sxs-lookup"><span data-stu-id="312d2-112">Select **OK**, which will register the manifest.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="312d2-113">Se o manifesto contiver um erro ou o caminho para o manifesto for inválido, você receberá uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="312d2-113">If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.</span></span>

1. <span data-ttu-id="312d2-114">Se o manifesto não contiver erros e o caminho for válido, o seu complemento agora será sideload e estará disponível na área de trabalho e no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="312d2-114">If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web.</span></span> <span data-ttu-id="312d2-115">Ele também será instalado em todos os dispositivos com suporte.</span><span class="sxs-lookup"><span data-stu-id="312d2-115">It will also be installed across all your supported devices.</span></span>

## <a name="sideload-manually"></a><span data-ttu-id="312d2-116">Sideload manualmente</span><span class="sxs-lookup"><span data-stu-id="312d2-116">Sideload manually</span></span>

<span data-ttu-id="312d2-117">Embora seja recomendável fazer sideload automaticamente pela linha de comando, conforme abordado na seção anterior, você também pode fazer sideload manualmente de um Outlook de entrada com base no cliente Outlook.</span><span class="sxs-lookup"><span data-stu-id="312d2-117">Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in based on the Outlook client.</span></span>

### <a name="outlook-on-the-web"></a><span data-ttu-id="312d2-118">Outlook na Web</span><span class="sxs-lookup"><span data-stu-id="312d2-118">Outlook on the web</span></span>

<span data-ttu-id="312d2-119">O processo de sideload de um complemento no Outlook na Web depende se você está usando a versão nova ou clássica.</span><span class="sxs-lookup"><span data-stu-id="312d2-119">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="312d2-120">Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no novo Outlook na Web](#new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="312d2-120">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#new-outlook-on-the-web).</span></span>

    ![Captura de tela parcial da nova Outlook na Web de ferramentas.](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="312d2-122">Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no Outlook na Web clássico](#classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="312d2-122">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#classic-outlook-on-the-web).</span></span>

    ![Captura de tela parcial da barra de ferramentas Outlook na Web clássica.](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="312d2-124">Se sua organização tiver incluído seu logotipo na barra de ferramentas da caixa de correio, você verá algo um pouco diferente do mostrado nas imagens anteriores.</span><span class="sxs-lookup"><span data-stu-id="312d2-124">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="new-outlook-on-the-web"></a><span data-ttu-id="312d2-125">Novo Outlook na Web</span><span class="sxs-lookup"><span data-stu-id="312d2-125">New Outlook on the web</span></span>

1. <span data-ttu-id="312d2-126">Acesse o [Outlook na Web](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="312d2-126">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="312d2-127">Crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="312d2-127">Create a new message.</span></span>

1. <span data-ttu-id="312d2-128">Escolha **...** na parte inferior da nova mensagem e selecione **Obter Suplementos** menu que aparecer.</span><span class="sxs-lookup"><span data-stu-id="312d2-128">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Janela de composição de mensagem na nova Outlook na Web com a opção Obter Complementos realçada.](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="312d2-130">Na caixa de diálogo **Suplementos do Outlook**, selecione **Meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="312d2-130">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Os complementos para Outlook caixa de diálogo no novo Outlook na Web com Meus complementos selecionados.](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="312d2-132">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="312d2-132">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="312d2-133">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="312d2-133">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Gerenciar captura de tela de complementos apontando para Adicionar a partir de uma opção de arquivo.](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="312d2-p106">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="312d2-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="classic-outlook-on-the-web"></a><span data-ttu-id="312d2-137">Clássico Outlook na Web</span><span class="sxs-lookup"><span data-stu-id="312d2-137">Classic Outlook on the web</span></span>

1. <span data-ttu-id="312d2-138">Acesse o [Outlook na Web](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="312d2-138">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="312d2-139">Escolha o ícone de engrenagem na seção superior direita da barra de ferramentas e selecione **Gerenciar suplementos**.</span><span class="sxs-lookup"><span data-stu-id="312d2-139">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Outlook na Web captura de tela apontando para a opção Gerenciar os complementos.](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="312d2-141">Na página **Gerenciar suplementos**, selecione **Suplementos** e **Meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="312d2-141">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook na Web de armazenamento com Meus complementos selecionados.](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="312d2-143">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="312d2-143">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="312d2-144">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="312d2-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Gerenciar captura de tela de complementos apontando para Adicionar a partir de uma opção de arquivo.](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="312d2-p108">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="312d2-p108">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-on-the-desktop"></a><span data-ttu-id="312d2-148">Outlook na área de trabalho</span><span class="sxs-lookup"><span data-stu-id="312d2-148">Outlook on the desktop</span></span>

#### <a name="outlook-2016-or-later"></a><span data-ttu-id="312d2-149">Outlook 2016 ou posterior</span><span class="sxs-lookup"><span data-stu-id="312d2-149">Outlook 2016 or later</span></span>

1. <span data-ttu-id="312d2-150">Abra Outlook 2016 ou posterior no Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="312d2-150">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="312d2-151">Selecione o botão **Obter Suplementos** na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="312d2-151">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Outlook 2016 faixa de opções apontando para o botão Obter Complementos.](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="312d2-153">Se você não vir o botão **Obter Complementos** na sua versão do Outlook, selecione:</span><span class="sxs-lookup"><span data-stu-id="312d2-153">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="312d2-154">**Botão Armazenar** na faixa de opções, se disponível.</span><span class="sxs-lookup"><span data-stu-id="312d2-154">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="312d2-155">OU</span><span class="sxs-lookup"><span data-stu-id="312d2-155">OR</span></span>
    >
    > - <span data-ttu-id="312d2-156">**Menu** Arquivo e, em seguida, selecione  o botão **Gerenciar Complementos** na guia Informações para abrir a caixa de diálogo **Add-ins** no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="312d2-156">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="312d2-157">Você pode ver mais sobre a experiência da Web na seção anterior [Sideload an add-in in Outlook na Web](#outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="312d2-157">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).</span></span>

1. <span data-ttu-id="312d2-158">Se houver guias próximas à parte superior da caixa de diálogo, verifique se a guia **Complementos** está selecionada.</span><span class="sxs-lookup"><span data-stu-id="312d2-158">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="312d2-159">Escolha **Meus complementos**.</span><span class="sxs-lookup"><span data-stu-id="312d2-159">Choose **My add-ins**.</span></span>

    ![Outlook 2016 de armazenamento com Meus complementos selecionados.](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="312d2-161">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="312d2-161">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="312d2-162">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="312d2-162">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela da Loja apontando para Adicionar de uma opção de arquivo.](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="312d2-p111">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="312d2-p111">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

#### <a name="outlook-2013"></a><span data-ttu-id="312d2-166">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="312d2-166">Outlook 2013</span></span>

1. <span data-ttu-id="312d2-167">Abra Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="312d2-167">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="312d2-168">Selecione o menu **Arquivo** e selecione o botão  **Gerenciar Complementos** na guia Informações. Outlook abrirá a versão da Web em um navegador.</span><span class="sxs-lookup"><span data-stu-id="312d2-168">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="312d2-169">Siga as etapas na [seção Sideload de](#outlook-on-the-web) um Outlook na Web de acordo com sua versão do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="312d2-169">Follow the steps in the [Sideload an add-in in Outlook on the web](#outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="312d2-170">Remover um complemento com sideload</span><span class="sxs-lookup"><span data-stu-id="312d2-170">Remove a sideloaded add-in</span></span>

<span data-ttu-id="312d2-171">Em todas as versões do Outlook, a chave para remover um complemento sideload é a caixa de diálogo Meus **Complementos** que lista seus complementos instalados. Escolha a reellipse ( `...` ) para o complemento e selecione **Remover**.</span><span class="sxs-lookup"><span data-stu-id="312d2-171">On all versions of Outlook, the key to removing a sideloaded add-in is the **My Add-ins** dialog which lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then select **Remove**.</span></span>

<span data-ttu-id="312d2-172">Para navegar até a caixa de diálogo Meus **Complementos** para seu cliente Outlook, use as últimas etapas listadas para [sideload manual](#sideload-manually) nas seções anteriores deste artigo.</span><span class="sxs-lookup"><span data-stu-id="312d2-172">To navigate to the **My Add-ins** dialog box for your Outlook client, use the last steps listed for [manual sideloading](#sideload-manually) in the previous sections of this article.</span></span>

<span data-ttu-id="312d2-173">Para remover um complemento sideload do Outlook, use as etapas descritas anteriormente neste artigo para encontrar o add-in na seção **Complementos personalizados** da caixa de diálogo que lista seus complementos instalados. Escolha a reellipse ( ) para o complemento e `...` escolha **Remover** para remover esse complemento específico.</span><span class="sxs-lookup"><span data-stu-id="312d2-173">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then choose **Remove** to remove that specific add-in.</span></span> <span data-ttu-id="312d2-174">Feche a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="312d2-174">Close the dialog.</span></span>
