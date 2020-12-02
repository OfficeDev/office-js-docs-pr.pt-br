---
title: Realizar sideload de suplementos do Outlook para teste
description: Use o sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.
ms.date: 12/01/2020
localization_priority: Normal
ms.openlocfilehash: dea2125ccd64eba2e3f1695c8ca1111a710321a4
ms.sourcegitcommit: c2fd7f982f3da748ef6be5c3a7434d859f8b46b9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/02/2020
ms.locfileid: "49530924"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="da78e-103">Realizar sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="da78e-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="da78e-104">Você pode usar sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="da78e-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a><span data-ttu-id="da78e-105">Realizar o sideload de um suplemento do Outlook na Web</span><span class="sxs-lookup"><span data-stu-id="da78e-105">Sideload an add-in in Outlook on the web</span></span>

<span data-ttu-id="da78e-106">O processo de Sideload de um suplemento no Outlook na Web depende se você está usando a versão nova ou clássica.</span><span class="sxs-lookup"><span data-stu-id="da78e-106">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="da78e-107">Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no novo Outlook na Web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="da78e-107">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span></span>

    ![captura de tela parcial da barra de ferramentas do novo Outlook na Web](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="da78e-109">Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no Outlook na Web clássico](#sideload-an-add-in-in-classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="da78e-109">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span></span>

    ![captura de tela parcial da barra de ferramentas do Outlook na Web clássico](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="da78e-111">Se sua organização tiver incluído seu logotipo na barra de ferramentas da caixa de correio, você verá algo um pouco diferente do mostrado nas imagens anteriores.</span><span class="sxs-lookup"><span data-stu-id="da78e-111">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a><span data-ttu-id="da78e-112">Realizar sideload de um suplemento no novo Outlook na Web</span><span class="sxs-lookup"><span data-stu-id="da78e-112">Sideload an add-in in the new Outlook on the web</span></span>

1. <span data-ttu-id="da78e-113">Acesse o [Outlook no Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="da78e-113">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="da78e-114">No Outlook na Web, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="da78e-114">In Outlook on the web, create a new message.</span></span>

1. <span data-ttu-id="da78e-115">Escolha **...** na parte inferior da nova mensagem e selecione **Obter Suplementos** menu que aparecer.</span><span class="sxs-lookup"><span data-stu-id="da78e-115">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Janela para redigir a mensagem no novo Outlook na Web com a opção Obter Suplementos realçada](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="da78e-117">Na caixa de diálogo **Suplementos do Outlook**, selecione **Meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="da78e-117">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Suplementos para a caixa de diálogo do Outlook no novo Outlook na Web com Meus suplementos selecionado](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="da78e-119">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="da78e-119">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="da78e-120">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="da78e-120">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="da78e-p102">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="da78e-p102">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a><span data-ttu-id="da78e-124">Realizar sideload de um suplemento no Outlook na Web clássico</span><span class="sxs-lookup"><span data-stu-id="da78e-124">Sideload an add-in in classic Outlook on the web</span></span>

1. <span data-ttu-id="da78e-125">Acesse o [Outlook no Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="da78e-125">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="da78e-126">Escolha o ícone de engrenagem na seção superior direita da barra de ferramentas e selecione **Gerenciar suplementos**.</span><span class="sxs-lookup"><span data-stu-id="da78e-126">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Captura de tela do Outlook na Web apontando para a opção Gerenciar suplementos](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="da78e-128">Na página **Gerenciar suplementos**, selecione **Suplementos** e **Meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="da78e-128">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Caixa de diálogo da Loja do Outlook na Web com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="da78e-130">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="da78e-130">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="da78e-131">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="da78e-131">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="da78e-p104">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="da78e-p104">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a><span data-ttu-id="da78e-135">Realizar sideload de um suplemento do Outlook na área de trabalho</span><span class="sxs-lookup"><span data-stu-id="da78e-135">Sideload an add-in in Outlook on the desktop</span></span>

### <a name="outlook-2016-or-later"></a><span data-ttu-id="da78e-136">Outlook 2016 ou posterior</span><span class="sxs-lookup"><span data-stu-id="da78e-136">Outlook 2016 or later</span></span>

1. <span data-ttu-id="da78e-137">Abra o Outlook 2016 ou posterior no Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="da78e-137">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="da78e-138">Selecione o botão **Obter Suplementos** na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="da78e-138">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Faixa de opções do Outlook 2016 apontando para obter o botão de suplementos](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="da78e-140">Se você não vir o botão **obter suplementos** em sua versão do Outlook, selecione:</span><span class="sxs-lookup"><span data-stu-id="da78e-140">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="da78e-141">O botão **armazenar** na faixa de opções, se disponível.</span><span class="sxs-lookup"><span data-stu-id="da78e-141">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="da78e-142">OU</span><span class="sxs-lookup"><span data-stu-id="da78e-142">OR</span></span>
    >
    > - <span data-ttu-id="da78e-143">Menu **arquivo** e, em seguida, selecione o botão **gerenciar suplementos** na guia **informações** para abrir a caixa de diálogo **suplementos** no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="da78e-143">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="da78e-144">Você pode ver mais sobre a experiência da Web na seção anterior [Sideload um suplemento no Outlook na Web](#sideload-an-add-in-in-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="da78e-144">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web).</span></span>

1. <span data-ttu-id="da78e-145">Se houver guias próximas à parte superior da caixa de diálogo, verifique se a guia **suplementos** está selecionada.</span><span class="sxs-lookup"><span data-stu-id="da78e-145">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="da78e-146">Escolha **meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="da78e-146">Choose **My add-ins**.</span></span>

    ![Caixa de diálogo da Loja do Outlook 2016 com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="da78e-148">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="da78e-148">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="da78e-149">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="da78e-149">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela da Loja apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="da78e-p107">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="da78e-p107">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-2013"></a><span data-ttu-id="da78e-153">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="da78e-153">Outlook 2013</span></span>

1. <span data-ttu-id="da78e-154">Abra o Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="da78e-154">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="da78e-155">Selecione o menu **arquivo** e, em seguida, selecione o botão **gerenciar suplementos** na guia **informações** . O Outlook abrirá a versão da Web em um navegador.</span><span class="sxs-lookup"><span data-stu-id="da78e-155">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="da78e-156">Siga as etapas na seção [Sideload um suplemento no Outlook na Web](#sideload-an-add-in-in-outlook-on-the-web) de acordo com a sua versão do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="da78e-156">Follow the steps in the [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="da78e-157">Remover um suplemento do suplementos foi feito</span><span class="sxs-lookup"><span data-stu-id="da78e-157">Remove a sideloaded add-in</span></span>

<span data-ttu-id="da78e-158">Para remover um suplemento do suplementos foi feito do Outlook, use as etapas descritas anteriormente neste artigo para localizar o suplemento na seção **suplementos personalizados** da caixa de diálogo que lista os suplementos instalados do. Escolha as reticências ( `...` ) para o suplemento e escolha **remover** para remover esse suplemento específico.</span><span class="sxs-lookup"><span data-stu-id="da78e-158">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>