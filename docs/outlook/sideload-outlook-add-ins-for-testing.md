---
title: Realizar sideload de suplementos do Outlook para teste
description: Use o sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.
ms.date: 06/24/2019
localization_priority: Normal
ms.openlocfilehash: 3543eeb58f441819edb2c129e6e14206e26de524
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605322"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="1ab8e-103">Realizar sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="1ab8e-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="1ab8e-104">Você pode usar sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>


## <a name="sideload-an-add-in-in-outlook-in-office-365"></a><span data-ttu-id="1ab8e-105">Realizar sideload de um suplemento do Outlook no Office 365</span><span class="sxs-lookup"><span data-stu-id="1ab8e-105">Sideload an add-in in Outlook in Office 365</span></span>

<span data-ttu-id="1ab8e-106">O processo de sideload de um suplemento do Outlook no Office 365 depende de se você está usando o novo Outlook na Web ou o Outlook na Web clássico.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-106">The process for sideloading an add-in in Outlook in Office 365 depends upon whether you are using the new Outlook on the web or classic Outlook on the web.</span></span>

- <span data-ttu-id="1ab8e-107">Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no novo Outlook na Web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="1ab8e-107">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span></span>

    ![captura de tela parcial da barra de ferramentas do novo Outlook na Web](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="1ab8e-109">Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no Outlook na Web clássico](#sideload-an-add-in-in-classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="1ab8e-109">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span></span>

    ![captura de tela parcial da barra de ferramentas do Outlook na Web clássico](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="1ab8e-111">Se sua organização tiver incluído seu logotipo na barra de ferramentas da caixa de correio, você verá algo um pouco diferente do mostrado nas imagens anteriores.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-111">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a><span data-ttu-id="1ab8e-112">Realizar sideload de um suplemento no novo Outlook na Web</span><span class="sxs-lookup"><span data-stu-id="1ab8e-112">Sideload an add-in in the new Outlook on the web</span></span>

1. <span data-ttu-id="1ab8e-113">Acesse o [Outlook no Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="1ab8e-113">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="1ab8e-114">No Outlook na Web, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-114">In Outlook on the web, create a new message.</span></span>   

1. <span data-ttu-id="1ab8e-115">Escolha **...** na parte inferior da nova mensagem e selecione **Obter Suplementos** menu que aparecer.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-115">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Janela para redigir a mensagem no novo Outlook na Web com a opção Obter Suplementos realçada](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="1ab8e-117">Na caixa de diálogo **Suplementos do Outlook**, selecione **Meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-117">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Suplementos para a caixa de diálogo do Outlook no novo Outlook na Web com Meus suplementos selecionado](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="1ab8e-119">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-119">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="1ab8e-120">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-120">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="1ab8e-p102">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-p102">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a><span data-ttu-id="1ab8e-124">Realizar sideload de um suplemento no Outlook na Web clássico</span><span class="sxs-lookup"><span data-stu-id="1ab8e-124">Sideload an add-in in classic Outlook on the web</span></span>

1. <span data-ttu-id="1ab8e-125">Acesse o [Outlook no Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="1ab8e-125">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="1ab8e-126">Escolha o ícone de engrenagem na seção superior direita da barra de ferramentas e selecione **Gerenciar suplementos**.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-126">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Captura de tela do Outlook na Web apontando para a opção Gerenciar suplementos](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="1ab8e-128">Na página **Gerenciar suplementos**, selecione **Suplementos** e **Meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-128">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Caixa de diálogo da Loja do Outlook na Web com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="1ab8e-130">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-130">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="1ab8e-131">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-131">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="1ab8e-p104">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-p104">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a><span data-ttu-id="1ab8e-135">Realizar sideload de um suplemento do Outlook na área de trabalho</span><span class="sxs-lookup"><span data-stu-id="1ab8e-135">Sideload an add-in in Outlook on the desktop</span></span>

1. <span data-ttu-id="1ab8e-136">Abra o Outlook 2013 ou posterior no Windows ou Outlook 2016 ou posterior no Mac.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-136">Open Outlook 2013 or later on Windows, or Outlook 2016 or later on Mac.</span></span>

1. <span data-ttu-id="1ab8e-137">Selecione o botão **Obter Suplementos** na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-137">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Faixa de opções do Outlook 2016 apontando para o botão Store](../images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > <span data-ttu-id="1ab8e-139">Caso não veja o botão **Obter Suplementos** em sua versão do Outlook, selecione o botão **Store** na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-139">If you don't see the **Get Add-ins** button in your version of Outlook, select the **Store** button on the ribbon instead.</span></span>

1. <span data-ttu-id="1ab8e-140">Selecione **Suplementos** e, depois, **Meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-140">Select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Caixa de diálogo da Loja do Outlook 2016 com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="1ab8e-142">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-142">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="1ab8e-143">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-143">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela da Loja apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="1ab8e-p106">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="1ab8e-147">Remover um suplemento do suplementos foi feito</span><span class="sxs-lookup"><span data-stu-id="1ab8e-147">Remove a sideloaded add-in</span></span>

<span data-ttu-id="1ab8e-148">Para remover um suplemento do suplementos foi feito do Outlook, use as etapas descritas anteriormente neste artigo para localizar o suplemento na seção **suplementos personalizados** da caixa de diálogo que lista seus suplementos instalados. escolha as reticências ( `...` ) para o suplemento e, em seguida, escolha **remover** para remover o suplemento específico do.</span><span class="sxs-lookup"><span data-stu-id="1ab8e-148">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>