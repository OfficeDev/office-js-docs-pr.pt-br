---
title: Realizar sideload de suplementos do Outlook para teste
description: Use o sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.
ms.date: 02/10/2021
localization_priority: Normal
ms.openlocfilehash: b783b815af84a7fd8b4abd52cdd8e0925bfb9ecf
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234244"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="dd5bd-103">Realizar sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="dd5bd-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="dd5bd-104">Você pode usar sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-automatically"></a><span data-ttu-id="dd5bd-105">Sideload automaticamente</span><span class="sxs-lookup"><span data-stu-id="dd5bd-105">Sideload automatically</span></span>

<span data-ttu-id="dd5bd-106">Se você criou seu complemento do Outlook usando o gerador [Yeoman](https://github.com/OfficeDev/generator-office)para Os Complementos do Office, o sideload será feito da melhor maneira por meio da linha de comando.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-106">If you created your Outlook add-in using [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), sideloading is best done through the command line.</span></span> <span data-ttu-id="dd5bd-107">Isso aproveitará nossas ferramentas e sideload em todos os seus dispositivos com suporte em um único comando.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-107">This will take advantage of our tooling and sideload across all of your supported devices in one command.</span></span>

1. <span data-ttu-id="dd5bd-108">Usando a linha de comando, navegue até o diretório raiz do seu projeto de complemento gerado pelo Yeoman.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-108">Using the command line, navigate to the root directory of your Yeoman generated add-in project.</span></span> <span data-ttu-id="dd5bd-109">Execute o comando `npm start`.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-109">Run the command `npm start`.</span></span>

2. <span data-ttu-id="dd5bd-110">O seu complemento do Outlook fará o sideload automático para o Outlook em seu computador desktop.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-110">Your Outlook add-in will automatically sideload to Outlook on your desktop computer.</span></span> <span data-ttu-id="dd5bd-111">Você verá uma caixa de diálogo aparecer, informando que há uma tentativa de sideload do complemento, listando o nome e o local do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-111">You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file.</span></span> <span data-ttu-id="dd5bd-112">Selecione **OK**, que registrará o manifesto.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-112">Select **OK**, which will register the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dd5bd-113">Se o manifesto contiver um erro ou o caminho para o manifesto for inválido, você receberá uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-113">If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.</span></span>

3. <span data-ttu-id="dd5bd-114">Se o manifesto não contiver erros e o caminho for válido, seu complemento agora será sideloaded e estará disponível na área de trabalho e no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-114">If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web.</span></span> <span data-ttu-id="dd5bd-115">Ele também será instalado em todos os seus dispositivos compatíveis.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-115">It will also be installed across all your supported devices.</span></span>

## <a name="sideload-manually"></a><span data-ttu-id="dd5bd-116">Realizar sideload manualmente</span><span class="sxs-lookup"><span data-stu-id="dd5bd-116">Sideload manually</span></span>

<span data-ttu-id="dd5bd-117">Embora seja recomendável realizar o sideload automaticamente pela linha de comando, conforme abordado na seção anterior, você também pode fazer o sideload manual de um complemento do Outlook com base no cliente do Outlook.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-117">Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in based on the Outlook client.</span></span>

### <a name="outlook-on-the-web"></a><span data-ttu-id="dd5bd-118">Outlook na Web</span><span class="sxs-lookup"><span data-stu-id="dd5bd-118">Outlook on the web</span></span>

<span data-ttu-id="dd5bd-119">O processo de sideload de um complemento no Outlook na Web depende se você está usando a versão nova ou clássica.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-119">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="dd5bd-120">Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no novo Outlook na Web](#new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="dd5bd-120">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#new-outlook-on-the-web).</span></span>

    ![captura de tela parcial da barra de ferramentas do novo Outlook na Web](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="dd5bd-122">Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no Outlook na Web clássico](#classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="dd5bd-122">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#classic-outlook-on-the-web).</span></span>

    ![captura de tela parcial da barra de ferramentas do Outlook na Web clássico](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="dd5bd-124">Se sua organização tiver incluído seu logotipo na barra de ferramentas da caixa de correio, você verá algo um pouco diferente do mostrado nas imagens anteriores.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-124">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="new-outlook-on-the-web"></a><span data-ttu-id="dd5bd-125">Novo Outlook na Web</span><span class="sxs-lookup"><span data-stu-id="dd5bd-125">New Outlook on the web</span></span>

1. <span data-ttu-id="dd5bd-126">Acesse o [Outlook na Web](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="dd5bd-126">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="dd5bd-127">Crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-127">Create a new message.</span></span>

1. <span data-ttu-id="dd5bd-128">Escolha **...** na parte inferior da nova mensagem e selecione **Obter Suplementos** menu que aparecer.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-128">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Janela para redigir a mensagem no novo Outlook na Web com a opção Obter Suplementos realçada](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="dd5bd-130">Na caixa de diálogo **Suplementos do Outlook**, selecione **Meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-130">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Suplementos para a caixa de diálogo do Outlook no novo Outlook na Web com Meus suplementos selecionado](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="dd5bd-132">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-132">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="dd5bd-133">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-133">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="dd5bd-p106">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="classic-outlook-on-the-web"></a><span data-ttu-id="dd5bd-137">Outlook na Web clássico</span><span class="sxs-lookup"><span data-stu-id="dd5bd-137">Classic Outlook on the web</span></span>

1. <span data-ttu-id="dd5bd-138">Acesse o [Outlook na Web](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="dd5bd-138">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="dd5bd-139">Escolha o ícone de engrenagem na seção superior direita da barra de ferramentas e selecione **Gerenciar suplementos**.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-139">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Captura de tela do Outlook na Web apontando para a opção Gerenciar suplementos](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="dd5bd-141">Na página **Gerenciar suplementos**, selecione **Suplementos** e **Meus suplementos**.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-141">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Caixa de diálogo da Loja do Outlook na Web com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="dd5bd-143">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-143">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="dd5bd-144">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="dd5bd-p108">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-p108">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-on-the-desktop"></a><span data-ttu-id="dd5bd-148">Outlook na área de trabalho</span><span class="sxs-lookup"><span data-stu-id="dd5bd-148">Outlook on the desktop</span></span>

#### <a name="outlook-2016-or-later"></a><span data-ttu-id="dd5bd-149">Outlook 2016 ou posterior</span><span class="sxs-lookup"><span data-stu-id="dd5bd-149">Outlook 2016 or later</span></span>

1. <span data-ttu-id="dd5bd-150">Abra o Outlook 2016 ou posterior no Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-150">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="dd5bd-151">Selecione o botão **Obter Suplementos** na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-151">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Faixa de opções do Outlook 2016 apontando para o botão Obter Complementos](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="dd5bd-153">Se você não vir o botão **Obter Complementos** em sua versão do Outlook, selecione:</span><span class="sxs-lookup"><span data-stu-id="dd5bd-153">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="dd5bd-154">**Botão Armazenar** na faixa de opções, se disponível.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-154">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="dd5bd-155">OU</span><span class="sxs-lookup"><span data-stu-id="dd5bd-155">OR</span></span>
    >
    > - <span data-ttu-id="dd5bd-156">**Menu** Arquivo e, em **seguida, selecione**  o botão Gerenciar Complementos na guia Informações para abrir a caixa de diálogo **Add-ins** no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-156">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="dd5bd-157">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="dd5bd-157">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).</span></span>

1. <span data-ttu-id="dd5bd-158">Se houver guias próximas à parte superior da caixa de diálogo, verifique se a guia **Add-ins** está selecionada.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-158">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="dd5bd-159">Escolha **Meus complementos.**</span><span class="sxs-lookup"><span data-stu-id="dd5bd-159">Choose **My add-ins**.</span></span>

    ![Caixa de diálogo da Loja do Outlook 2016 com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="dd5bd-161">Localize a seção **Suplementos personalizados** no final da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-161">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="dd5bd-162">Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-162">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Captura de tela da Loja apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="dd5bd-p111">Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-p111">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

#### <a name="outlook-2013"></a><span data-ttu-id="dd5bd-166">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="dd5bd-166">Outlook 2013</span></span>

1. <span data-ttu-id="dd5bd-167">Abra o Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-167">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="dd5bd-168">Selecione **o** menu Arquivo e selecione **o botão Gerenciar Complementos** na **guia** Informações. O Outlook abrirá a versão da Web em um navegador.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-168">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="dd5bd-169">Siga as etapas na [seção Sideload](#outlook-on-the-web) de um complemento no Outlook na Web de acordo com sua versão do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-169">Follow the steps in the [Sideload an add-in in Outlook on the web](#outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="dd5bd-170">Remover um complemento de sideload</span><span class="sxs-lookup"><span data-stu-id="dd5bd-170">Remove a sideloaded add-in</span></span>

<span data-ttu-id="dd5bd-171">Em todas as versões do Outlook, a chave para remover um complemento de sideload é a caixa de diálogo Meus **Complementos,** que lista seus complementos instalados. Escolha as reellipses ( `...` ) para o complemento e selecione **Remover**.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-171">On all versions of Outlook, the key to removing a sideloaded add-in is the **My Add-ins** dialog which lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then select **Remove**.</span></span>

<span data-ttu-id="dd5bd-172">Para navegar até a caixa de diálogo Meus **Complementos** para seu cliente do Outlook, use as últimas etapas listadas para [sideload manual](#sideload-manually) nas seções anteriores deste artigo.</span><span class="sxs-lookup"><span data-stu-id="dd5bd-172">To navigate to the **My Add-ins** dialog box for your Outlook client, use the last steps listed for [manual sideloading](#sideload-manually) in the previous sections of this article.</span></span>

<span data-ttu-id="dd5bd-173">Para remover um complemento de sideload do Outlook, use as etapas descritas anteriormente neste artigo para encontrar o add-in na seção de **complementos personalizados** da caixa de diálogo que lista seus complementos instalados. Escolha as reellipses ( ) para o complemento e, em seguida, escolha Remover para `...` remover esse complemento específico. </span><span class="sxs-lookup"><span data-stu-id="dd5bd-173">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>

