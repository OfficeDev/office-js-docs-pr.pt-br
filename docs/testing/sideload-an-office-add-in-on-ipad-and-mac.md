---
title: Realizar o sideload de suplementos do Office em um iPad ou Mac para teste
description: Testar o suplemento do Office no iPad e Mac por Sideload
ms.date: 02/18/2020
localization_priority: Normal
ms.openlocfilehash: 4863a55d21ab37411e76810a744f103cc364f7c1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719774"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="464c5-103">Realizar o sideload de suplementos do Office em um iPad ou Mac para teste</span><span class="sxs-lookup"><span data-stu-id="464c5-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="464c5-p101">Para ver como seu suplemento será executado no Office no iOS, você pode realizar o sideload do manifesto do seu suplemento em um iPad usando o iTunes ou realizar o sideload do manifesto do suplemento diretamente no Office no Mac. Esta ação não permite definir pontos de interrupção e depurar o código do seu suplemento enquanto ele estiver em execução, mas é possível ver como ele se comporta e verificar se a interface do usuário é utilizável e está sendo processada adequadamente.</span><span class="sxs-lookup"><span data-stu-id="464c5-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="464c5-106">Pré-requisitos do Office no iOS</span><span class="sxs-lookup"><span data-stu-id="464c5-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="464c5-107">Um computador com Windows ou Mac com [iTunes](https://www.apple.com/itunes/download/) instalado.</span><span class="sxs-lookup"><span data-stu-id="464c5-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>

- <span data-ttu-id="464c5-108">Um iPad executando o iOS 8.2 ou posterior com [Excel no iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) instalado e um cabo de sincronização.</span><span class="sxs-lookup"><span data-stu-id="464c5-108">An iPad running iOS 8.2 or later with [Excel on iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>

- <span data-ttu-id="464c5-109">O arquivo de manifesto .xml para o suplemento que você deseja testar.</span><span class="sxs-lookup"><span data-stu-id="464c5-109">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="464c5-110">Pré-requisitos do Office no Mac</span><span class="sxs-lookup"><span data-stu-id="464c5-110">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="464c5-111">Um Mac executando OS X v10.10 “Yosemite” ou posterior com [Office no Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) instalado.</span><span class="sxs-lookup"><span data-stu-id="464c5-111">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="464c5-112">Word no Mac versão 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="464c5-112">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="464c5-113">Excel no Mac versão 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="464c5-113">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="464c5-114">PowerPoint no Mac versão 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="464c5-114">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="464c5-115">O arquivo de manifesto .xml para o suplemento que você deseja testar.</span><span class="sxs-lookup"><span data-stu-id="464c5-115">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad"></a><span data-ttu-id="464c5-116">Realizar um sideload de um suplemento no Excel ou no Word no iPad</span><span class="sxs-lookup"><span data-stu-id="464c5-116">Sideload an add-in on Excel or Word on iPad</span></span>

1. <span data-ttu-id="464c5-117">Use um cabo de sincronização para conectar seu iPad ao computador.</span><span class="sxs-lookup"><span data-stu-id="464c5-117">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="464c5-118">Se você estiver conectando o iPad ao computador pela primeira vez, você será solicitado a **confiar neste computador?**.</span><span class="sxs-lookup"><span data-stu-id="464c5-118">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="464c5-119">Escolha **Confiar** para continuar.</span><span class="sxs-lookup"><span data-stu-id="464c5-119">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="464c5-120">No iTunes, escolha o ícone do **iPad** abaixo da barra de menus.</span><span class="sxs-lookup"><span data-stu-id="464c5-120">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="464c5-121">Em **configurações** no lado esquerdo do iTunes, escolha **aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="464c5-121">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="464c5-122">No lado direito do iTunes, role para baixo até **compartilhamento de arquivos**e, em seguida, escolha **Excel** ou **Word** na coluna **suplementos** .</span><span class="sxs-lookup"><span data-stu-id="464c5-122">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="464c5-123">Na parte inferior da coluna documentos do **Excel** ou do **Word** , escolha **Adicionar arquivo**e, em seguida, selecione o arquivo manifest. XML do suplemento que você deseja Sideload.</span><span class="sxs-lookup"><span data-stu-id="464c5-123">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="464c5-124">Abra o aplicativo Excel ou Word em seu iPad.</span><span class="sxs-lookup"><span data-stu-id="464c5-124">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="464c5-125">Se o aplicativo Excel ou Word já estiver em execução, escolha o botão **página inicial** e, em seguida, feche e reinicie o aplicativo.</span><span class="sxs-lookup"><span data-stu-id="464c5-125">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="464c5-126">Abra um documento.</span><span class="sxs-lookup"><span data-stu-id="464c5-126">Open a document.</span></span>

8. <span data-ttu-id="464c5-127">Escolha **suplementos** na guia **Inserir** . O suplemento do suplementos foi feito está disponível para inserção sob o título do **desenvolvedor** na interface do usuário de **suplementos** .</span><span class="sxs-lookup"><span data-stu-id="464c5-127">Choose **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Inserir Suplementos no aplicativo do Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="464c5-129">Realizar sideload de um suplemento no Office no Mac</span><span class="sxs-lookup"><span data-stu-id="464c5-129">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="464c5-130">Para realizar o sideload de um suplemento do Outlook no Mac, confira [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="464c5-130">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="464c5-131">Abra o **terminal** e vá para uma das seguintes pastas onde você salvará o arquivo de manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="464c5-131">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="464c5-132">Se a pasta `wef` não existir em seu computador, crie-a.</span><span class="sxs-lookup"><span data-stu-id="464c5-132">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="464c5-133">Para o Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="464c5-133">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>    
    - <span data-ttu-id="464c5-134">Para o Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="464c5-134">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="464c5-135">Para o PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="464c5-135">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="464c5-136">Abra a pasta no **Finder** usando o comando `open .` (incluindo o ponto ou ponto).</span><span class="sxs-lookup"><span data-stu-id="464c5-136">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="464c5-137">Copie o arquivo de manifesto do suplemento nessa pasta.</span><span class="sxs-lookup"><span data-stu-id="464c5-137">Copy your add-in's manifest file to this folder.</span></span>

    ![Pasta Wef no Office no Mac](../images/all-my-files.png)

3. <span data-ttu-id="464c5-p106">Abra o Word e abra um documento. Reinicie o Word se já estiver em execução.</span><span class="sxs-lookup"><span data-stu-id="464c5-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="464c5-141">No Word, escolha **Inserir** > **suplementos** > **meus** suplementos (menu suspenso) e, em seguida, escolha seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="464c5-141">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Meus Suplementos no Office no Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="464c5-p107">Aplicativos em que foi feito o sideload não aparecerão na caixa de diálogo Meus Suplementos. Eles só ficam visíveis dentro do menu suspenso (pequena seta para baixo à direita de Meus Suplementos na guia **Inserir**). Os suplementos em que foi feito o sideload são exibidos na lista sob o título **Suplementos do Desenvolvedor** nesse menu.</span><span class="sxs-lookup"><span data-stu-id="464c5-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="464c5-146">Verifique se o seu suplemento é exibido no Word.</span><span class="sxs-lookup"><span data-stu-id="464c5-146">Verify that your add-in is displayed in Word.</span></span>

    ![Suplemento do Office exibido no Office no Mac](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="464c5-148">Remover um suplemento do suplementos foi feito</span><span class="sxs-lookup"><span data-stu-id="464c5-148">Remove a sideloaded add-in</span></span>

<span data-ttu-id="464c5-149">Você pode remover um suplemento suplementos foi feito anteriormente limpando o cache do Office em seu computador.</span><span class="sxs-lookup"><span data-stu-id="464c5-149">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="464c5-150">Detalhes sobre como limpar o cache para cada plataforma e host podem ser encontrados no artigo [limpar o cache do Office](clear-cache.md).</span><span class="sxs-lookup"><span data-stu-id="464c5-150">Details on how to clear the cache for each platform and host can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="464c5-151">Confira também</span><span class="sxs-lookup"><span data-stu-id="464c5-151">See also</span></span>

- [<span data-ttu-id="464c5-152">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="464c5-152">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
