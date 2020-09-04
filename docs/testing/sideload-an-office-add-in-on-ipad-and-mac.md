---
title: Realizar o sideload de suplementos do Office em um iPad ou Mac para teste
description: Teste o suplemento do Office no iPad e no Mac por Sideload.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 7c5e9542c6e6f9abc96defde389b9543421b8529
ms.sourcegitcommit: 604361e55dee45c7a5d34c2fa6937693c154fc24
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/03/2020
ms.locfileid: "47364050"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="0ac84-103">Realizar o sideload de suplementos do Office em um iPad ou Mac para teste</span><span class="sxs-lookup"><span data-stu-id="0ac84-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="0ac84-p101">Para ver como seu suplemento será executado no Office no iOS, você pode realizar o sideload do manifesto do seu suplemento em um iPad usando o iTunes ou realizar o sideload do manifesto do suplemento diretamente no Office no Mac. Esta ação não permite definir pontos de interrupção e depurar o código do seu suplemento enquanto ele estiver em execução, mas é possível ver como ele se comporta e verificar se a interface do usuário é utilizável e está sendo processada adequadamente.</span><span class="sxs-lookup"><span data-stu-id="0ac84-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="0ac84-106">Pré-requisitos do Office no iOS</span><span class="sxs-lookup"><span data-stu-id="0ac84-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="0ac84-107">Um computador com Windows ou Mac com [iTunes](https://www.apple.com/itunes/download/) instalado.</span><span class="sxs-lookup"><span data-stu-id="0ac84-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
  > [!IMPORTANT]
  > <span data-ttu-id="0ac84-108">Se você estiver executando o macOS Catalina, o [iTunes não estará mais disponível](https://support.apple.com/HT210200) , portanto, você deve seguir as instruções na seção [Sideload um suplemento no Excel ou no Word no iPad usando o MacOS,](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="0ac84-108">If you're running macOS Catalina, [iTunes is no longer available](https://support.apple.com/HT210200) so you should follow the instructions in the section [Sideload an add-in on Excel or Word on iPad using macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) later in this article.</span></span>

- <span data-ttu-id="0ac84-109">Um iPad executando o iOS 8,2 ou posterior com o [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) ou [Word](https://apps.apple.com/app/microsoft-word/id586447913) instalado e um cabo de sincronização.</span><span class="sxs-lookup"><span data-stu-id="0ac84-109">An iPad running iOS 8.2 or later with [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) or [Word](https://apps.apple.com/app/microsoft-word/id586447913) installed, and a sync cable.</span></span>

- <span data-ttu-id="0ac84-110">O arquivo de manifesto .xml para o suplemento que você deseja testar.</span><span class="sxs-lookup"><span data-stu-id="0ac84-110">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="0ac84-111">Pré-requisitos do Office no Mac</span><span class="sxs-lookup"><span data-stu-id="0ac84-111">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="0ac84-112">Um Mac executando OS X v10.10 “Yosemite” ou posterior com [Office no Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) instalado.</span><span class="sxs-lookup"><span data-stu-id="0ac84-112">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="0ac84-113">Word no Mac versão 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="0ac84-113">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="0ac84-114">Excel no Mac versão 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="0ac84-114">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="0ac84-115">PowerPoint no Mac versão 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="0ac84-115">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="0ac84-116">O arquivo de manifesto .xml para o suplemento que você deseja testar.</span><span class="sxs-lookup"><span data-stu-id="0ac84-116">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a><span data-ttu-id="0ac84-117">Sideload um suplemento no Excel ou no Word no iPad usando iTunes</span><span class="sxs-lookup"><span data-stu-id="0ac84-117">Sideload an add-in on Excel or Word on iPad using iTunes</span></span>

1. <span data-ttu-id="0ac84-118">Use um cabo de sincronização para conectar seu iPad ao computador.</span><span class="sxs-lookup"><span data-stu-id="0ac84-118">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="0ac84-119">Se você estiver conectando o iPad ao computador pela primeira vez, você será solicitado a **confiar neste computador?**.</span><span class="sxs-lookup"><span data-stu-id="0ac84-119">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="0ac84-120">Escolha **Confiar** para continuar.</span><span class="sxs-lookup"><span data-stu-id="0ac84-120">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="0ac84-121">No iTunes, escolha o ícone do **iPad** abaixo da barra de menus.</span><span class="sxs-lookup"><span data-stu-id="0ac84-121">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="0ac84-122">Em **configurações** no lado esquerdo do iTunes, escolha **aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="0ac84-122">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="0ac84-123">No lado direito do iTunes, role para baixo até **compartilhamento de arquivos**e, em seguida, escolha **Excel** ou **Word** na coluna **suplementos** .</span><span class="sxs-lookup"><span data-stu-id="0ac84-123">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="0ac84-124">Na parte inferior da coluna documentos do **Excel** ou do **Word** , escolha **Adicionar arquivo**e, em seguida, selecione o arquivo manifest. XML do suplemento que você deseja Sideload.</span><span class="sxs-lookup"><span data-stu-id="0ac84-124">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="0ac84-125">Abra o aplicativo Excel ou Word em seu iPad.</span><span class="sxs-lookup"><span data-stu-id="0ac84-125">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="0ac84-126">Se o aplicativo Excel ou Word já estiver em execução, escolha o botão **página inicial** e, em seguida, feche e reinicie o aplicativo.</span><span class="sxs-lookup"><span data-stu-id="0ac84-126">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="0ac84-127">Abra um documento.</span><span class="sxs-lookup"><span data-stu-id="0ac84-127">Open a document.</span></span>

8. <span data-ttu-id="0ac84-128">Escolha **suplementos** na guia **Inserir** . (na guia **Inserir** , talvez seja necessário rolar horizontalmente até que você veja o botão **suplementos** .) O suplemento do suplementos foi feito está disponível para inserção sob o título do **desenvolvedor** na interface do usuário de **suplementos** .</span><span class="sxs-lookup"><span data-stu-id="0ac84-128">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Inserir Suplementos no aplicativo do Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a><span data-ttu-id="0ac84-130">Sideload um suplemento no Excel ou no Word no iPad com o macOS Catalina</span><span class="sxs-lookup"><span data-stu-id="0ac84-130">Sideload an add-in on Excel or Word on iPad using macOS Catalina</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0ac84-131">Com a introdução do macOS Catalina, a [Apple descontinuava o iTunes no Mac](https://support.apple.com/HT210200) e a funcionalidade integrada necessária para os aplicativos do Sideload no **Finder**.</span><span class="sxs-lookup"><span data-stu-id="0ac84-131">With the introduction of macOS Catalina, [Apple discontinued iTunes on Mac](https://support.apple.com/HT210200) and integrated functionality required to sideload apps into **Finder**.</span></span>

1. <span data-ttu-id="0ac84-132">Use um cabo de sincronização para conectar seu iPad ao computador.</span><span class="sxs-lookup"><span data-stu-id="0ac84-132">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="0ac84-133">Se você estiver conectando o iPad ao computador pela primeira vez, você será solicitado a **confiar neste computador?**.</span><span class="sxs-lookup"><span data-stu-id="0ac84-133">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="0ac84-134">Escolha **Confiar** para continuar.</span><span class="sxs-lookup"><span data-stu-id="0ac84-134">Choose **Trust** to continue.</span></span> <span data-ttu-id="0ac84-135">Você também pode ser perguntado se este é um novo iPad ou se você está restaurando um.</span><span class="sxs-lookup"><span data-stu-id="0ac84-135">You may also be asked if this is a new iPad or if you're restoring one.</span></span>

2. <span data-ttu-id="0ac84-136">No Finder, em **locais**, escolha o ícone do **iPad** abaixo da barra de menus.</span><span class="sxs-lookup"><span data-stu-id="0ac84-136">In Finder, under **Locations**, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="0ac84-137">Na parte superior da janela Localizador, clique em **arquivos**e, em seguida, localize **Excel** ou **Word**.</span><span class="sxs-lookup"><span data-stu-id="0ac84-137">On the top of the Finder window, click on **Files**, and then locate **Excel** or **Word**.</span></span>

4. <span data-ttu-id="0ac84-138">Em uma janela de localizador diferente, arraste e solte o manifest.xml arquivo do suplemento que você deseja carregar no lado do arquivo do **Excel** ou **Word** na primeira janela do Finder.</span><span class="sxs-lookup"><span data-stu-id="0ac84-138">From a different Finder window, drag and drop the manifest.xml file of the add-in you want to side load onto the **Excel** or **Word** file in the first Finder window.</span></span>

5. <span data-ttu-id="0ac84-139">Abra o aplicativo Excel ou Word em seu iPad.</span><span class="sxs-lookup"><span data-stu-id="0ac84-139">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="0ac84-140">Se o aplicativo Excel ou Word já estiver em execução, escolha o botão **página inicial** e, em seguida, feche e reinicie o aplicativo.</span><span class="sxs-lookup"><span data-stu-id="0ac84-140">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

6. <span data-ttu-id="0ac84-141">Abra um documento.</span><span class="sxs-lookup"><span data-stu-id="0ac84-141">Open a document.</span></span>

7. <span data-ttu-id="0ac84-142">Escolha **suplementos** na guia **Inserir** . (na guia **Inserir** , talvez seja necessário rolar horizontalmente até que você veja o botão **suplementos** .) O suplemento do suplementos foi feito está disponível para inserção sob o título do **desenvolvedor** na interface do usuário de **suplementos** .</span><span class="sxs-lookup"><span data-stu-id="0ac84-142">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Inserir Suplementos no aplicativo do Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="0ac84-144">Realizar sideload de um suplemento no Office no Mac</span><span class="sxs-lookup"><span data-stu-id="0ac84-144">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="0ac84-145">Para realizar o sideload de um suplemento do Outlook no Mac, confira [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-an-add-in-in-outlook-on-the-desktop).</span><span class="sxs-lookup"><span data-stu-id="0ac84-145">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-an-add-in-in-outlook-on-the-desktop).</span></span>

1. <span data-ttu-id="0ac84-146">Abra o **terminal** e vá para uma das seguintes pastas onde você salvará o arquivo de manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="0ac84-146">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="0ac84-147">Se a pasta `wef` não existir em seu computador, crie-a.</span><span class="sxs-lookup"><span data-stu-id="0ac84-147">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="0ac84-148">Para o Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="0ac84-148">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>
    - <span data-ttu-id="0ac84-149">Para o Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="0ac84-149">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="0ac84-150">Para o PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="0ac84-150">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="0ac84-151">Abra a pasta no **Finder** usando o comando `open .` (incluindo o ponto ou ponto).</span><span class="sxs-lookup"><span data-stu-id="0ac84-151">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="0ac84-152">Copie o arquivo de manifesto do suplemento nessa pasta.</span><span class="sxs-lookup"><span data-stu-id="0ac84-152">Copy your add-in's manifest file to this folder.</span></span>

    ![Pasta Wef no Office no Mac](../images/all-my-files.png)

3. <span data-ttu-id="0ac84-p108">Abra o Word e abra um documento. Reinicie o Word se já estiver em execução.</span><span class="sxs-lookup"><span data-stu-id="0ac84-p108">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="0ac84-156">No Word, escolha **Inserir**  >  **suplementos**  >  **meus** suplementos (menu suspenso) e, em seguida, escolha seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="0ac84-156">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Meus Suplementos no Office no Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="0ac84-p109">Aplicativos em que foi feito o sideload não aparecerão na caixa de diálogo Meus Suplementos. Eles só ficam visíveis dentro do menu suspenso (pequena seta para baixo à direita de Meus Suplementos na guia **Inserir**). Os suplementos em que foi feito o sideload são exibidos na lista sob o título **Suplementos do Desenvolvedor** nesse menu.</span><span class="sxs-lookup"><span data-stu-id="0ac84-p109">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="0ac84-161">Verifique se o seu suplemento é exibido no Word.</span><span class="sxs-lookup"><span data-stu-id="0ac84-161">Verify that your add-in is displayed in Word.</span></span>

    ![Suplemento do Office exibido no Office no Mac](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="0ac84-163">Remover um suplemento do suplementos foi feito</span><span class="sxs-lookup"><span data-stu-id="0ac84-163">Remove a sideloaded add-in</span></span>

<span data-ttu-id="0ac84-164">Você pode remover um suplemento suplementos foi feito anteriormente limpando o cache do Office em seu computador.</span><span class="sxs-lookup"><span data-stu-id="0ac84-164">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="0ac84-165">Detalhes sobre como limpar o cache para cada plataforma e aplicativo podem ser encontrados no artigo [limpar o cache do Office](clear-cache.md).</span><span class="sxs-lookup"><span data-stu-id="0ac84-165">Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0ac84-166">Confira também</span><span class="sxs-lookup"><span data-stu-id="0ac84-166">See also</span></span>

- [<span data-ttu-id="0ac84-167">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="0ac84-167">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
