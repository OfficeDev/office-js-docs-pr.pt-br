---
title: Realizar o sideload de suplementos do Office em um iPad ou Mac para teste
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5ec6924917f2351da77c8b9a84eb8de77b3864e
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348125"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="d940a-102">Realizar o sideload de suplementos do Office em um iPad ou Mac para teste</span><span class="sxs-lookup"><span data-stu-id="d940a-102">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="d940a-p101">Para ver como seu suplemento será executado no Office para iOS, você pode realizar o sideload do manifesto do suplemento em um iPad usando o iTunes ou diretamente no Office para Mac. Esta ação não permite definir pontos de interrupção e depurar o código do suplemento enquanto ele estiver sendo executado, mas é possível ver como ele se comporta e se a interface do usuário é utilizável e está sendo processada adequadamente.</span><span class="sxs-lookup"><span data-stu-id="d940a-p101">To see how your add-in will run in Office for iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office for Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span> 

## <a name="prerequisites-for-office-for-ios"></a><span data-ttu-id="d940a-105">Pré-requisitos do Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d940a-105">Prerequisites for Office for iOS</span></span>

- <span data-ttu-id="d940a-106">Um computador com Windows ou Mac com [iTunes](http://www.apple.com/itunes/download/) instalado.</span><span class="sxs-lookup"><span data-stu-id="d940a-106">A Windows or Mac computer with [iTunes](http://www.apple.com/itunes/download/) installed.</span></span>
    
- <span data-ttu-id="d940a-107">Um iPad executando o iOS 8.2 ou posterior com [Excel para iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) instalado e um cabo de sincronização.</span><span class="sxs-lookup"><span data-stu-id="d940a-107">An iPad running iOS 8.2 or later with [Excel for iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>
    
- <span data-ttu-id="d940a-108">O arquivo de manifesto .xml do suplemento que você deseja testar.</span><span class="sxs-lookup"><span data-stu-id="d940a-108">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="prerequisites-for-office-for-mac"></a><span data-ttu-id="d940a-109">Pré-requisitos do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="d940a-109">Prerequisites for Office for Mac</span></span>

- <span data-ttu-id="d940a-110">Um Mac executando o OS X v10.10 "Yosemite" ou posterior com [Office para Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) instalado.</span><span class="sxs-lookup"><span data-stu-id="d940a-110">A Mac running OS X v10.10 "Yosemite" or later with [Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>
    
- <span data-ttu-id="d940a-111">Word para Mac versão 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="d940a-111">Word for Mac version 15.18 (160109).</span></span>
   
- <span data-ttu-id="d940a-112">Excel para Mac versão 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="d940a-112">Excel for Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="d940a-113">PowerPoint para Mac versão 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="d940a-113">PowerPoint for Mac version 15.24 (160614)</span></span>
    
- <span data-ttu-id="d940a-114">O arquivo de manifesto .xml do suplemento que você deseja testar.</span><span class="sxs-lookup"><span data-stu-id="d940a-114">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="sideload-an-add-in-on-excel-or-word-for-ipad"></a><span data-ttu-id="d940a-115">Realizar o sideload de um suplemento no Excel ou no Word para iPad</span><span class="sxs-lookup"><span data-stu-id="d940a-115">Sideload an add-in on Excel or Word for iPad</span></span>

1. <span data-ttu-id="d940a-p102">Use um cabo de sincronização para conectar seu iPad ao computador. Se estiver conectando o iPad ao computador pela primeira vez, deverá responder **Confiar neste computador?** Escolha **Confiar** para continuar.</span><span class="sxs-lookup"><span data-stu-id="d940a-p102">Use a sync cable to connect your iPad to your computer. If you're connecting the iPad to your computer for the first time, you'll be prompted with  **Trust This Computer?**. Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="d940a-119">No iTunes, escolha o ícone do **iPad** abaixo da barra de menus.</span><span class="sxs-lookup"><span data-stu-id="d940a-119">In iTunes, choose the  **iPad** icon below the menu bar.</span></span>
    
    ![O ícone do iPad no iTunes](../images/ipad.png)

3. <span data-ttu-id="d940a-121">Em **Ajustes** no lado esquerdo do iTunes, escolha **Aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="d940a-121">Under  **Settings** on the left side of iTunes, choose **Apps**.</span></span>
    
    ![Configurações de aplicativos no iTunes](../images/file-settings-apps.png)

4. <span data-ttu-id="d940a-123">No lado direito do iTunes, role para baixo até **Compartilhamento de Arquivos**, e escolha **Excel** ou **Word** na coluna **Aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="d940a-123">On the right side of iTunes, scroll down to  **File Sharing**, and then choose  **Excel** or **Word** in the **Add-ins** column.</span></span>
    
    ![Compartilhamento de arquivos no iTunes](../images/file-sharing.png)

5. <span data-ttu-id="d940a-125">Na parte inferior da coluna Documentos do **Excel** ou do **Word**, escolha **Adicionar Arquivo** e selecione o arquivo de manifesto .xml do suplemento para o qual você deseja realizar sideload.</span><span class="sxs-lookup"><span data-stu-id="d940a-125">At the bottom of the  **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span> 
    
6. <span data-ttu-id="d940a-p103">Abra o aplicativo Excel ou Word no seu iPad. Se o aplicativo Excel ou Word já estiver em execução, escolha o botão **Início**, feche e reinicie o aplicativo.</span><span class="sxs-lookup"><span data-stu-id="d940a-p103">Open the Excel or Word app on your iPad. If the Excel or Word app is already running, choose the  **Home** button, and then close and restart the app.</span></span>
    
7. <span data-ttu-id="d940a-128">Abra um documento.</span><span class="sxs-lookup"><span data-stu-id="d940a-128">Open a document.</span></span>
    
8. <span data-ttu-id="d940a-129">Escolha **Suplementos** na guia **Inserir**. O suplemento com sideload está disponível para inserção no cabeçalho **Desenvolvedor** da interface de usuário **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="d940a-129">Choose  **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>
    
    ![Inserir Suplementos no aplicativo do Excel](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-on-office-for-mac"></a><span data-ttu-id="d940a-131">Realizar sideload de um suplemento no Office para Mac</span><span class="sxs-lookup"><span data-stu-id="d940a-131">Sideload an add-in on Office for Mac</span></span>

> [!NOTE]
> <span data-ttu-id="d940a-132">Para realizar o sideload de um suplemento do Outlook para Mac, confira [Realizar sideload de suplementos do Outlook para teste](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="d940a-132">To sideload Outlook 2016 for Mac add-in, see [Sideload Outlook add-ins for testing](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

1. <span data-ttu-id="d940a-p104">Abra o **Terminal** e navegue até uma das pastas a seguir, onde você salvará o arquivo de manifesto do suplemento. Se a pasta `wef` não existir em seu computador, crie-a.</span><span class="sxs-lookup"><span data-stu-id="d940a-p104">Open  **Terminal** and go to one of the following folders where you'll save your add-in's manifest file. If the `wef` folder doesn't exist on your computer, create it.</span></span>
    
    - <span data-ttu-id="d940a-135">Para o Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="d940a-135">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span></span>    
    - <span data-ttu-id="d940a-136">Para o Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="d940a-136">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span></span>
    - <span data-ttu-id="d940a-137">Para o PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="d940a-137">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span></span>
    
2. <span data-ttu-id="d940a-p105">Abra a pasta no **Finder** usando o comando `open .` (incluindo o ponto final). Copie o arquivo de manifesto do suplemento nessa pasta.</span><span class="sxs-lookup"><span data-stu-id="d940a-p105">Open the folder in  **Finder** using the command `open .` (including the period or dot). Copy your add-in's manifest file to this folder.</span></span>
    
    ![Pasta Wef no Office para Mac](../images/all-my-files.png)

3. <span data-ttu-id="d940a-p106">Abra o Word e, em seguida, abra um documento. Reinicie o Word se ele já estiver em execução.</span><span class="sxs-lookup"><span data-stu-id="d940a-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>
    
4. <span data-ttu-id="d940a-143">No Word, escolha **Inserir** > **Suplementos** > **Meus Suplementos** (menu suspenso) e escolha o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="d940a-143">In Word, choose  **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>
    
    ![Meus Suplementos no Office para Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="d940a-p107">Os suplementos para os quais foi realizado o sideload não aparecerão na caixa de diálogo Meus Suplementos. Eles só ficam visíveis dentro do menu suspenso (pequena seta para baixo à direita de Meus Suplementos na guia **Inserir**). Esse suplementos são exibidos na lista sob o título **Suplementos do Desenvolvedor** nesse menu.</span><span class="sxs-lookup"><span data-stu-id="d940a-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span> 
    
5. <span data-ttu-id="d940a-148">Verifique se o seu suplemento é exibido no Word.</span><span class="sxs-lookup"><span data-stu-id="d940a-148">Verify that your add-in is displayed in Word.</span></span>
    
    ![Suplemento do Office exibido no Office para Mac](../images/lorem-ipsum-wikipedia.png)
    
    > [!NOTE]
    > <span data-ttu-id="d940a-p108">Por motivos de desempenho, o Office para Mac costuma armazenar os suplementos no cache. Se você precisar forçar um novo carregamento do suplemento durante o desenvolvimento, pode limpar a pasta `Users/<usr>/Library/Containers/com.Microsoft.OsfWebHost/Data/`. Se essa pasta não existir, exclua os arquivos da pasta `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`.</span><span class="sxs-lookup"><span data-stu-id="d940a-p108">Add-ins are cached often in Office for Mac, for performance reasons. If you need to force a reload of your add-in while you're developing it, you can clear the `Users/<usr>/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="d940a-153">Confira também</span><span class="sxs-lookup"><span data-stu-id="d940a-153">See also</span></span>

- [<span data-ttu-id="d940a-154">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="d940a-154">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
    
