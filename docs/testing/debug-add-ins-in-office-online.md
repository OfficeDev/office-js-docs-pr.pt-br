---
title: Depurar suplementos no Office na Web
description: Como usar o Office na Web para testar e depurar seus suplementos.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 1ec844af555ab55d7592a9d764c6173ea9780016
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350179"
---
# <a name="debug-add-ins-in-office-on-the-web"></a><span data-ttu-id="15df5-103">Depurar suplementos no Office na Web</span><span class="sxs-lookup"><span data-stu-id="15df5-103">Debug add-ins in Office on the web</span></span>

<span data-ttu-id="15df5-104">Você pode criar e depurar suplementos em um computador que não esteja executando o Windows ou os clientes de área de trabalho do Office 2013 ou do Office 2016, por exemplo, se você estiver desenvolvendo no Mac. Este artigo descreve como usar o Office Online para testar e depurar seus suplementos.</span><span class="sxs-lookup"><span data-stu-id="15df5-104">You can build and debug add-ins on a computer that isn't running Windows or the Office desktop client&mdash;for example, if you're developing on a Mac.</span></span> <span data-ttu-id="15df5-105">Este artigo descreve como usar o Office na Web para testar e depurar seus suplementos.</span><span class="sxs-lookup"><span data-stu-id="15df5-105">This article describes how to use Office on the web to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="15df5-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="15df5-106">Prerequisites</span></span>

<span data-ttu-id="15df5-107">Para começar:</span><span class="sxs-lookup"><span data-stu-id="15df5-107">To get started:</span></span>

- <span data-ttu-id="15df5-108">Obter uma Microsoft 365 de desenvolvedor se você ainda não tiver uma ou tiver acesso a um SharePoint site.</span><span class="sxs-lookup"><span data-stu-id="15df5-108">Get a Microsoft 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>

  > [!NOTE]
  > <span data-ttu-id="15df5-109">Para obter uma assinatura de desenvolvedor renovável de 90 dias Microsoft 365, participe do nosso [programa Microsoft 365 desenvolvedor.](https://developer.microsoft.com/office/dev-program)</span><span class="sxs-lookup"><span data-stu-id="15df5-109">To get a free, 90-day renewable Microsoft 365 developer subscription, join our [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span> <span data-ttu-id="15df5-110">Consulte a [documentação Microsoft 365](/office/developer-program/office-365-developer-program) do programa de desenvolvedor para obter instruções passo a passo sobre como ingressar no programa Microsoft 365 desenvolvedor e configurar sua assinatura.</span><span class="sxs-lookup"><span data-stu-id="15df5-110">See the [Microsoft 365 developer program documentation](/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Microsoft 365 developer program and configure your subscription.</span></span>

- <span data-ttu-id="15df5-111">Configurar um catálogo de aplicativos no SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="15df5-111">Set up an app catalog on SharePoint Online.</span></span> <span data-ttu-id="15df5-112">Um catálogo de aplicativos é um conjunto de sites dedicado no SharePoint Online que hospeda bibliotecas de documentos para Office de complementos. Se você tiver seu próprio site SharePoint, poderá configurar uma biblioteca de documentos do catálogo de aplicativos.</span><span class="sxs-lookup"><span data-stu-id="15df5-112">An app catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an app catalog document library.</span></span> <span data-ttu-id="15df5-113">Para obter mais informações, [consulte Publish task pane and content add-ins to an app catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span><span class="sxs-lookup"><span data-stu-id="15df5-113">For more information, see [Publish task pane and content add-ins to an app catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a><span data-ttu-id="15df5-114">Depurar seu suplemento do Excel ou Word na Web</span><span class="sxs-lookup"><span data-stu-id="15df5-114">Debug your add-in from Excel or Word on the web</span></span>

<span data-ttu-id="15df5-115">Para depurar seu suplemento usando o Office na Web:</span><span class="sxs-lookup"><span data-stu-id="15df5-115">To debug your add-in by using Office on the web:</span></span>

1. <span data-ttu-id="15df5-116">Implante o suplemento em um servidor que dê suporte a SSL.</span><span class="sxs-lookup"><span data-stu-id="15df5-116">Deploy your add-in to a server that supports SSL.</span></span>

    > [!NOTE]
    > <span data-ttu-id="15df5-117">Recomendamos que você use o [gerador Yeoman](https://github.com/OfficeDev/generator-office) para criar e hospedar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="15df5-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>

2. <span data-ttu-id="15df5-p104">No seu [arquivo de manifesto de suplemento](../develop/add-in-manifests.md), atualize o valor do elemento **SourceLocation** para incluir um URI absoluto, em vez de relativo. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="15df5-p104">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:</span></span>

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. <span data-ttu-id="15df5-120">Carregue o manifesto para a biblioteca de suplementos do Office no catálogo de aplicativos no SharePoint.</span><span class="sxs-lookup"><span data-stu-id="15df5-120">Upload the manifest to the Office Add-ins library in the app catalog on SharePoint.</span></span>

4. <span data-ttu-id="15df5-121">Iniciar Excel ou Word na Web do launcher de aplicativos no Microsoft 365 e abra um novo documento.</span><span class="sxs-lookup"><span data-stu-id="15df5-121">Launch Excel or Word on the web from the app launcher in Microsoft 365, and open a new document.</span></span>

5. <span data-ttu-id="15df5-122">Na guia Inserir, escolha **Meus Complementos** ou **Office-ins** para inserir seu complemento e testá-lo no aplicativo.</span><span class="sxs-lookup"><span data-stu-id="15df5-122">On the Insert tab, choose **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>

6. <span data-ttu-id="15df5-123">Use seu depurador de navegador favorito para depurar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="15df5-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="15df5-124">Possíveis problemas</span><span class="sxs-lookup"><span data-stu-id="15df5-124">Potential issues</span></span>

<span data-ttu-id="15df5-125">A seguir estão alguns problemas que você pode encontrar ao depurar.</span><span class="sxs-lookup"><span data-stu-id="15df5-125">The following are some issues that you might encounter as you debug.</span></span>

- <span data-ttu-id="15df5-126">Alguns erros de JavaScript que você vê podem vir do Office na Web.</span><span class="sxs-lookup"><span data-stu-id="15df5-126">Some JavaScript errors that you see might originate from Office on the web.</span></span>

- <span data-ttu-id="15df5-127">O navegador pode mostrar um erro de certificado inválido que você deve ignorar.</span><span class="sxs-lookup"><span data-stu-id="15df5-127">The browser might show an invalid certificate error that you will need to bypass.</span></span> <span data-ttu-id="15df5-128">O processo para fazer isso varia com o navegador e as interfaces de usuário dos vários navegadores para fazer essa alteração periodicamente.</span><span class="sxs-lookup"><span data-stu-id="15df5-128">The process for doing this varies with the browser and the various browsers' UIs for doing this change periodically.</span></span> <span data-ttu-id="15df5-129">Você deve pesquisar na ajuda do navegador ou pesquisar online para obter instruções.</span><span class="sxs-lookup"><span data-stu-id="15df5-129">You should search the browser's help or search online for instructions.</span></span> <span data-ttu-id="15df5-130">(Por exemplo, procure por "Aviso de certificado inválido do Microsoft Edge".) A maioria dos navegadores terá um link na página de aviso que permite que você clique na página do suplemento.</span><span class="sxs-lookup"><span data-stu-id="15df5-130">(For example, search for "Microsoft Edge invalid certificate warning".) Most browsers will have a link on the warning page that enables you to click through to the add-in page.</span></span> <span data-ttu-id="15df5-131">Por exemplo, o Microsoft Edge possui um link "Ir para a página da Web (não recomendado)".</span><span class="sxs-lookup"><span data-stu-id="15df5-131">For example, Microsoft Edge has a link "Go on to the webpage (Not recommended)".</span></span> <span data-ttu-id="15df5-132">Mas você geralmente terá que passar por este link toda vez que o suplemento for recarregado.</span><span class="sxs-lookup"><span data-stu-id="15df5-132">But you will usually have to go through this link every time the add-in reloads.</span></span> <span data-ttu-id="15df5-133">Para um bypass mais duradouro, consulte a ajuda, como sugerido.</span><span class="sxs-lookup"><span data-stu-id="15df5-133">For a longer lasting bypass, see the help as suggested.</span></span>

- <span data-ttu-id="15df5-134">Se você definir pontos de interrupção no seu código, o Office na Web pode lançar uma mensagem de erro indicando que não é possível salvar.</span><span class="sxs-lookup"><span data-stu-id="15df5-134">If you set breakpoints in your code, Office on the web might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="15df5-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="15df5-135">See also</span></span>

- [<span data-ttu-id="15df5-136">Práticas recomendadas para o desenvolvimento de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="15df5-136">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="15df5-137">Políticas de validação do AppSource</span><span class="sxs-lookup"><span data-stu-id="15df5-137">AppSource validation policies</span></span>](/legal/marketplace/certification-policies)  
- [<span data-ttu-id="15df5-138">Criar aplicativos e suplementos eficazes para o AppSource</span><span class="sxs-lookup"><span data-stu-id="15df5-138">Create effective AppSource apps and add-ins</span></span>](/office/dev/store/create-effective-office-store-listings)  
- [<span data-ttu-id="15df5-139">Solucionar erros de usuários com suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="15df5-139">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
