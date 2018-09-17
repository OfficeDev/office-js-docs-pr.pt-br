---
title: Depurar suplementos no Office Online
description: Como usar o Office Online para testar e depurar seus suplementos.
ms.date: 03/14/2018
ms.openlocfilehash: ee458352c78a3bb7828e66df9fcde12958f3df93
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945761"
---
# <a name="debug-add-ins-in-office-online"></a><span data-ttu-id="1f446-103">Depurar suplementos no Office Online</span><span class="sxs-lookup"><span data-stu-id="1f446-103">Debug add-ins in Office Online</span></span>


<span data-ttu-id="1f446-104">Você pode compilar e depurar suplementos em um computador que não esteja executando o Windows ou o cliente de desktop do Office&mdash;por exemplo, se você estiver desenvolvendo em um Mac.</span><span class="sxs-lookup"><span data-stu-id="1f446-104">You can build and debug add-ins on a computer that isn't running Windows or the Office 2013 or Office 2016 desktop client - for example, if you're developing on a Mac. This article describes how to use Office Online to test and debug your add-ins.</span></span> <span data-ttu-id="1f446-105">Este artigo descreve como usar o Office Online para testar e depurar seus suplementos.</span><span class="sxs-lookup"><span data-stu-id="1f446-105">How to use Office Online to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="1f446-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="1f446-106">Prerequisites</span></span>

<span data-ttu-id="1f446-107">Para começar:</span><span class="sxs-lookup"><span data-stu-id="1f446-107">To get started:</span></span>

- <span data-ttu-id="1f446-108">Obtenha uma conta de desenvolvedor do Office 365, se já não tiver uma, ou o acesso a um site do SharePoint.</span><span class="sxs-lookup"><span data-stu-id="1f446-108">Get an Office 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>
    
  > [!NOTE]
  > <span data-ttu-id="1f446-109">Para se inscrever para uma assinatura gratuita de desenvolvedor do Office 365, ingresse no [Programa de Desenvolvedor do Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="1f446-109">To sign up for a free Office 365 developer subscription, join our [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span> <span data-ttu-id="1f446-110">Confira o passo a passo de como participar do Programa para Desenvolvedores do Office 365, entre e configure sua assinatura na [documentação do Programa para Desenvolvedores do Office 365](https://docs.microsoft.com/office/developer-program/office-365-developer-program).</span><span class="sxs-lookup"><span data-stu-id="1f446-110">See the [Office 365 Developer Program documentation](https://docs.microsoft.com/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Office 365 Developer Program and sign up and configure your subscription.</span></span>
     
- <span data-ttu-id="1f446-p103">Configure um catálogo de suplementos no Office 365 (SharePoint Online). Um catálogo de suplementos é um conjunto de sites dedicado no SharePoint Online que hospeda bibliotecas de documentos para suplementos do Office. Se você tiver seu próprio site do SharePoint, pode configurar uma biblioteca de documentos do catálogo de suplementos. Para saber mais, confira [Publicar suplementos de conteúdo e de painel de tarefas em um catálogo de suplementos no SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span><span class="sxs-lookup"><span data-stu-id="1f446-p103">Set up an add-in catalog on Office 365 (SharePoint Online). An add-in catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an add-in catalog document library. For more information, see [Publish task pane and content add-ins to an add-in catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a><span data-ttu-id="1f446-114">Depurar seu suplemento do Excel Online ou do Word Online</span><span class="sxs-lookup"><span data-stu-id="1f446-114">Debug your add-in from Excel Online or Word Online</span></span>

<span data-ttu-id="1f446-115">Para depurar seu suplemento usando o Office Online:</span><span class="sxs-lookup"><span data-stu-id="1f446-115">To debug your add-in by using Office Online:</span></span>

1. <span data-ttu-id="1f446-116">Implante o suplemento em um servidor que dê suporte a SSL.</span><span class="sxs-lookup"><span data-stu-id="1f446-116">Deploy your add-in to a server that supports SSL.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="1f446-117">Recomendamos que você use o [gerador Yeoman](https://github.com/OfficeDev/generator-office) para criar e hospedar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="1f446-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>
     
2. <span data-ttu-id="1f446-p104">No seu [arquivo de manifesto de suplemento](../develop/add-in-manifests.md), atualize o valor do elemento **SourceLocation** para incluir um URI absoluto, em vez de relativo. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="1f446-p104">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:</span></span>
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. <span data-ttu-id="1f446-120">Carregue o manifesto na biblioteca de Suplementos do Office no catálogo de suplementos no SharePoint.</span><span class="sxs-lookup"><span data-stu-id="1f446-120">Upload the manifest to the Office Add-ins library in the add-in catalog on SharePoint.</span></span>
    
4. <span data-ttu-id="1f446-121">Inicie o Excel Online ou o Word Online do inicializador de aplicativos no Office 365 e abra um novo documento.</span><span class="sxs-lookup"><span data-stu-id="1f446-121">Launch Excel Online or Word Online from the app launcher in Office 365, and open a new document.</span></span>
    
5. <span data-ttu-id="1f446-122">Na guia Inserir, escolha **Meus Suplementos** ou **Suplementos do Office** para inserir seu suplemento e testá-lo no aplicativo.</span><span class="sxs-lookup"><span data-stu-id="1f446-122">On the Insert tab, choose  **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>
    
6. <span data-ttu-id="1f446-123">Use seu depurador de navegador favorito para depurar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="1f446-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="1f446-124">Possíveis problemas</span><span class="sxs-lookup"><span data-stu-id="1f446-124">Potential issues</span></span>    

<span data-ttu-id="1f446-125">A seguir apresentamos alguns problemas que você pode encontrar ao depurar:</span><span class="sxs-lookup"><span data-stu-id="1f446-125">The following are some issues that you might encounter as you debug:</span></span>
    
- <span data-ttu-id="1f446-126">Alguns erros de JavaScript que você vê podem vir do Office Online.</span><span class="sxs-lookup"><span data-stu-id="1f446-126">Some JavaScript errors that you see might originate from Office Online.</span></span>
      
- <span data-ttu-id="1f446-127">O navegador pode mostrar um erro de certificado inválido que você deve ignorar.</span><span class="sxs-lookup"><span data-stu-id="1f446-127">The browser might show an invalid certificate error that you will need to bypass.</span></span>
      
- <span data-ttu-id="1f446-128">Se você definir pontos de interrupção no seu código, o Office Online pode lançar uma mensagem de erro indicando que não é possível salvar.</span><span class="sxs-lookup"><span data-stu-id="1f446-128">If you set breakpoints in your code, Office Online might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="1f446-129">Veja também</span><span class="sxs-lookup"><span data-stu-id="1f446-129">See also</span></span>

- [<span data-ttu-id="1f446-130">Práticas recomendadas para o desenvolvimento de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1f446-130">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="1f446-131">Políticas de validação do AppSource</span><span class="sxs-lookup"><span data-stu-id="1f446-131">AppSource validation policies</span></span>](https://docs.microsoft.com/office/dev/store/validation-policies)  
- [<span data-ttu-id="1f446-132">Criar aplicativos e suplementos eficazes para o AppSource</span><span class="sxs-lookup"><span data-stu-id="1f446-132">Create effective AppSource apps and add-ins</span></span>](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)  
- [<span data-ttu-id="1f446-133">Solucionar erros de usuários com suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1f446-133">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
    
