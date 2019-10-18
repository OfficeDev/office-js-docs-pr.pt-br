---
title: Instale a última versão do Office
description: Informações sobre como desativar essa opção para obter as versões mais recentes do Office.
ms.date: 07/17/2019
localization_priority: Priority
ms.openlocfilehash: 2308bf8f825201f1074d56bba2ddc67c32ba7f24
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940623"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="fa7cf-103">Instale a última versão do Office</span><span class="sxs-lookup"><span data-stu-id="fa7cf-103">Install the latest version of Office</span></span>

<span data-ttu-id="fa7cf-104">Novos recursos de desenvolvedor, inclusive os que ainda estão na visualização, são fornecidos primeiro aos assinantes que aceitam obter as últimas compilações do Office.</span><span class="sxs-lookup"><span data-stu-id="fa7cf-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span>

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="fa7cf-105">Aceitar para receber as versões mais recentes</span><span class="sxs-lookup"><span data-stu-id="fa7cf-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="fa7cf-106">Aceitar para receber as versões mais recentes do Office:</span><span class="sxs-lookup"><span data-stu-id="fa7cf-106">To opt in to getting the latest builds of Office:</span></span>

- <span data-ttu-id="fa7cf-107">Se você é assinante do Office 365 Home, Personal ou University, confira [Ser um Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="fa7cf-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="fa7cf-108">Se você for um cliente corporativo do Office 365, confira [Instalar a versão de Primeiro Lançamento do Office 365 para clientes corporativos](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="fa7cf-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="fa7cf-109">Se você estiver executando o Office em um Mac:</span><span class="sxs-lookup"><span data-stu-id="fa7cf-109">If you're running Office on a Mac:</span></span>
  - <span data-ttu-id="fa7cf-110">Abra um aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="fa7cf-110">Start the Office application.</span></span>
  - <span data-ttu-id="fa7cf-111">Selecione **Verificar Atualizações** no menu Ajuda.</span><span class="sxs-lookup"><span data-stu-id="fa7cf-111">Select **Check for Updates** on the Help menu.</span></span>
  - <span data-ttu-id="fa7cf-112">Na caixa Microsoft AutoUpdate, marque a caixa para participar do programa Office Insider.</span><span class="sxs-lookup"><span data-stu-id="fa7cf-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span>

## <a name="get-the-latest-build"></a><span data-ttu-id="fa7cf-113">Obtenha a versão mais recente:</span><span class="sxs-lookup"><span data-stu-id="fa7cf-113">Get the latest build</span></span>

<span data-ttu-id="fa7cf-114">Para receber as versões mais recentes do Office:</span><span class="sxs-lookup"><span data-stu-id="fa7cf-114">To get the latest build of Office:</span></span>

1. <span data-ttu-id="fa7cf-115">Baixar a [Ferramenta de Implantação do Office](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="fa7cf-115">Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span>
2. <span data-ttu-id="fa7cf-p101">Execute a ferramenta. Isso extrai estes dois arquivos: Setup.exe e configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="fa7cf-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="fa7cf-118">Substitua o arquivo configuration.xml pelo [Arquivo de Configuração do Primeiro Lançamento](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="fa7cf-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="fa7cf-119">Execute o seguinte comando como administrador: `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="fa7cf-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span>

> [!NOTE]
> <span data-ttu-id="fa7cf-120">O comando pode demorar muito para ser executado sem indicar o progresso.</span><span class="sxs-lookup"><span data-stu-id="fa7cf-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="fa7cf-121">Quando o processo de instalação for concluído, você terá os últimos aplicativos do Office instalados.</span><span class="sxs-lookup"><span data-stu-id="fa7cf-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="fa7cf-122">Para verificar se você tem a última compilação, vá até **arquivo** > **conta** em qualquer aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="fa7cf-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="fa7cf-123">Em Atualizações do Office, você verá o rótulo (Office Insiders) acima do número de versão.</span><span class="sxs-lookup"><span data-stu-id="fa7cf-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Uma captura de tela que mostra informações do produto com o rótulo Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="fa7cf-125">Builds mínimos do Office para conjuntos de requisitos de API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="fa7cf-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="fa7cf-126">Para saber mais sobre os builds mínimos de produtos para cada plataforma dos conjuntos de requisitos de API, confira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="fa7cf-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="fa7cf-127">Conjuntos de requisitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="fa7cf-127">Excel JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)
- [<span data-ttu-id="fa7cf-128">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="fa7cf-128">OneNote JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets)
- [<span data-ttu-id="fa7cf-129">Conjuntos de requisitos de API JavaScript do Outlook</span><span class="sxs-lookup"><span data-stu-id="fa7cf-129">Outlook JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
- [<span data-ttu-id="fa7cf-130">Conjuntos de requisitos de API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="fa7cf-130">OneNote JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets)
- [<span data-ttu-id="fa7cf-131">Conjuntos de requisitos da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="fa7cf-131">Word JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets)
- [<span data-ttu-id="fa7cf-132">Conjuntos de requisitos da API de Caixa de Diálogo</span><span class="sxs-lookup"><span data-stu-id="fa7cf-132">Dialog API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="fa7cf-133">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="fa7cf-133">Office Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
