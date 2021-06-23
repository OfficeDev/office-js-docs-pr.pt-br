---
title: Instale a última versão do Office
description: Informações sobre como desativar essa opção para obter as versões mais recentes do Office.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: d72012e3c2e642c74d8573c4d9bb3b29a8fc0274
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076017"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="93918-103">Instale a última versão do Office</span><span class="sxs-lookup"><span data-stu-id="93918-103">Install the latest version of Office</span></span>

<span data-ttu-id="93918-104">Novos recursos de desenvolvedor, inclusive os que ainda estão na visualização, são fornecidos primeiro aos assinantes que aceitam obter as últimas compilações do Office.</span><span class="sxs-lookup"><span data-stu-id="93918-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span>

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="93918-105">Aceitar para receber as versões mais recentes</span><span class="sxs-lookup"><span data-stu-id="93918-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="93918-106">Aceitar para receber as versões mais recentes do Office:</span><span class="sxs-lookup"><span data-stu-id="93918-106">To opt in to getting the latest builds of Office:</span></span>

- <span data-ttu-id="93918-107">Se você for um Microsoft 365 Family, pessoal ou assinante da Universidade, consulte Ser um Office [Insider](https://insider.office.com).</span><span class="sxs-lookup"><span data-stu-id="93918-107">If you're a Microsoft 365 Family, Personal, or University subscriber, see [Be an Office Insider](https://insider.office.com).</span></span>
- <span data-ttu-id="93918-108">Se você for um cliente Microsoft 365 Apps para Pequenos e Médios negócios, consulte [Install the First Release build for Microsoft 365 Apps para Pequenos e Médios negócios customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="93918-108">If you're a Microsoft 365 Apps for business customer, see [Install the First Release build for Microsoft 365 Apps for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="93918-109">Se você estiver executando o Office em um Mac:</span><span class="sxs-lookup"><span data-stu-id="93918-109">If you're running Office on a Mac:</span></span>
  - <span data-ttu-id="93918-110">Abra um aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="93918-110">Start an Office application.</span></span>
  - <span data-ttu-id="93918-111">Selecione **Verificar Atualizações** no menu Ajuda.</span><span class="sxs-lookup"><span data-stu-id="93918-111">Select **Check for Updates** on the Help menu.</span></span>
  - <span data-ttu-id="93918-112">Na caixa Microsoft AutoUpdate, marque a caixa para participar do programa Office Insider.</span><span class="sxs-lookup"><span data-stu-id="93918-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span>

## <a name="get-the-latest-build"></a><span data-ttu-id="93918-113">Obtenha a versão mais recente:</span><span class="sxs-lookup"><span data-stu-id="93918-113">Get the latest build</span></span>

<span data-ttu-id="93918-114">Para receber as versões mais recentes do Office:</span><span class="sxs-lookup"><span data-stu-id="93918-114">To get the latest build of Office:</span></span>

1. <span data-ttu-id="93918-115">Baixar a [Ferramenta de Implantação do Office](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="93918-115">Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span>
2. <span data-ttu-id="93918-p101">Execute a ferramenta. Isso extrai estes dois arquivos: Setup.exe e configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="93918-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="93918-118">Substitua o arquivo configuration.xml pelo [Arquivo de Configuração do Primeiro Lançamento](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="93918-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="93918-119">Execute o seguinte comando como administrador: `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="93918-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span>

> [!NOTE]
> <span data-ttu-id="93918-120">O comando pode demorar muito para ser executado sem indicar o progresso.</span><span class="sxs-lookup"><span data-stu-id="93918-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="93918-121">Quando o processo de instalação for concluído, você terá os últimos aplicativos do Office instalados.</span><span class="sxs-lookup"><span data-stu-id="93918-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="93918-122">Para verificar se você tem a última compilação, vá até **arquivo** > **conta** em qualquer aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="93918-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="93918-123">Em Atualizações do Office, você verá o rótulo (Office Insiders) acima do número de versão.</span><span class="sxs-lookup"><span data-stu-id="93918-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Uma captura de tela que mostra informações do produto com o rótulo Office Insiders.](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="93918-125">Builds mínimos do Office para conjuntos de requisitos de API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="93918-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="93918-126">Para saber mais sobre os builds mínimos de produtos para cada plataforma dos conjuntos de requisitos de API, confira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="93918-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="93918-127">Conjuntos de requisitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="93918-127">Excel JavaScript API requirement sets</span></span>](../reference/requirement-sets/excel-api-requirement-sets.md)
- [<span data-ttu-id="93918-128">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="93918-128">OneNote JavaScript API requirement sets</span></span>](../reference/requirement-sets/onenote-api-requirement-sets.md)
- [<span data-ttu-id="93918-129">Conjuntos de requisitos de API JavaScript do Outlook</span><span class="sxs-lookup"><span data-stu-id="93918-129">Outlook JavaScript API requirement sets</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [<span data-ttu-id="93918-130">Conjuntos de requisitos de API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="93918-130">PowerPoint JavaScript API requirement sets</span></span>](../reference/requirement-sets/powerpoint-api-requirement-sets.md)
- [<span data-ttu-id="93918-131">Conjuntos de requisitos da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="93918-131">Word JavaScript API requirement sets</span></span>](../reference/requirement-sets/word-api-requirement-sets.md)
- [<span data-ttu-id="93918-132">Conjuntos de requisitos da API de Caixa de Diálogo</span><span class="sxs-lookup"><span data-stu-id="93918-132">Dialog API requirement sets</span></span>](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [<span data-ttu-id="93918-133">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="93918-133">Office Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
