---
title: Instalar a última versão do Office
description: Informações sobre como aceitar para obter as versões mais recentes do Office.
ms.date: 12/04/2017
ms.openlocfilehash: 14e26d9fa9f7ec3b2724cbf2e9787cde9dbe4094
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943877"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="6a01a-103">Instalar a última versão do Office</span><span class="sxs-lookup"><span data-stu-id="6a01a-103">Install the latest version of Office</span></span>

<span data-ttu-id="6a01a-104">Novos recursos de desenvolvedor, inclusive os que ainda estão na visualização, são fornecidos primeiro aos assinantes que aceitam obter as últimas versões do Office.</span><span class="sxs-lookup"><span data-stu-id="6a01a-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="6a01a-105">Aceitar para receber as versões mais recentes</span><span class="sxs-lookup"><span data-stu-id="6a01a-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="6a01a-106">Aceitar para receber as versões mais recentes do Office:</span><span class="sxs-lookup"><span data-stu-id="6a01a-106">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="6a01a-107">Se você é assinante do Office 365 Home, Personal ou University, confira [Ser um Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="6a01a-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="6a01a-108">Se você for um cliente corporativo do Office 365, confira [Instalar a versão de Primeiro Lançamento do Office 365 para clientes corporativos](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="6a01a-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="6a01a-109">Se você estiver executando o Office em um Mac:</span><span class="sxs-lookup"><span data-stu-id="6a01a-109">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="6a01a-110">Inicie um programa do Office para Mac.</span><span class="sxs-lookup"><span data-stu-id="6a01a-110">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="6a01a-111">Selecione **Verificar Atualizações** no menu Ajuda.</span><span class="sxs-lookup"><span data-stu-id="6a01a-111">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="6a01a-112">Na caixa Microsoft AutoUpdate, marque a caixa para participar do programa Office Insider.</span><span class="sxs-lookup"><span data-stu-id="6a01a-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="6a01a-113">Obter a versão mais recente</span><span class="sxs-lookup"><span data-stu-id="6a01a-113">Get the latest build</span></span>

<span data-ttu-id="6a01a-114">Para obter as versões mais recentes do Office:</span><span class="sxs-lookup"><span data-stu-id="6a01a-114">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="6a01a-115">Baixe a [Ferramenta de Implantação do Office](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="6a01a-115">Download the Office Deployment Tool</span></span> 
2. <span data-ttu-id="6a01a-p101">Execute a ferramenta. Isso extrai estes dois arquivos: Setup.exe e configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="6a01a-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="6a01a-118">Substitua o arquivo configuration.xml pelo [Arquivo de Configuração do Primeiro Lançamento](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="6a01a-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="6a01a-119">Execute o seguinte comando como administrador:  `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="6a01a-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="6a01a-120">O comando pode demorar muito para ser executado sem indicar o progresso.</span><span class="sxs-lookup"><span data-stu-id="6a01a-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="6a01a-121">Quando o processo de instalação for concluído, você terá os aplicativos do Office mais recentes instalados.</span><span class="sxs-lookup"><span data-stu-id="6a01a-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="6a01a-122">Para verificar se você tem a versão mais recente, vá para o **Arquivo** > **Conta** de qualquer aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="6a01a-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="6a01a-123">Em atualizações do Office, você verá o rótulo (Office Insiders) acima do número de versão.</span><span class="sxs-lookup"><span data-stu-id="6a01a-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Uma captura de tela que mostra informações do produto com o rótulo Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="6a01a-125">Builds mínimos do Office para conjuntos de requisitos de API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="6a01a-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="6a01a-126">Para saber mais sobre os builds mínimos de produtos para cada plataforma dos conjuntos de requisitos de API, confira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="6a01a-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="6a01a-127">Conjuntos de requisitos da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="6a01a-127">Word JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets?view=office-js)
- [<span data-ttu-id="6a01a-128">Conjuntos de requisitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="6a01a-128">Excel JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js)
- [<span data-ttu-id="6a01a-129">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="6a01a-129">OneNote JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [<span data-ttu-id="6a01a-130">Conjuntos de requisitos da API de caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="6a01a-130">Dialog API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [<span data-ttu-id="6a01a-131">Conjuntos de requisitos de API comum do Office</span><span class="sxs-lookup"><span data-stu-id="6a01a-131">Office common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js)
