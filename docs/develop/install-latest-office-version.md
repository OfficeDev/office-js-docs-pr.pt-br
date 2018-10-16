---
title: Instalar a última versão do Office
description: Informações sobre como aceitar para obter as versões mais recentes do Office.
ms.date: 12/04/2017
ms.openlocfilehash: 0e6e147144757004575fa086e1066b7cdf133ee8
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505787"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="b71e6-103">Instalar a última versão do Office</span><span class="sxs-lookup"><span data-stu-id="b71e6-103">Install the latest version of Office</span></span>

<span data-ttu-id="b71e6-104">Novos recursos de desenvolvedor, inclusive os que ainda estão na visualização, são fornecidos primeiro aos assinantes que aceitam obter as últimas versões do Office.</span><span class="sxs-lookup"><span data-stu-id="b71e6-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="b71e6-105">Aceitar para receber as versões mais recentes</span><span class="sxs-lookup"><span data-stu-id="b71e6-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="b71e6-106">Aceitar para receber as versões mais recentes do Office:</span><span class="sxs-lookup"><span data-stu-id="b71e6-106">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="b71e6-107">Se você é assinante do Office 365 Home, Personal ou University, confira [Ser um Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="b71e6-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="b71e6-108">Se você for um cliente corporativo do Office 365, confira [Instalar a versão de Primeiro Lançamento do Office 365 para clientes corporativos](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="b71e6-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="b71e6-109">Se você estiver executando o Office em um Mac:</span><span class="sxs-lookup"><span data-stu-id="b71e6-109">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="b71e6-110">Inicie um programa do Office para Mac.</span><span class="sxs-lookup"><span data-stu-id="b71e6-110">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="b71e6-111">Selecione **Verificar Atualizações** no menu Ajuda.</span><span class="sxs-lookup"><span data-stu-id="b71e6-111">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="b71e6-112">Na caixa Microsoft AutoUpdate, marque a caixa para participar do programa Office Insider.</span><span class="sxs-lookup"><span data-stu-id="b71e6-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="b71e6-113">Obter a versão mais recente</span><span class="sxs-lookup"><span data-stu-id="b71e6-113">Get the latest build</span></span>

<span data-ttu-id="b71e6-114">Para obter as versões mais recentes do Office:</span><span class="sxs-lookup"><span data-stu-id="b71e6-114">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="b71e6-115">Baixe a [Ferramenta de Implantação do Office](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="b71e6-115">Download the Office Deployment Tool</span></span> 
2. <span data-ttu-id="b71e6-p101">Execute a ferramenta. Isso extrai estes dois arquivos: Setup.exe e configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="b71e6-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="b71e6-118">Substitua o arquivo configuration.xml pelo [Arquivo de Configuração do Primeiro Lançamento](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="b71e6-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="b71e6-119">Execute o seguinte comando como administrador:  `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="b71e6-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="b71e6-120">O comando pode demorar muito para ser executado sem indicação do andamento.</span><span class="sxs-lookup"><span data-stu-id="b71e6-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="b71e6-p102">Quando o processo de instalação for concluído, você terá os aplicativos mais recentes do Office instalados. Para verificar se você tem a compilação mais recente, vá para **Arquivo** > **Conta** em qualquer aplicativo do Office. Em Atualizações do Office, você verá o rótulo (Office Insiders) acima do número de versão.</span><span class="sxs-lookup"><span data-stu-id="b71e6-p102">When the installation process finishes, you will have the latest Office applications installed. To verify that you have the latest build, go to **File** > **Account** from any Office application. Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Uma captura de tela que mostra informações do produto com o rótulo Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="b71e6-125">Builds mínimos do Office para conjuntos de requisitos de API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="b71e6-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="b71e6-126">Para saber mais sobre os builds mínimos de produtos para cada plataforma dos conjuntos de requisitos de API, confira os seguintes:</span><span class="sxs-lookup"><span data-stu-id="b71e6-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="b71e6-127">Conjuntos de requisitos da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="b71e6-127">Word JavaScript API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js)
- [<span data-ttu-id="b71e6-128">Conjuntos de requisitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="b71e6-128">Excel JavaScript API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js)
- [<span data-ttu-id="b71e6-129">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="b71e6-129">OneNote JavaScript API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [<span data-ttu-id="b71e6-130">Conjuntos de requisitos da API de Caixa de Diálogo</span><span class="sxs-lookup"><span data-stu-id="b71e6-130">Dialog API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [<span data-ttu-id="b71e6-131">Conjuntos de requisitos comuns da API do Office</span><span class="sxs-lookup"><span data-stu-id="b71e6-131">Office common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
