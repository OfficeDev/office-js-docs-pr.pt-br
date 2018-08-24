---
title: Instalar a última versão do Office 2016
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 98dc69a7971a94b96bc3f7304fc7905f31013a87
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925231"
---
# <a name="install-the-latest-version-of-office-2016"></a><span data-ttu-id="4f008-102">Instalar a última versão do Office 2016</span><span class="sxs-lookup"><span data-stu-id="4f008-102">Install the latest version of Office 2016</span></span>

<span data-ttu-id="4f008-103">Novos recursos de desenvolvedor, inclusive os que ainda estão na visualização, são fornecidos primeiro aos assinantes que aceitam obter as últimas compilações do Office.</span><span class="sxs-lookup"><span data-stu-id="4f008-103">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="4f008-104">Aceitar para receber as versões mais recentes</span><span class="sxs-lookup"><span data-stu-id="4f008-104">Opt in to getting the latest builds</span></span>

<span data-ttu-id="4f008-105">Aceitar para receber as versões mais recentes do Office 2016:</span><span class="sxs-lookup"><span data-stu-id="4f008-105">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="4f008-106">Se você é assinante do Office 365 Home, Personal ou University, confira [Ser um Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="4f008-106">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="4f008-107">Se você for um cliente corporativo do Office 365, confira [Instalar a versão de Primeiro Lançamento do Office 365 para clientes corporativos](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="4f008-107">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="4f008-108">Se você estiver executando o Office 2016 em um Mac:</span><span class="sxs-lookup"><span data-stu-id="4f008-108">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="4f008-109">Inicie um programa do Office 2016 para Mac.</span><span class="sxs-lookup"><span data-stu-id="4f008-109">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="4f008-110">Selecione **Verificar Atualizações** no menu Ajuda.</span><span class="sxs-lookup"><span data-stu-id="4f008-110">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="4f008-111">Na caixa Microsoft AutoUpdate, marque a caixa para participar do programa Office Insider.</span><span class="sxs-lookup"><span data-stu-id="4f008-111">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="4f008-112">Para a versão mais recente:</span><span class="sxs-lookup"><span data-stu-id="4f008-112">Get the latest build</span></span>

<span data-ttu-id="4f008-113">Para receber as versões mais recentes do Office 2016:</span><span class="sxs-lookup"><span data-stu-id="4f008-113">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="4f008-114">Baixe a [Ferramenta de Implantação do Office 2016](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="4f008-114">Download the [Office 2016 Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span> 
2. <span data-ttu-id="4f008-p101">Execute a ferramenta. Isso extrai estes dois arquivos: Setup.exe e configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="4f008-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="4f008-117">Substitua o arquivo configuration.xml pelo [Arquivo de Configuração do Primeiro Lançamento](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="4f008-117">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="4f008-118">Execute o seguinte comando como administrador: `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="4f008-118">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="4f008-119">O comando pode demorar muito para ser executado sem indicar o progresso.</span><span class="sxs-lookup"><span data-stu-id="4f008-119">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="4f008-p102">Quando o processo de instalação for concluído, você terá os últimos aplicativos do Office 2016 instalados. Para verificar se você tem a última compilação, vá para **Arquivo**  >  **Conta** em qualquer aplicativo do Office. Em Atualizações do Office, você verá o rótulo (Office Insiders) acima do número de versão.</span><span class="sxs-lookup"><span data-stu-id="4f008-p102">When the installation process finishes, you will have the latest Office 2016 applications installed. To verify that you have the latest build, go to **File** > **Account** from any Office application. Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Uma captura de tela que mostra informações do produto com o rótulo Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="4f008-124">Builds mínimos do Office para conjuntos de requisitos de API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="4f008-124">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="4f008-125">Para saber mais sobre os builds mínimos de produtos para cada plataforma dos conjuntos de requisitos de API, confira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="4f008-125">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="4f008-126">Conjuntos de requisitos da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="4f008-126">Word JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets)
- [<span data-ttu-id="4f008-127">Conjuntos de requisitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="4f008-127">Excel JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)
- [<span data-ttu-id="4f008-128">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="4f008-128">OneNote JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets)
- [<span data-ttu-id="4f008-129">Conjuntos de requisitos da API de caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="4f008-129">Dialog API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="4f008-130">Conjuntos de requisitos de API comum do Office</span><span class="sxs-lookup"><span data-stu-id="4f008-130">Office common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
