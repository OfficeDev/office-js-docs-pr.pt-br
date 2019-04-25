---
title: Atualize para a API de JavaScript mais recente da biblioteca do Office e o esquema de manifesto do suplemento da versão 1.1
description: Atualize seus arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no seu projeto de Suplemento do Office para a versão 1.1.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7cbda821897b33a19e4bc9eeac27a096e01bc217
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448711"
---
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a><span data-ttu-id="24f71-103">Atualize para a API de JavaScript mais recente da biblioteca do Office e o esquema de manifesto do suplemento da versão 1.1</span><span class="sxs-lookup"><span data-stu-id="24f71-103">Update to the latest JavaScript API for Office library and version 1.1 add-in manifest schema</span></span>

<span data-ttu-id="24f71-104">Este artigo descreve como atualizar os arquivos do JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação do manifesto do suplemento no projeto do suplemento do Office para a versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="24f71-104">This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="24f71-105">Projetos criados no Visual Studio 2017 já usam a versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="24f71-105">Projects created in Visual Studio 2017 will already use version 1.1.</span></span> <span data-ttu-id="24f71-106">No entanto, há atualizações secundárias ocasionais para a versão 1.1 que você pode aplicar ao usar as técnicas neste artigo.</span><span class="sxs-lookup"><span data-stu-id="24f71-106">However there are occasional minor updates to version 1.1 that you can apply by using the techniques in this article.</span></span>

## <a name="use-the-most-up-to-date-project-files"></a><span data-ttu-id="24f71-107">Usar os arquivos de projeto mais atualizados</span><span class="sxs-lookup"><span data-stu-id="24f71-107">Use the most up-to-date project files</span></span>

<span data-ttu-id="24f71-108">Se estiver usando o Visual Studio para desenvolver o suplemento, para usar os [membros mais recentes](/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office) da API JavaScript para Office e os [recursos da v1.1 do manifesto do suplemento](../develop/add-in-manifests.md) (que é validado em relação a offappmanifest 1.1.xsd), é preciso baixar o Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="24f71-108">If you use Visual Studio to develop your add-in, to use the [newest API members](/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office) of the JavaScript API for Office and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download Visual Studio 2017.</span></span> <span data-ttu-id="24f71-109">Para baixar o Visual Studio 2017, confira a [página IDE do Visual Studio](https://visualstudio.microsoft.com/vs/).</span><span class="sxs-lookup"><span data-stu-id="24f71-109">To download Visual Studio 2017, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="24f71-110">Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.</span><span class="sxs-lookup"><span data-stu-id="24f71-110">During installation you'll need to select the Office/SharePoint development workload.</span></span>

<span data-ttu-id="24f71-111">Se estiver usando um editor de texto ou IDE que não o Visual Studio para desenvolver o suplemento, é precisa atualizar as referências à CDN para o Office.js e a versão do esquema consultada pelo manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="24f71-111">If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.</span></span>

<span data-ttu-id="24f71-112">Para executar um suplemento desenvolvido usando recursos novos e atualizados da API do Office.js e do suplemento do manifesto, seus clientes devem estar executando o Office 2013 SP1 ou uma versão posterior de produtos locais e, quando aplicável, o SharePoint Server 2013 SP1 e produtos de servidor relacionados, o Exchange Server 2013 Service Pack 1 (SP1) ou produtos hospedados online equivalentes: Office 365, SharePoint Online e Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="24f71-112">To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Office 365, SharePoint Online, and Exchange Online.</span></span>

<span data-ttu-id="24f71-113">Para baixar os produtos do Office, SharePoint e Exchange SP1, consulte o seguinte:</span><span class="sxs-lookup"><span data-stu-id="24f71-113">To download Office, SharePoint, and Exchange SP1 products, see the following:</span></span>

- [<span data-ttu-id="24f71-114">Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft Office 2013 e produtos da área de trabalho relacionados</span><span class="sxs-lookup"><span data-stu-id="24f71-114">List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products</span></span>](https://support.microsoft.com/kb/2850036)

- [<span data-ttu-id="24f71-115">Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft SharePoint Server 2013 e produtos do servidor relacionados</span><span class="sxs-lookup"><span data-stu-id="24f71-115">List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products</span></span>](https://support.microsoft.com/kb/2850035)

- [<span data-ttu-id="24f71-116">Descrição do Exchange Server 2013 Service Pack 1</span><span class="sxs-lookup"><span data-stu-id="24f71-116">Description of Exchange Server 2013 Service Pack 1</span></span>](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a><span data-ttu-id="24f71-117">Atualização de um projeto de suplemento do Office criado com o Visual Studio</span><span class="sxs-lookup"><span data-stu-id="24f71-117">Updating an Office Add-in project created with Visual Studio</span></span>

<span data-ttu-id="24f71-118">Para projetos criados antes do lançamento da v1.1 da API JavaScript para Office e o esquema de manifesto do suplemento, é possível atualizar os arquivos de um projeto usando o **NuGet Package Manager** e, em seguida, atualizar as páginas HTML do suplemento para fazer referência a eles.</span><span class="sxs-lookup"><span data-stu-id="24f71-118">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you can update a project's files using the  **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.</span></span> 

<span data-ttu-id="24f71-119">Observe que o processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="24f71-119">Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a><span data-ttu-id="24f71-120">Atualizar os arquivos da biblioteca da API JavaScript para Office em seu projeto para a versão mais recente</span><span class="sxs-lookup"><span data-stu-id="24f71-120">Update the JavaScript API for Office library files in your project to the newest release</span></span>
<span data-ttu-id="24f71-121">As etapas a seguir atualizam seus arquivos da biblioteca do Office para a versão mais recente.</span><span class="sxs-lookup"><span data-stu-id="24f71-121">The following steps will update your Office library files to the latest version.</span></span> <span data-ttu-id="24f71-122">As etapas usam o Visual Studio 2017, mas são semelhantes para o Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="24f71-122">The steps use Visual Studio 2017, but they are similar for Visual Studio 2015.</span></span>

1. <span data-ttu-id="24f71-123">No Visual Studio 2017, abra ou crie um novo projeto de **Suplemento do Office**.</span><span class="sxs-lookup"><span data-stu-id="24f71-123">In Visual Studio 2017, open or create a new  **Office Add-in** project.</span></span>
2. <span data-ttu-id="24f71-124">Escolha **Ferramentas** > **Gerenciador de Pacotes NuGet** > **Gerenciar Pacotes Nuget para a Solução**.</span><span class="sxs-lookup"><span data-stu-id="24f71-124">Choose  **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.</span></span>
3. <span data-ttu-id="24f71-125">No **Gerenciador de Pacotes NuGet**, escolha **nuget.org** como **Origem do pacote**.</span><span class="sxs-lookup"><span data-stu-id="24f71-125">In the  **NuGet Package Manager**, select  **nuget.org** for **Package source**.</span></span>
4. <span data-ttu-id="24f71-126">Escolha a guia **Atualizações**.</span><span class="sxs-lookup"><span data-stu-id="24f71-126">Choose the **Updates** tab.</span></span>
5. <span data-ttu-id="24f71-127">Selecione Microsoft.Office.js.</span><span class="sxs-lookup"><span data-stu-id="24f71-127">Select Microsoft.Office.js.</span></span>
6. <span data-ttu-id="24f71-128">No painel à esquerda, escolha **Atualizar** e conclua o processo de atualização do pacote.</span><span class="sxs-lookup"><span data-stu-id="24f71-128">In the left pane, choose **Update** and complete the package update process.</span></span>

<span data-ttu-id="24f71-129">Você precisará realizar algumas etapas adicionais para concluir a atualização.</span><span class="sxs-lookup"><span data-stu-id="24f71-129">You'll need to take a few additional steps to complete the update.</span></span> <span data-ttu-id="24f71-130">Na marca **head** das páginas HTML do suplemento, comente ou exclua quaisquer referências existentes ao script office.js e faça referência à biblioteca atualizada da API JavaScript para Office da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="24f71-130">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > <span data-ttu-id="24f71-131">O `/1/` em `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.</span><span class="sxs-lookup"><span data-stu-id="24f71-131">The `/1/` in the `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="24f71-132">Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema</span><span class="sxs-lookup"><span data-stu-id="24f71-132">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="24f71-133">No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).</span><span class="sxs-lookup"><span data-stu-id="24f71-133">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="24f71-134">Após atualizar a versão do esquema do manifesto do suplemento para 1.1, será preciso remover os elementos **Capabilities** e **Capability** e substituí-los pelos [Hosts](/office/dev/add-ins/reference/manifest/hosts) e elementos [Host](/office/dev/add-ins/reference/manifest/host) ou pelos [elementos Requirements e Requirement](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="24f71-134">After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](/office/dev/add-ins/reference/manifest/hosts) and [Host](/office/dev/add-ins/reference/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a><span data-ttu-id="24f71-135">Atualização de um projeto de suplemento do Office criado com um editor de texto ou outro IDE</span><span class="sxs-lookup"><span data-stu-id="24f71-135">Updating an Office Add-in project created with a text editor or other IDE</span></span>

<span data-ttu-id="24f71-136">Para projetos criados antes do lançamento da v1.1 da API JavaScript para Office e o esquema de manifesto de suplemento, é preciso atualizar as páginas HTML do suplemento para fazerem referência à CDN da biblioteca v1.1 e atualizar o arquivo de manifesto do suplemento para usar a v1.1 do esquema.</span><span class="sxs-lookup"><span data-stu-id="24f71-136">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.</span></span> 

<span data-ttu-id="24f71-137">O processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="24f71-137">The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

<span data-ttu-id="24f71-138">Você não precisa de cópias locais dos arquivos da API JavaScript para Office (Office.js e arquivos .js específicos do aplicativo) para desenvolver um suplemento do Office (a referência à CDN para Office.js baixa os arquivos necessários no tempo de execução). Porém, se desejar uma cópia local dos arquivos da biblioteca, pode usar o [Utilitário de Linha de Comando NuGet](https://docs.nuget.org/consume/installing-nuget) e o comando `Install-Package Microsoft.Office.js` para baixá-los.</span><span class="sxs-lookup"><span data-stu-id="24f71-138">You don't need local copies of the JavaScript API for Office files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](https://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.</span></span>

> [!NOTE]
> <span data-ttu-id="24f71-139">Para obter uma cópia da XSD (Definição de esquema XML) para o manifesto do suplemento v1.1, confira a listagem em [Referência de esquema para manifestos de Suplementos do Office (v1.1)](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="24f71-139">To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a><span data-ttu-id="24f71-140">Atualizar os arquivos da biblioteca da API JavaScript para Office em seu projeto para usar a versão mais recente</span><span class="sxs-lookup"><span data-stu-id="24f71-140">Update the JavaScript API for Office library files in your project to use the newest release</span></span>

1. <span data-ttu-id="24f71-141">Abra as páginas HTML do suplemento no editor de texto ou IDE.</span><span class="sxs-lookup"><span data-stu-id="24f71-141">Open the HTML pages for your add-in in your text editor or IDE.</span></span>

2. <span data-ttu-id="24f71-142">Na marca **head** das páginas HTML do suplemento, comente ou exclua quaisquer referências existentes ao script office.js e faça referência à biblioteca atualizada da API JavaScript para Office da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="24f71-142">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > <span data-ttu-id="24f71-143">O `/1/` na frente de `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.</span><span class="sxs-lookup"><span data-stu-id="24f71-143">The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="24f71-144">Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema</span><span class="sxs-lookup"><span data-stu-id="24f71-144">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="24f71-145">No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).</span><span class="sxs-lookup"><span data-stu-id="24f71-145">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="24f71-146">Após atualizar a versão do esquema do manifesto do suplemento para 1.1, será preciso remover os elementos **Capabilities** e **Capability** e substituí-los pelos [Hosts](/office/dev/add-ins/reference/manifest/hosts) e elementos [Host](/office/dev/add-ins/reference/manifest/host) ou pelos [elementos Requirements e Requirement](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="24f71-146">After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](/office/dev/add-ins/reference/manifest/hosts) and [Host](/office/dev/add-ins/reference/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="24f71-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="24f71-147">See also</span></span>

- <span data-ttu-id="24f71-148">[Especificar hosts do Office e requisitos de API](specify-office-hosts-and-api-requirements.md) ]</span><span class="sxs-lookup"><span data-stu-id="24f71-148">[Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md) ]</span></span>
- [<span data-ttu-id="24f71-149">Noções básicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="24f71-149">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="24f71-150">API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="24f71-150">JavaScript API for Office</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="24f71-151">Referência de esquema para manifestos de suplementos do Office (versão 1.1)</span><span class="sxs-lookup"><span data-stu-id="24f71-151">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
