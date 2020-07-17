---
title: Atualizar para a biblioteca de API JavaScript do Office mais recente e o esquema de manifesto de suplemento versão 1,1
description: Atualize seus arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no seu projeto de Suplemento do Office para a versão 1.1.
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: 34127b3920af1309d4e4c2e1c265c676640a1c24
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093550"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a><span data-ttu-id="41fec-103">Atualizar para a biblioteca de API JavaScript do Office mais recente e o esquema de manifesto de suplemento versão 1,1</span><span class="sxs-lookup"><span data-stu-id="41fec-103">Update to the latest Office JavaScript API library and version 1.1 add-in manifest schema</span></span>

<span data-ttu-id="41fec-104">Este artigo descreve como atualizar os arquivos do JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação do manifesto do suplemento no projeto do suplemento do Office para a versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="41fec-104">This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="41fec-105">Os projetos criados no Visual Studio 2019 já usarão a versão 1,1.</span><span class="sxs-lookup"><span data-stu-id="41fec-105">Projects created in Visual Studio 2019 will already use version 1.1.</span></span> <span data-ttu-id="41fec-106">No entanto, há atualizações secundárias ocasionais para a versão 1.1 que você pode aplicar ao usar as técnicas neste artigo.</span><span class="sxs-lookup"><span data-stu-id="41fec-106">However there are occasional minor updates to version 1.1 that you can apply by using the techniques in this article.</span></span>

## <a name="use-the-most-up-to-date-project-files"></a><span data-ttu-id="41fec-107">Usar os arquivos de projeto mais atualizados</span><span class="sxs-lookup"><span data-stu-id="41fec-107">Use the most up-to-date project files</span></span>

<span data-ttu-id="41fec-108">Se você usar o Visual Studio para desenvolver seu suplemento, para usar os membros mais recentes da API da API JavaScript do Office e os [recursos do v 1.1 do manifesto do suplemento](../develop/add-in-manifests.md) (que é validado no offappmanifest-1.1. xsd), será necessário baixar o Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="41fec-108">If you use Visual Studio to develop your add-in, to use the newest API members of the Office JavaScript API and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download Visual Studio 2019.</span></span> <span data-ttu-id="41fec-109">Para baixar o Visual Studio 2019, confira a [página IDE do Visual Studio](https://visualstudio.microsoft.com/vs/).</span><span class="sxs-lookup"><span data-stu-id="41fec-109">To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="41fec-110">Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.</span><span class="sxs-lookup"><span data-stu-id="41fec-110">During installation you'll need to select the Office/SharePoint development workload.</span></span>

<span data-ttu-id="41fec-111">Se estiver usando um editor de texto ou IDE que não o Visual Studio para desenvolver o suplemento, é precisa atualizar as referências à CDN para o Office.js e a versão do esquema consultada pelo manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="41fec-111">If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.</span></span>

<span data-ttu-id="41fec-112">Para executar um suplemento desenvolvido usando recursos novos e atualizados do Office.js API e suplementos de suplemento, seus clientes devem estar executando o Office 2013 SP1 ou versões posteriores, produtos locais, e, quando aplicável, SharePoint Server 2013 SP1 e produtos de servidor relacionados, Exchange Server 2013 Service Pack 1 (SP1) ou os produtos hospedados online equivalentes: Microsoft 365, SharePoint Online e Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="41fec-112">To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Microsoft 365, SharePoint Online, and Exchange Online.</span></span>

<span data-ttu-id="41fec-113">Para baixar os produtos do Office, SharePoint e Exchange SP1, consulte o seguinte:</span><span class="sxs-lookup"><span data-stu-id="41fec-113">To download Office, SharePoint, and Exchange SP1 products, see the following:</span></span>

- [<span data-ttu-id="41fec-114">Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft Office 2013 e produtos da área de trabalho relacionados</span><span class="sxs-lookup"><span data-stu-id="41fec-114">List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products</span></span>](https://support.microsoft.com/kb/2850036)

- [<span data-ttu-id="41fec-115">Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft SharePoint Server 2013 e produtos do servidor relacionados</span><span class="sxs-lookup"><span data-stu-id="41fec-115">List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products</span></span>](https://support.microsoft.com/kb/2850035)

- [<span data-ttu-id="41fec-116">Descrição do Exchange Server 2013 Service Pack 1</span><span class="sxs-lookup"><span data-stu-id="41fec-116">Description of Exchange Server 2013 Service Pack 1</span></span>](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a><span data-ttu-id="41fec-117">Atualização de um projeto de suplemento do Office criado com o Visual Studio</span><span class="sxs-lookup"><span data-stu-id="41fec-117">Updating an Office Add-in project created with Visual Studio</span></span>

<span data-ttu-id="41fec-118">Para projetos criados antes do lançamento da versão v 1.1 da API JavaScript do Office e do esquema de manifesto de suplemento, você pode atualizar os arquivos de um projeto usando o **Gerenciador de pacotes do NuGet**e, em seguida, atualizar as páginas HTML do suplemento para fazer referência a eles.</span><span class="sxs-lookup"><span data-stu-id="41fec-118">For projects created before the release of v1.1 of the Office JavaScript API and add-in manifest schema, you can update a project's files using the **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.</span></span> 

<span data-ttu-id="41fec-119">Observe que o processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="41fec-119">Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a><span data-ttu-id="41fec-120">Atualizar os arquivos da biblioteca da API JavaScript do Office em seu projeto para a versão mais recente</span><span class="sxs-lookup"><span data-stu-id="41fec-120">Update the Office JavaScript API library files in your project to the newest release</span></span>
<span data-ttu-id="41fec-121">As etapas a seguir atualizarão seus arquivos de biblioteca do Office.js para a versão mais recente.</span><span class="sxs-lookup"><span data-stu-id="41fec-121">The following steps will update your Office.js library files to the latest version.</span></span> <span data-ttu-id="41fec-122">As etapas usam o Visual Studio 2019, mas são semelhantes para versões anteriores do Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="41fec-122">The steps use Visual Studio 2019, but they are similar for previous versions of Visual Studio.</span></span>

1. <span data-ttu-id="41fec-123">No Visual Studio 2019, abra ou crie um novo projeto de **suplemento do Office** .</span><span class="sxs-lookup"><span data-stu-id="41fec-123">In Visual Studio 2019, open or create a new **Office Add-in** project.</span></span>
2. <span data-ttu-id="41fec-124">Escolha **ferramentas**  >  **NuGet Package Manager**  >  **gerenciar pacotes NuGet para solução**.</span><span class="sxs-lookup"><span data-stu-id="41fec-124">Choose **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.</span></span>
3. <span data-ttu-id="41fec-125">Escolha a guia **Atualizações**.</span><span class="sxs-lookup"><span data-stu-id="41fec-125">Choose the **Updates** tab.</span></span>
4. <span data-ttu-id="41fec-126">Selecione Microsoft.Office.js.</span><span class="sxs-lookup"><span data-stu-id="41fec-126">Select Microsoft.Office.js.</span></span> <span data-ttu-id="41fec-127">Verifique se a origem do pacote é de **NuGet.org**.</span><span class="sxs-lookup"><span data-stu-id="41fec-127">Ensure the package source is from **nuget.org**.</span></span>
5. <span data-ttu-id="41fec-128">No painel esquerdo, escolha **instalar** e concluir o processo de atualização do pacote.</span><span class="sxs-lookup"><span data-stu-id="41fec-128">In the left pane, choose **Install** and complete the package update process.</span></span>

<span data-ttu-id="41fec-129">Você precisará realizar algumas etapas adicionais para concluir a atualização.</span><span class="sxs-lookup"><span data-stu-id="41fec-129">You'll need to take a few additional steps to complete the update.</span></span> <span data-ttu-id="41fec-130">Na marca **Head** das páginas HTML do seu suplemento, comente ou exclua quaisquer referências de script office.js existentes e faça referência à biblioteca de API JavaScript do Office atualizada da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="41fec-130">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API library as follows:</span></span>

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > <span data-ttu-id="41fec-131">O `/1/` em `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.</span><span class="sxs-lookup"><span data-stu-id="41fec-131">The `/1/` in the `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="41fec-132">Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema</span><span class="sxs-lookup"><span data-stu-id="41fec-132">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="41fec-133">No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).</span><span class="sxs-lookup"><span data-stu-id="41fec-133">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="41fec-134">Após a atualização da versão do esquema de manifesto do suplemento para 1,1, você precisará remover os **recursos** e os elementos de **capacidade** e substituí-los pelos elementos [hosts](../reference/manifest/hosts.md) e [host](../reference/manifest/host.md) ou nos [elementos requirements](specify-office-hosts-and-api-requirements.md)e requirement.</span><span class="sxs-lookup"><span data-stu-id="41fec-134">After updating the version of the add-in manifest schema to 1.1, you will need to remove the **Capabilities** and **Capability** elements, and replace them with either the [Hosts](../reference/manifest/hosts.md) and [Host](../reference/manifest/host.md) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a><span data-ttu-id="41fec-135">Atualização de um projeto de suplemento do Office criado com um editor de texto ou outro IDE</span><span class="sxs-lookup"><span data-stu-id="41fec-135">Updating an Office Add-in project created with a text editor or other IDE</span></span>

<span data-ttu-id="41fec-136">Para projetos criados antes da versão do v 1.1 da API JavaScript do Office e do esquema de manifesto de suplemento, você precisa atualizar suas páginas HTML do suplemento para fazer referência à CDN da biblioteca v 1.1 e atualizar o arquivo de manifesto do suplemento para usar o esquema v 1.1.</span><span class="sxs-lookup"><span data-stu-id="41fec-136">For projects created before the release of v1.1 of the Office JavaScript API and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.</span></span> 

<span data-ttu-id="41fec-137">O processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="41fec-137">The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

<span data-ttu-id="41fec-138">Você não precisa de cópias locais dos arquivos da API JavaScript do Office (Office.js e arquivos. js específicos do aplicativo) para desenvolver o suplemento do Office (fazer referência à CDN para Office.js baixa os arquivos necessários no tempo de execução), mas se você quiser uma cópia local dos arquivos da biblioteca, poderá usar o [Utilitário de linha de comando do NuGet](https://docs.nuget.org/consume/installing-nuget) e o `Install-Package Microsoft.Office.js` comando para baixá-los.</span><span class="sxs-lookup"><span data-stu-id="41fec-138">You don't need local copies of the Office JavaScript API files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](https://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.</span></span>

> [!NOTE]
> <span data-ttu-id="41fec-139">Para obter uma cópia da XSD (Definição de esquema XML) para o manifesto do suplemento v1.1, confira a listagem em [Referência de esquema para manifestos de Suplementos do Office (v1.1)](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="41fec-139">To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>


### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a><span data-ttu-id="41fec-140">Atualizar os arquivos da biblioteca da API JavaScript do Office em seu projeto para usar a versão mais recente</span><span class="sxs-lookup"><span data-stu-id="41fec-140">Update the Office JavaScript API library files in your project to use the newest release</span></span>

1. <span data-ttu-id="41fec-141">Abra as páginas HTML do suplemento no editor de texto ou IDE.</span><span class="sxs-lookup"><span data-stu-id="41fec-141">Open the HTML pages for your add-in in your text editor or IDE.</span></span>

2. <span data-ttu-id="41fec-142">Na marca **Head** das páginas HTML do seu suplemento, comente ou exclua quaisquer referências de script office.js existentes e faça referência à biblioteca de API JavaScript do Office atualizada da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="41fec-142">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API library as follows:</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > <span data-ttu-id="41fec-143">O `/1/` na frente de `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.</span><span class="sxs-lookup"><span data-stu-id="41fec-143">The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="41fec-144">Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema</span><span class="sxs-lookup"><span data-stu-id="41fec-144">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="41fec-145">No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).</span><span class="sxs-lookup"><span data-stu-id="41fec-145">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="41fec-146">Após a atualização da versão do esquema de manifesto do suplemento para 1,1, você precisará remover os **recursos** e os elementos de **capacidade** e substituí-los pelos elementos [hosts](../reference/manifest/hosts.md) e [host](../reference/manifest/host.md) ou nos [elementos requirements](specify-office-hosts-and-api-requirements.md)e requirement.</span><span class="sxs-lookup"><span data-stu-id="41fec-146">After updating the version of the add-in manifest schema to 1.1, you will need to remove the **Capabilities** and **Capability** elements, and replace them with either the [Hosts](../reference/manifest/hosts.md) and [Host](../reference/manifest/host.md) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="41fec-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="41fec-147">See also</span></span>

- <span data-ttu-id="41fec-148">[Especificar hosts do Office e requisitos de API](specify-office-hosts-and-api-requirements.md) ]</span><span class="sxs-lookup"><span data-stu-id="41fec-148">[Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md) ]</span></span>
- [<span data-ttu-id="41fec-149">Entendendo a API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="41fec-149">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="41fec-150">API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="41fec-150">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="41fec-151">Referência de esquema para manifestos de suplementos do Office (versão 1.1)</span><span class="sxs-lookup"><span data-stu-id="41fec-151">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
