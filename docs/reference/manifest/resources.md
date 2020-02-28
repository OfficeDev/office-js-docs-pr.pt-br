---
title: Elemento Resources no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7e1cd9fcb383fa4d5881917b3dd8d6dec3bbe4f8
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324824"
---
# <a name="resources-element"></a><span data-ttu-id="5adb1-102">Elemento Resources</span><span class="sxs-lookup"><span data-stu-id="5adb1-102">Resources element</span></span>

<span data-ttu-id="5adb1-p101">Contém ícones, cadeias de caracteres e URLs para o nó [VersionOverrides](versionoverrides.md). Um elemento de manifesto especifica um recurso usando a **d** do recurso. Isso ajuda a manter o tamanho do manifesto manejável, especialmente quando os recursos tiverem versões para localidades diferentes. Uma **id** deve ser exclusiva dentro do manifesto e pode ter no máximo 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="5adb1-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="5adb1-107">Cada recurso pode ter um ou mais elementos filhos **Override** para definir um recurso diferente para uma localidade específica.</span><span class="sxs-lookup"><span data-stu-id="5adb1-107">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5adb1-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="5adb1-108">Child elements</span></span>

|  <span data-ttu-id="5adb1-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="5adb1-109">Element</span></span> |  <span data-ttu-id="5adb1-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="5adb1-110">Type</span></span>  |  <span data-ttu-id="5adb1-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="5adb1-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5adb1-112">Imagens</span><span class="sxs-lookup"><span data-stu-id="5adb1-112">Images</span></span>](#images)            |  <span data-ttu-id="5adb1-113">image</span><span class="sxs-lookup"><span data-stu-id="5adb1-113">image</span></span>   |  <span data-ttu-id="5adb1-114">Fornece a URL HTTPS de uma imagem para um ícone.</span><span class="sxs-lookup"><span data-stu-id="5adb1-114">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="5adb1-115">**URLs**</span><span class="sxs-lookup"><span data-stu-id="5adb1-115">**Urls**</span></span>                |  <span data-ttu-id="5adb1-116">url</span><span class="sxs-lookup"><span data-stu-id="5adb1-116">url</span></span>     |  <span data-ttu-id="5adb1-p102">Fornece um local para a URL HTTPS. A URL pode ter 2.048 caracteres no máximo.</span><span class="sxs-lookup"><span data-stu-id="5adb1-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="5adb1-119">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="5adb1-119">**ShortStrings**</span></span> |  <span data-ttu-id="5adb1-120">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5adb1-120">string</span></span>  |  <span data-ttu-id="5adb1-p103">O texto para os elementos **Label** e **Title**. Cada **String** contém no máximo 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="5adb1-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="5adb1-123">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="5adb1-123">**LongStrings**</span></span>  |  <span data-ttu-id="5adb1-124">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5adb1-124">string</span></span>  | <span data-ttu-id="5adb1-p104">O texto para atributos **Description**. Cada **String** contém no máximo 250 caracteres.</span><span class="sxs-lookup"><span data-stu-id="5adb1-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="5adb1-127">Use o protocolo SSL (Secure Sockets Layer) para todas as URLs nos elementos **Image** e **Url**.</span><span class="sxs-lookup"><span data-stu-id="5adb1-127">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="5adb1-128">Imagens</span><span class="sxs-lookup"><span data-stu-id="5adb1-128">Images</span></span>
<span data-ttu-id="5adb1-129">Cada ícone deve ter três elementos **images** , um para cada um dos três tamanhos obrigatórios:</span><span class="sxs-lookup"><span data-stu-id="5adb1-129">Each icon must have three **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="5adb1-130">16 x 16</span><span class="sxs-lookup"><span data-stu-id="5adb1-130">16x16</span></span>
- <span data-ttu-id="5adb1-131">32x32</span><span class="sxs-lookup"><span data-stu-id="5adb1-131">32x32</span></span>
- <span data-ttu-id="5adb1-132">80x80</span><span class="sxs-lookup"><span data-stu-id="5adb1-132">80x80</span></span>

<span data-ttu-id="5adb1-133">Os seguintes tamanhos adicionais também têm suporte, mas não são obrigatórios:</span><span class="sxs-lookup"><span data-stu-id="5adb1-133">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="5adb1-134">20x20</span><span class="sxs-lookup"><span data-stu-id="5adb1-134">20x20</span></span>
- <span data-ttu-id="5adb1-135">24x24</span><span class="sxs-lookup"><span data-stu-id="5adb1-135">24x24</span></span>
- <span data-ttu-id="5adb1-136">40x40</span><span class="sxs-lookup"><span data-stu-id="5adb1-136">40x40</span></span>
- <span data-ttu-id="5adb1-137">48x48</span><span class="sxs-lookup"><span data-stu-id="5adb1-137">48x48</span></span>
- <span data-ttu-id="5adb1-138">64x64</span><span class="sxs-lookup"><span data-stu-id="5adb1-138">64x64</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="5adb1-139">O Outlook requer a capacidade de armazenar em cache os recursos de imagem para fins de desempenho.</span><span class="sxs-lookup"><span data-stu-id="5adb1-139">Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="5adb1-140">Por esse motivo, o servidor que hospeda um recurso de imagem não deve adicionar nenhuma diretriz CACHE-CONTROL ao cabeçalho da resposta.</span><span class="sxs-lookup"><span data-stu-id="5adb1-140">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="5adb1-141">Isso fará com que o Outlook substitua automaticamente uma imagem padrão ou genérica.</span><span class="sxs-lookup"><span data-stu-id="5adb1-141">This will result in Outlook automatically substituting a generic or default image.</span></span>    

## <a name="resources-examples"></a><span data-ttu-id="5adb1-142">Exemplos de recursos</span><span class="sxs-lookup"><span data-stu-id="5adb1-142">Resources examples</span></span> 

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
