---
title: Elemento ExtendedPermission no arquivo de manifesto
description: Define uma permissão estendida que o suplemento precisa para acessar a API ou o recurso associado.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 996cac59c44220d05165c7be6ae7c3d79d853271
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626397"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="ef343-103">`ExtendedPermission` pseudoelemento</span><span class="sxs-lookup"><span data-stu-id="ef343-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="ef343-104">Define uma permissão estendida que o suplemento precisa para acessar a API ou o recurso associado.</span><span class="sxs-lookup"><span data-stu-id="ef343-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="ef343-105">O `ExtendedPermission` elemento é um elemento filho de [ExtendedPermissions](extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="ef343-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ef343-106">O suporte para este elemento foi introduzido no conjunto de requisitos 1,9.</span><span class="sxs-lookup"><span data-stu-id="ef343-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="ef343-107">Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="ef343-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="ef343-108">Permissões estendidas disponíveis</span><span class="sxs-lookup"><span data-stu-id="ef343-108">Available extended permissions</span></span>

<span data-ttu-id="ef343-109">Estes são os valores disponíveis.</span><span class="sxs-lookup"><span data-stu-id="ef343-109">The following are the available values.</span></span>

|<span data-ttu-id="ef343-110">Valor disponível</span><span class="sxs-lookup"><span data-stu-id="ef343-110">Available value</span></span>|<span data-ttu-id="ef343-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="ef343-111">Description</span></span>|<span data-ttu-id="ef343-112">Hosts</span><span class="sxs-lookup"><span data-stu-id="ef343-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="ef343-113">Declara que o suplemento está usando a API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="ef343-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="ef343-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="ef343-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="ef343-115">`ExtendedPermission` como</span><span class="sxs-lookup"><span data-stu-id="ef343-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="ef343-116">Veja a seguir um exemplo do `ExtendedPermission` elemento.</span><span class="sxs-lookup"><span data-stu-id="ef343-116">The following is an example of the `ExtendedPermission` element.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a><span data-ttu-id="ef343-117">Contido em</span><span class="sxs-lookup"><span data-stu-id="ef343-117">Contained in</span></span>

[<span data-ttu-id="ef343-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="ef343-118">ExtendedPermissions</span></span>](extendedpermissions.md)
