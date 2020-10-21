---
title: Elemento ExtendedPermissions no arquivo de manifesto
description: Define o conjunto de permissões estendidas que o suplemento precisa para acessar as APIs ou recursos associados.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626439"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="dc6bd-103">Elemento ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="dc6bd-103">ExtendedPermissions element</span></span>

<span data-ttu-id="dc6bd-104">Define o conjunto de permissões estendidas que o suplemento precisa para acessar as APIs ou recursos associados.</span><span class="sxs-lookup"><span data-stu-id="dc6bd-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="dc6bd-105">O `ExtendedPermissions` elemento é um elemento filho de [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="dc6bd-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dc6bd-106">O suporte para este elemento foi introduzido no conjunto de requisitos 1,9.</span><span class="sxs-lookup"><span data-stu-id="dc6bd-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="dc6bd-107">Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="dc6bd-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="child-elements"></a><span data-ttu-id="dc6bd-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="dc6bd-108">Child elements</span></span>

|  <span data-ttu-id="dc6bd-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="dc6bd-109">Element</span></span> |  <span data-ttu-id="dc6bd-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="dc6bd-110">Required</span></span>  |  <span data-ttu-id="dc6bd-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="dc6bd-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="dc6bd-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="dc6bd-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="dc6bd-113">Não</span><span class="sxs-lookup"><span data-stu-id="dc6bd-113">No</span></span>   | <span data-ttu-id="dc6bd-114">Define uma permissão estendida necessária para que o suplemento acesse a API ou o recurso associado.</span><span class="sxs-lookup"><span data-stu-id="dc6bd-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="dc6bd-115">`ExtendedPermissions` como</span><span class="sxs-lookup"><span data-stu-id="dc6bd-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="dc6bd-116">Veja a seguir um exemplo do `ExtendedPermissions` elemento.</span><span class="sxs-lookup"><span data-stu-id="dc6bd-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="dc6bd-117">Contido em</span><span class="sxs-lookup"><span data-stu-id="dc6bd-117">Contained in</span></span>

[<span data-ttu-id="dc6bd-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="dc6bd-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="dc6bd-119">Pode conter</span><span class="sxs-lookup"><span data-stu-id="dc6bd-119">Can contain</span></span>

[<span data-ttu-id="dc6bd-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="dc6bd-120">ExtendedPermission</span></span>](extendedpermission.md)
