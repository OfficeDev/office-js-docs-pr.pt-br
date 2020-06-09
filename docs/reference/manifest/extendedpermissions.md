---
title: Elemento ExtendedPermissions no arquivo de manifesto
description: Define o conjunto de permissões estendidas que o suplemento precisa para acessar as APIs ou recursos associados.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: cf59d13d794f8f303da6cc0ca39066584bc3f56c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611530"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="44f1c-103">Elemento ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="44f1c-103">ExtendedPermissions element</span></span>

<span data-ttu-id="44f1c-104">Define o conjunto de permissões estendidas que o suplemento precisa para acessar as APIs ou recursos associados.</span><span class="sxs-lookup"><span data-stu-id="44f1c-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="44f1c-105">O `ExtendedPermissions` elemento é um elemento filho de [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="44f1c-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="44f1c-106">Esse elemento só está disponível no [conjunto de requisitos de visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) de suplementos do Outlook em relação ao Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="44f1c-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="44f1c-107">Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="44f1c-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="44f1c-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="44f1c-108">Child elements</span></span>

|  <span data-ttu-id="44f1c-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="44f1c-109">Element</span></span> |  <span data-ttu-id="44f1c-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="44f1c-110">Required</span></span>  |  <span data-ttu-id="44f1c-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="44f1c-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="44f1c-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="44f1c-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="44f1c-113">Não</span><span class="sxs-lookup"><span data-stu-id="44f1c-113">No</span></span>   | <span data-ttu-id="44f1c-114">Define uma permissão estendida necessária para que o suplemento acesse a API ou o recurso associado.</span><span class="sxs-lookup"><span data-stu-id="44f1c-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="44f1c-115">`ExtendedPermissions`como</span><span class="sxs-lookup"><span data-stu-id="44f1c-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="44f1c-116">Veja a seguir um exemplo do `ExtendedPermissions` elemento.</span><span class="sxs-lookup"><span data-stu-id="44f1c-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="44f1c-117">Contido em</span><span class="sxs-lookup"><span data-stu-id="44f1c-117">Contained in</span></span>

[<span data-ttu-id="44f1c-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="44f1c-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="44f1c-119">Pode conter</span><span class="sxs-lookup"><span data-stu-id="44f1c-119">Can contain</span></span>

[<span data-ttu-id="44f1c-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="44f1c-120">ExtendedPermission</span></span>](extendedpermission.md)
