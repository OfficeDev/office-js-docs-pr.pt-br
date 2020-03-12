---
title: Elemento ExtendedPermissions no arquivo de manifesto
description: ''
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 966378b8bbed66960d7a99c4a82df75ace1c9161
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605797"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="382e4-102">Elemento ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="382e4-102">ExtendedPermissions element</span></span>

<span data-ttu-id="382e4-103">Define o conjunto de permissões estendidas que o suplemento precisa para acessar as APIs ou recursos associados.</span><span class="sxs-lookup"><span data-stu-id="382e4-103">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="382e4-104">O `ExtendedPermissions` elemento é um elemento filho de [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="382e4-104">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="382e4-105">Esse elemento só está disponível no [conjunto de requisitos de visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) de suplementos do Outlook em relação ao Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="382e4-105">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="382e4-106">Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="382e4-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="382e4-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="382e4-107">Child elements</span></span>

|  <span data-ttu-id="382e4-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="382e4-108">Element</span></span> |  <span data-ttu-id="382e4-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="382e4-109">Required</span></span>  |  <span data-ttu-id="382e4-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="382e4-110">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="382e4-111">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="382e4-111">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="382e4-112">Não</span><span class="sxs-lookup"><span data-stu-id="382e4-112">No</span></span>   | <span data-ttu-id="382e4-113">Define uma permissão estendida necessária para que o suplemento acesse a API ou o recurso associado.</span><span class="sxs-lookup"><span data-stu-id="382e4-113">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="382e4-114">`ExtendedPermissions`como</span><span class="sxs-lookup"><span data-stu-id="382e4-114">`ExtendedPermissions` example</span></span>

<span data-ttu-id="382e4-115">Veja a seguir um exemplo do `ExtendedPermissions` elemento.</span><span class="sxs-lookup"><span data-stu-id="382e4-115">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="382e4-116">Contido em</span><span class="sxs-lookup"><span data-stu-id="382e4-116">Contained in</span></span>

[<span data-ttu-id="382e4-117">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="382e4-117">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="382e4-118">Pode conter</span><span class="sxs-lookup"><span data-stu-id="382e4-118">Can contain</span></span>

[<span data-ttu-id="382e4-119">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="382e4-119">ExtendedPermission</span></span>](extendedpermission.md)
