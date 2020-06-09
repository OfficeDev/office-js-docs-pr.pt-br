---
title: Elemento ExtendedPermission no arquivo de manifesto
description: Define uma permissão estendida que o suplemento precisa para acessar a API ou o recurso associado.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: ca4c2d12cb9a5be159c22712b631d0bde42e48ed
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611537"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="f82f2-103">`ExtendedPermission`pseudoelemento</span><span class="sxs-lookup"><span data-stu-id="f82f2-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="f82f2-104">Define uma permissão estendida que o suplemento precisa para acessar a API ou o recurso associado.</span><span class="sxs-lookup"><span data-stu-id="f82f2-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="f82f2-105">O `ExtendedPermission` elemento é um elemento filho de [ExtendedPermissions](extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="f82f2-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f82f2-106">Esse elemento só está disponível no [conjunto de requisitos de visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) de suplementos do Outlook em relação ao Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="f82f2-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="f82f2-107">Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="f82f2-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="f82f2-108">Permissões estendidas disponíveis</span><span class="sxs-lookup"><span data-stu-id="f82f2-108">Available extended permissions</span></span>

<span data-ttu-id="f82f2-109">Estes são os valores disponíveis.</span><span class="sxs-lookup"><span data-stu-id="f82f2-109">The following are the available values.</span></span>

|<span data-ttu-id="f82f2-110">Valor disponível</span><span class="sxs-lookup"><span data-stu-id="f82f2-110">Available value</span></span>|<span data-ttu-id="f82f2-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="f82f2-111">Description</span></span>|<span data-ttu-id="f82f2-112">Hosts</span><span class="sxs-lookup"><span data-stu-id="f82f2-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="f82f2-113">Declara que o suplemento está usando a API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="f82f2-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="f82f2-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="f82f2-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="f82f2-115">`ExtendedPermission`como</span><span class="sxs-lookup"><span data-stu-id="f82f2-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="f82f2-116">Veja a seguir um exemplo do `ExtendedPermission` elemento.</span><span class="sxs-lookup"><span data-stu-id="f82f2-116">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="f82f2-117">Contido em</span><span class="sxs-lookup"><span data-stu-id="f82f2-117">Contained in</span></span>

[<span data-ttu-id="f82f2-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="f82f2-118">ExtendedPermissions</span></span>](extendedpermissions.md)
