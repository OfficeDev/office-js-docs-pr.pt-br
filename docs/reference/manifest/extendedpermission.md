---
title: Elemento ExtendedPermission no arquivo de manifesto
description: Define uma permissão estendida que o suplemento precisa para acessar a API ou o recurso associado.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 7ff17312ae487d20f4d7af0ed4405cedd8820253
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720599"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="f7bd1-103">`ExtendedPermission`pseudoelemento</span><span class="sxs-lookup"><span data-stu-id="f7bd1-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="f7bd1-104">Define uma permissão estendida que o suplemento precisa para acessar a API ou o recurso associado.</span><span class="sxs-lookup"><span data-stu-id="f7bd1-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="f7bd1-105">O `ExtendedPermission` elemento é um elemento filho de [ExtendedPermissions](extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="f7bd1-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f7bd1-106">Esse elemento só está disponível no [conjunto de requisitos de visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) de suplementos do Outlook em relação ao Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="f7bd1-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="f7bd1-107">Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="f7bd1-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="f7bd1-108">Permissões estendidas disponíveis</span><span class="sxs-lookup"><span data-stu-id="f7bd1-108">Available extended permissions</span></span>

<span data-ttu-id="f7bd1-109">Estes são os valores disponíveis.</span><span class="sxs-lookup"><span data-stu-id="f7bd1-109">The following are the available values.</span></span>

|<span data-ttu-id="f7bd1-110">Valor disponível</span><span class="sxs-lookup"><span data-stu-id="f7bd1-110">Available value</span></span>|<span data-ttu-id="f7bd1-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="f7bd1-111">Description</span></span>|<span data-ttu-id="f7bd1-112">Hosts</span><span class="sxs-lookup"><span data-stu-id="f7bd1-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="f7bd1-113">Declara que o suplemento está usando a API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="f7bd1-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="f7bd1-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="f7bd1-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="f7bd1-115">`ExtendedPermission`como</span><span class="sxs-lookup"><span data-stu-id="f7bd1-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="f7bd1-116">Veja a seguir um exemplo do `ExtendedPermission` elemento.</span><span class="sxs-lookup"><span data-stu-id="f7bd1-116">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="f7bd1-117">Contido em</span><span class="sxs-lookup"><span data-stu-id="f7bd1-117">Contained in</span></span>

[<span data-ttu-id="f7bd1-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="f7bd1-118">ExtendedPermissions</span></span>](extendedpermissions.md)
