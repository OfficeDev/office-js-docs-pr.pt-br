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
# <a name="extendedpermission-element"></a>`ExtendedPermission` pseudoelemento

Define uma permissão estendida que o suplemento precisa para acessar a API ou o recurso associado. O `ExtendedPermission` elemento é um elemento filho de [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> O suporte para este elemento foi introduzido no conjunto de requisitos 1,9. Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="available-extended-permissions"></a>Permissões estendidas disponíveis

Estes são os valores disponíveis.

|Valor disponível|Descrição|Hosts|
|---|---|---|
|`AppendOnSend`|Declara que o suplemento está usando a API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) .|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` como

Veja a seguir um exemplo do `ExtendedPermission` elemento.

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

## <a name="contained-in"></a>Contido em

[ExtendedPermissions](extendedpermissions.md)
