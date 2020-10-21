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
# <a name="extendedpermissions-element"></a>Elemento ExtendedPermissions

Define o conjunto de permissões estendidas que o suplemento precisa para acessar as APIs ou recursos associados. O `ExtendedPermissions` elemento é um elemento filho de [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> O suporte para este elemento foi introduzido no conjunto de requisitos 1,9. Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Não   | Define uma permissão estendida necessária para que o suplemento acesse a API ou o recurso associado. |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` como

Veja a seguir um exemplo do `ExtendedPermissions` elemento.

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

[VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>Pode conter

[ExtendedPermission](extendedpermission.md)
