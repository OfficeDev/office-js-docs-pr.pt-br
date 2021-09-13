---
title: Elemento ExtendedPermissions no arquivo de manifesto
description: Define a coleção de permissões estendidas que o add-in precisa para acessar APIs ou recursos associados.
ms.date: 10/15/2020
ms.localizationpriority: medium
ms.openlocfilehash: 633609e43b9de656b5bc483fc59a5b4c24556254
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152083"
---
# <a name="extendedpermissions-element"></a>Elemento ExtendedPermissions

Define a coleção de permissões estendidas que o add-in precisa para acessar APIs ou recursos associados. O `ExtendedPermissions` elemento é um elemento filho de [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> O suporte a esse elemento foi introduzido no conjunto de requisitos 1.9. Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Não   | Define uma permissão estendida necessária para que o add-in acesse a API ou recurso associado. |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` exemplo

A seguir, um exemplo do `ExtendedPermissions` elemento.

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
