---
title: Elemento ExtendedPermissions no arquivo de manifesto
description: Define a coleção de permissões estendidas que o add-in precisa para acessar APIs ou recursos associados.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: c3f021adfcc2f3a4ba7b7d7aeeb52f3213d92788d401130abbc92618930d09fe
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097890"
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
