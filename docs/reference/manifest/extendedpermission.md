---
title: Elemento ExtendedPermission no arquivo de manifesto
description: Define uma permissão estendida que o complemento precisa para acessar a API ou recurso associado.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 29e4a16c09982a2f7b52fd085960b33dba1efb8781d8a33b63bf138a571ef00e
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57090177"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` elemento

Define uma permissão estendida que o complemento precisa para acessar a API ou recurso associado. O `ExtendedPermission` elemento é um elemento filho de [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> O suporte a esse elemento foi introduzido no conjunto de requisitos 1.9. Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="available-extended-permissions"></a>Permissões estendidas disponíveis

A seguir estão os valores disponíveis.

|Valor disponível|Descrição|Hosts|
|---|---|---|
|`AppendOnSend`|Declara que o complemento está usando o [Office. API Body.appendOnSendAsync.](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendOnSendAsync_data__options__callback_)|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` exemplo

A seguir, um exemplo do `ExtendedPermission` elemento.

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
