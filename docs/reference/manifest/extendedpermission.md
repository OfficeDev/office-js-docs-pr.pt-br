---
title: Elemento ExtendedPermission no arquivo de manifesto
description: ''
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 6c41684fc922f5845559250311edd8182788cfc5
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605796"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission`pseudoelemento

Define uma permissão estendida que o suplemento precisa para acessar a API ou o recurso associado. O `ExtendedPermission` elemento é um elemento filho de [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> Esse elemento só está disponível no [conjunto de requisitos de visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) de suplementos do Outlook em relação ao Exchange Online. Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.

## <a name="available-extended-permissions"></a>Permissões estendidas disponíveis

Estes são os valores disponíveis.

|Valor disponível|Descrição|Hosts|
|---|---|---|
|`AppendOnSend`|Declara que o suplemento está usando a API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) .|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission`como

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
