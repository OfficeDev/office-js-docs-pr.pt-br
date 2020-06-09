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
