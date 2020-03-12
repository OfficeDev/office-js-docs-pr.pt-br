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
# <a name="extendedpermissions-element"></a>Elemento ExtendedPermissions

Define o conjunto de permissões estendidas que o suplemento precisa para acessar as APIs ou recursos associados. O `ExtendedPermissions` elemento é um elemento filho de [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> Esse elemento só está disponível no [conjunto de requisitos de visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) de suplementos do Outlook em relação ao Exchange Online. Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Não   | Define uma permissão estendida necessária para que o suplemento acesse a API ou o recurso associado. |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions`como

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
