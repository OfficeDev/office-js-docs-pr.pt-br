---
title: Elemento ExtendedPermissions no arquivo de manifesto
description: Define a coleção de permissões estendidas que o add-in precisa para acessar APIs ou recursos associados.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 46ca6e3e2fb992755d9067b4251200073f07ade1
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042123"
---
# <a name="extendedpermissions-element"></a>Elemento ExtendedPermissions

Define a coleção de permissões estendidas que o add-in precisa para acessar APIs ou recursos associados. O `ExtendedPermissions` elemento é um elemento filho de [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> O suporte a esse elemento foi introduzido no conjunto de requisitos 1.9. Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

**Tipo de suplemento:** Email

**Válido somente nestes esquemas VersionOverrides:**

- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos:**

- [Caixa de correio 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

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
