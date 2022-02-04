---
title: Elemento ExtendedPermission no arquivo de manifesto
description: Define uma permissão estendida que o complemento precisa para acessar a API ou recurso associado.
ms.date: 01/04/2022
ms.localizationpriority: medium
---

# <a name="extendedpermission-element"></a>`ExtendedPermission` elemento

Define uma permissão estendida que o complemento precisa para acessar a API ou recurso associado. O `ExtendedPermission` elemento é um elemento filho de [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> O suporte a esse elemento foi introduzido no conjunto de requisitos 1.9. Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

**Tipo de suplemento:** Email

**Válido somente nesses esquemas VersionOverrides**:

- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [Caixa de correio 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

## <a name="available-extended-permissions"></a>Permissões estendidas disponíveis

A seguir estão os valores disponíveis.

|Valor disponível|Descrição|Hosts|
|---|---|---|
|`AppendOnSend`|Declara que o complemento está usando o [Office. API Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-appendonsendasync-member(1)).|Outlook|

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
