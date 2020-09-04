---
title: Habilitar cenários de acesso de representante em um suplemento do Outlook
description: Descreve brevemente o acesso de representante e discute como configurar o suporte a suplementos.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 965798995638c9ec518d1e69ed23763d845fb7ec
ms.sourcegitcommit: 604361e55dee45c7a5d34c2fa6937693c154fc24
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/03/2020
ms.locfileid: "47363825"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Habilitar cenários de acesso de representante em um suplemento do Outlook

Um proprietário de caixa de correio pode usar o recurso de acesso de representante para [permitir que outra pessoa gerencie seus emails e calendários](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926). Este artigo especifica a quais permissões de representante a API JavaScript do Office oferece suporte e descreve como habilitar cenários de acesso de representante no suplemento do Outlook.

> [!IMPORTANT]
> O acesso de representante não está disponível atualmente no Outlook no Android e iOS. Além disso, esse recurso não está disponível atualmente com [caixas de correio compartilhadas em grupo](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide#shared-mailboxes) no Outlook na Web. Essa funcionalidade pode ser disponibilizada no futuro.
>
> O suporte para esse recurso foi introduzido no conjunto de requisitos 1,8. Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="supported-permissions-for-delegate-access"></a>Permissões com suporte para acesso de representante

A tabela a seguir descreve as permissões de representante que a API JavaScript do Office suporta.

|Permissão|Valor|Descrição|
|---|---:|---|
|Ler|1 (000001)|Pode ler itens.|
|Gravar|2 (000010)|Pode criar itens.|
|DeleteOwn|4 (000100)|Só pode excluir os itens que eles criaram.|
|DeleteAll|8 (001000)|Pode excluir qualquer item.|
|EditOwn|16 (010000)|Só pode editar os itens que eles criaram.|
|EditAll|32 (100000)|Pode editar qualquer item.|

> [!NOTE]
> Atualmente, a API oferece suporte para obter permissões de representante existentes, mas não definir permissões de representante.

O objeto [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma bitmask para indicar as permissões do representante. Cada posição na bitmask representa uma permissão específica e, se estiver definida como `1` , o representante tem a respectiva permissão. Por exemplo, se o segundo bit à direita é `1` , o representante tem permissão de **gravação** . Você pode ver um exemplo de como verificar se há uma permissão específica na seção [executar uma operação como representante,](#perform-an-operation-as-delegate) posteriormente neste artigo.

## <a name="sync-across-mailbox-clients"></a>Sincronizar entre clientes de caixa de correio

As atualizações de um representante para a caixa de correio do proprietário costumam ser sincronizadas imediatamente nas caixas de correio.

No entanto, se as operações REST ou Exchange Web Services (EWS) foram usadas para definir uma propriedade estendida em um item, essas alterações poderão levar algumas horas para a sincronização. Em vez disso, recomendamos usar o objeto [CustomProperties](/javascript/api/outlook/office.customproperties) e APIs relacionadas para evitar esse atraso. Para saber mais, confira a [seção Propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "obter e definir metadados em um suplemento do Outlook".

> [!IMPORTANT]
> Em um cenário de representante, não é possível usar o EWS com os tokens atualmente fornecidos pela API office.js.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para habilitar cenários de acesso de representante no suplemento, você deve definir o elemento [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` no manifesto no elemento pai `DesktopFormFactor` . No momento, não há suporte para outros fatores de formulário.

Para oferecer suporte a chamadas REST de um representante, defina o nó de [permissões](../reference/manifest/permissions.md) no manifesto como `ReadWriteMailbox` .

O exemplo a seguir mostra o `SupportsSharedFolders` elemento definido como `true` em uma seção do manifesto.

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
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate"></a>Executar uma operação como representante

Você pode obter as propriedades compartilhadas de um item no modo de redação ou leitura chamando o método [Item. getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) . Isso retorna um objeto [SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do representante, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.

> [!IMPORTANT]
> Em um cenário de representante, o suplemento pode usar REST, mas não EWS, e a permissão do suplemento deve ser definida para `ReadWriteMailbox` habilitar o acesso REST à caixa de correio do proprietário.

O exemplo a seguir mostra como obter as propriedades compartilhadas de uma mensagem ou compromisso, verificar se o representante tem permissão de **gravação** e fazer uma chamada REST.

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> Como representante, você pode usar o REST para [obter o conteúdo de uma mensagem do Outlook anexada a um item do Outlook ou a uma postagem de grupo](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).

## <a name="see-also"></a>Confira também

- [Permitir que outra pessoa Gerencie seu email e calendário](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Compartilhamento de calendário no Office 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Como solicitar elementos de manifesto](../develop/manifest-element-ordering.md)
- [Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Operadores de JavaScript bit a bit](https://www.w3schools.com/js/js_bitwise.asp)