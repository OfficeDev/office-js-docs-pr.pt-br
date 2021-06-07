---
title: Habilitar cenários de acesso de representante em um Outlook de entrada
description: Descreve brevemente o acesso de representantes e discute como configurar o suporte ao complemento.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 256c37087b10eaf9c8025e19a4990852f9550458
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783488"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Habilitar cenários de acesso de representante em um Outlook de entrada

Um proprietário de caixa de correio pode usar o recurso de acesso de representante [para permitir que outra pessoa gerencie seus emails e calendários.](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926) Este artigo especifica quais permissões delegar Office API JavaScript suporta e descreve como habilitar cenários de acesso de representante em seu Outlook de usuário.

> [!IMPORTANT]
> O acesso de representante não está disponível no Outlook android e iOS. Além disso, esse recurso não está disponível no momento com caixas de correio [compartilhadas](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) em grupo Outlook na Web. Essa funcionalidade pode ser disponibilizada no futuro.
>
> O suporte para esse recurso foi introduzido no conjunto de requisitos 1.8. Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="supported-permissions-for-delegate-access"></a>Permissões com suporte para acesso de representante

A tabela a seguir descreve as permissões de representante suportadas Office API JavaScript.

|Permissão|Valor|Descrição|
|---|---:|---|
|Leitura|1 (000001)|Pode ler itens.|
|Gravar|2 (000010)|Pode criar itens.|
|DeleteOwn|4 (000100)|Pode excluir apenas os itens criados.|
|DeleteAll|8 (001000)|Pode excluir qualquer item.|
|EditOwn|16 (010000)|Pode editar apenas os itens criados.|
|EditAll|32 (100000)|Pode editar todos os itens.|

> [!NOTE]
> Atualmente, a API oferece suporte para obter permissões de representante existentes, mas não definir permissões de representante.

O [objeto DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma máscara de bits para indicar as permissões do representante. Cada posição na máscara de bits representa uma permissão específica e, se estiver definida como, o `1` representante terá a respectiva permissão. Por exemplo, se o segundo bit da direita for `1` , o representante terá permissão **Gravar.** Você pode ver um exemplo de como verificar se há uma permissão específica na seção [Executar](#perform-an-operation-as-delegate) uma operação como representante posteriormente neste artigo.

## <a name="sync-across-mailbox-clients"></a>Sincronizar entre clientes de caixa de correio

As atualizações de um representante para a caixa de correio do proprietário geralmente são sincronizadas entre caixas de correio imediatamente.

No entanto, se as operações REST ou Exchange Web Services (EWS) foram usadas para definir uma propriedade estendida em um item, essas alterações podem levar algumas horas para sincronizar. Em vez disso, recomendamos que você use o [objeto CustomProperties](/javascript/api/outlook/office.customproperties) e APIs relacionadas para evitar esse atraso. Para saber mais, consulte a seção [propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "Obter e definir metadados em um Outlook de complemento".

> [!IMPORTANT]
> Em um cenário de representante, você não pode usar o EWS com os tokens atualmente fornecidos pela API office.js.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para habilitar cenários de acesso de representante no seu add-in, você deve definir o [elemento SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) como no `true` manifesto sob o elemento pai `DesktopFormFactor` . Atualmente, outros fatores de formulário não são suportados.

Para dar suporte a chamadas REST de um representante, de definir o nó [Permissões](../reference/manifest/permissions.md) no manifesto como `ReadWriteMailbox` .

O exemplo a seguir mostra `SupportsSharedFolders` o elemento definido como em uma seção do `true` manifesto.

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

Você pode obter as propriedades compartilhadas de um item no modo Redação ou Leitura chamando o [método item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) Isso retorna um [objeto SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do representante, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.

O exemplo a seguir mostra como obter as propriedades compartilhadas de uma mensagem ou compromisso, verificar se o representante tem permissão **Gravar** e fazer uma chamada REST.

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
> Como representante, você pode usar REST para obter o conteúdo de uma mensagem Outlook anexada a um item Outlook [ou postagem de grupo.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Manipular a chamada REST em itens compartilhados e não compartilhados

Se você quiser chamar uma operação REST em um item, se o item é compartilhado ou não, você pode usar a API para determinar se o `getSharedPropertiesAsync` item é compartilhado. Depois disso, você pode construir a URL REST para a operação usando o objeto apropriado.

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a>Limitações

Dependendo dos cenários do seu complemento, há algumas limitações a considerar ao lidar com situações de representante.

### <a name="rest-and-ews"></a>REST e EWS

Seu complemento pode usar REST, mas não EWS, e a permissão do add-in deve ser definida para habilitar o acesso REST à caixa de correio `ReadWriteMailbox` do proprietário.

### <a name="message-compose-mode"></a>Modo De composição de Mensagens

No modo Redação de Mensagem, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) não é suportado Outlook na Web ou Windows a menos que as seguintes condições sejam atendidas.

1. O proprietário compartilha pelo menos uma pasta de caixa de correio com o representante.
1. O representante esboça uma mensagem na pasta compartilhada.

    Exemplos:

    - O representante responde ou encaminha um email na pasta compartilhada.
    - O representante salva uma mensagem de rascunho e a move de sua própria pasta **Rascunhos** para a pasta compartilhada. O representante abre o rascunho da pasta compartilhada e continua compondo.

Depois que a mensagem é enviada, ela geralmente é encontrada na pasta Itens **Enviados do** representante.

## <a name="see-also"></a>Confira também

- [Permitir que outra pessoa gerencie seu email e calendário](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Compartilhamento de calendário em Microsoft 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Como solicitar elementos de manifesto](../develop/manifest-element-ordering.md)
- [Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Operadores de bit do JavaScript](https://www.w3schools.com/js/js_bitwise.asp)