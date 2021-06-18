---
title: Habilitar pastas compartilhadas e cenários de caixa de correio compartilhadas em um Outlook de entrada
description: Discute como configurar o suporte ao complemento para pastas compartilhadas (a.k.a. acesso delegado) e caixas de correio compartilhadas.
ms.date: 06/17/2021
localization_priority: Normal
ms.openlocfilehash: 5d7fb712b8f814184c2a444c32416d35fb1da49c
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007766"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>Habilitar pastas compartilhadas e cenários de caixa de correio compartilhadas em um Outlook de entrada

Este artigo descreve como habilitar pastas compartilhadas (também conhecidas como acesso de representante) e cenários de caixa de correio compartilhada (agora em visualização [)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)no seu Outlook add-in, incluindo quais permissões a API JavaScript Office suporta.

> [!IMPORTANT]
> O suporte a esse recurso foi introduzido no [conjunto de requisitos 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md). Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="supported-setups"></a>Configurações com suporte

As seções a seguir descrevem configurações com suporte para caixas de correio compartilhadas (agora em visualização) e pastas compartilhadas. As APIs de recurso podem não funcionar conforme o esperado em outras configurações. Selecione a plataforma que você gostaria de aprender a configurar.

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>Pastas compartilhadas

O proprietário da caixa de correio [deve primeiro fornecer acesso a um representante](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926). O representante deve seguir as instruções descritas na seção "Adicionar caixa de correio de outra pessoa ao seu perfil" do artigo Gerenciar itens de calendário e email de [outra pessoa.](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5)

#### <a name="shared-mailboxes-preview"></a>Caixas de correio compartilhadas (visualização)

Exchange administradores de servidor podem criar e gerenciar caixas de correio compartilhadas para conjuntos de usuários acessarem. No momento, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) é a única versão de servidor com suporte para esse recurso.

Um recurso Exchange Server conhecido como "automapping" está ativado por [](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) padrão, o que significa que, posteriormente, a caixa de correio compartilhada deve aparecer automaticamente no aplicativo Outlook do usuário depois que o Outlook tiver sido fechado e reaberto. No entanto, se um administrador tiver desabilitado a automação, o usuário deverá seguir as etapas manuais descritas na seção "Adicionar uma caixa de correio compartilhada ao Outlook" do artigo Abrir e usar uma caixa de correio compartilhada no [Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).

> [!WARNING]
> Não **entre** na caixa de correio compartilhada com uma senha. As APIs de recurso não funcionarão nesse caso.

### <a name="web-browser---modern-outlook"></a>[Navegador da Web – Outlook moderno](#tab/modern)

#### <a name="shared-folders"></a>Pastas compartilhadas

O proprietário da caixa de correio [deve primeiro fornecer acesso a um representante](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) atualizando as permissões de pasta de caixa de correio. O representante deve seguir as instruções descritas na seção "Adicionar caixa de correio de outra pessoa à sua lista de pastas Outlook Web App" do artigo Acessar a caixa de correio [de outra pessoa](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081).

#### <a name="shared-mailboxes-preview"></a>Caixas de correio compartilhadas (visualização)

Exchange administradores de servidor podem criar e gerenciar caixas de correio compartilhadas para conjuntos de usuários acessarem. No momento, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) é a única versão de servidor com suporte para esse recurso.

Depois de receber acesso, um usuário de caixa de correio compartilhada deve seguir as etapas descritas na seção "Adicionar a caixa de correio compartilhada para que ela seja exibida em sua caixa de correio principal" do artigo Abrir e usar uma caixa de correio compartilhada no [Outlook na Web](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).

> [!WARNING]
> NÃO **use** outras opções como "Abrir outra caixa de correio". As APIs de recurso podem não funcionar corretamente.

---

Para saber mais sobre onde os complementos fazem e não são ativados em geral, consulte a seção Itens de Caixa de Correio disponíveis para os [complementos](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) da página de visão geral de Outlook de complementos.

## <a name="supported-permissions"></a>Permissões com suporte

A tabela a seguir descreve as permissões que a API JavaScript Office suporta para representantes e usuários de caixa de correio compartilhados.

|Permissão|Valor|Descrição|
|---|---:|---|
|Leitura|1 (000001)|Pode ler itens.|
|Gravar|2 (000010)|Pode criar itens.|
|DeleteOwn|4 (000100)|Pode excluir apenas os itens criados.|
|DeleteAll|8 (001000)|Pode excluir qualquer item.|
|EditOwn|16 (010000)|Pode editar apenas os itens criados.|
|EditAll|32 (100000)|Pode editar todos os itens.|

> [!NOTE]
> Atualmente, a API oferece suporte para obter permissões existentes, mas não para definir permissões.

O [objeto DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma máscara de bits para indicar as permissões. Cada posição na máscara de bits representa uma permissão específica e, se estiver definida como, o `1` usuário terá a respectiva permissão. Por exemplo, se o segundo bit da direita for `1` , o usuário terá permissão **Gravar.** Você pode ver um exemplo de como verificar uma permissão específica na seção Executar uma operação como representante ou usuário de caixa de correio [compartilhada](#perform-an-operation-as-delegate-or-shared-mailbox-user) mais adiante neste artigo.

## <a name="sync-across-shared-folder-clients"></a>Sincronizar entre clientes de pasta compartilhada

As atualizações de um representante para a caixa de correio do proprietário geralmente são sincronizadas entre caixas de correio imediatamente.

No entanto, se as operações REST ou Exchange Web Services (EWS) foram usadas para definir uma propriedade estendida em um item, essas alterações podem levar algumas horas para sincronizar. Em vez disso, recomendamos que você use o [objeto CustomProperties](/javascript/api/outlook/office.customproperties) e APIs relacionadas para evitar esse atraso. Para saber mais, consulte a seção [propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "Obter e definir metadados em um Outlook de complemento".

> [!IMPORTANT]
> Em um cenário de representante, você não pode usar o EWS com os tokens atualmente fornecidos pela API office.js.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para habilitar pastas compartilhadas e cenários de caixa de correio compartilhadas no seu complemento, você deve definir o [elemento SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) como no manifesto sob `true` o elemento pai `DesktopFormFactor` . Atualmente, outros fatores de formulário não são suportados.

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

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a>Executar uma operação como representante ou usuário de caixa de correio compartilhada

Você pode obter as propriedades compartilhadas de um item no modo Redação ou Leitura chamando o [método item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) Isso retorna um [objeto SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do usuário, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.

O exemplo a seguir mostra como obter as propriedades compartilhadas de uma  mensagem ou compromisso, verificar se o representante ou usuário de caixa de correio compartilhada tem permissão Gravar e fazer uma chamada REST.

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

Dependendo dos cenários do seu complemento, há algumas limitações a considerar ao lidar com situações de pasta compartilhada ou de caixa de correio compartilhada.

### <a name="message-compose-mode"></a>Modo De composição de Mensagens

No modo Redação de Mensagem, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) não é suportado no Outlook na Web ou no Windows a menos que as seguintes condições sejam atendidas.

a. **Delegar acesso/pastas compartilhadas**

1. O proprietário da caixa de correio inicia uma mensagem. Pode ser uma nova mensagem, uma resposta ou um encaminhamento.
1. Eles salvam a mensagem e a movem de sua própria pasta **Rascunhos** para uma pasta compartilhada com o representante.
1. O representante abre o rascunho da pasta compartilhada e continua compondo.

b. **Caixa de correio compartilhada**

1. Um usuário de caixa de correio compartilhado inicia uma mensagem. Pode ser uma nova mensagem, uma resposta ou um encaminhamento.
1. Eles salvam a mensagem e a movem de sua própria pasta **Rascunhos** para uma pasta na caixa de correio compartilhada.
1. Outro usuário de caixa de correio compartilhada abre o rascunho da caixa de correio compartilhada e continua compondo.

A mensagem agora está em um contexto compartilhado e os complementos que suportam esses cenários compartilhados podem obter as propriedades compartilhadas do item. Depois que a mensagem é enviada, ela geralmente é encontrada na pasta Itens **Enviados do** remetente.

### <a name="rest-and-ews"></a>REST e EWS

Seu complemento pode usar REST e a permissão do complemento deve ser definida como para habilitar o acesso REST à caixa de correio do proprietário ou à caixa de correio compartilhada conforme `ReadWriteMailbox` aplicável. Não há suporte para EWS.

## <a name="see-also"></a>Confira também

- [Permitir que outra pessoa gerencie seu email e calendário](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Compartilhamento de calendário em Microsoft 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Adicionar uma caixa de correio compartilhada Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [Como solicitar elementos de manifesto](../develop/manifest-element-ordering.md)
- [Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Operadores de bit do JavaScript](https://www.w3schools.com/js/js_bitwise.asp)