---
title: Habilitar pastas compartilhadas e cenários de caixa de correio compartilhada em um suplemento do Outlook
description: Discute como configurar o suporte a suplementos para pastas compartilhadas (também conhecido como acesso delegado) e caixas de correio compartilhadas.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 707be0fb71931b80314750b435dca18d23247a23
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467164"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>Habilitar pastas compartilhadas e cenários de caixa de correio compartilhada em um suplemento do Outlook

Este artigo descreve como habilitar pastas compartilhadas (também conhecidas como acesso [delegado) e](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview#shared-mailboxes) cenários de caixa de correio compartilhada (agora em versão prévia) em seu suplemento do Outlook, incluindo quais permissões a API JavaScript do Office dá suporte.

## <a name="supported-clients-and-platforms"></a>Clientes e plataformas com suporte

A tabela a seguir mostra combinações de cliente-servidor com suporte para esse recurso, incluindo a Atualização Cumulativa mínima necessária, quando aplicável. Não há suporte para combinações excluídas.

| Client | Exchange Online | Exchange 2019 local<br>(Atualização Cumulativa 1 ou posterior) | Exchange 2016 local<br>(Atualização Cumulativa 6 ou posterior) | Exchange 2013 local |
|---|:---:|:---:|:---:|:---:|
|Windows:<br>Versão 1910 (Build 12130.20272) ou posterior|Sim|Sim\*|Sim\*|Sim\*|
|Mac:<br>build 16.47 ou posterior|Sim|Sim|Sim|Sim|
|Navegador da Web:<br>interface do usuário moderna do Outlook|Sim|Não aplicável|Não aplicável|Não aplicável|
|Navegador da Web:<br>interface do usuário clássica do Outlook|Não aplicável|Não|Não|Não|

> [!NOTE]
> \* O suporte para esse recurso em um ambiente do Exchange local está disponível a partir da versão 2206 (Build 15330.20000) para o Canal Atual e a Versão 2207 (Build 15427.20000) para o Canal Empresarial Mensal.

> [!IMPORTANT]
> O suporte para esse recurso foi introduzido no conjunto de requisitos [1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) (para obter detalhes, consulte [clientes e plataformas](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)). No entanto, observe que a matriz de suporte do recurso é um superconjunto do conjunto de requisitos.

## <a name="supported-setups"></a>Configurações com suporte

As seções a seguir descrevem as configurações com suporte para caixas de correio compartilhadas (agora em versão prévia) e pastas compartilhadas. As APIs de recurso podem não funcionar conforme o esperado em outras configurações. Selecione a plataforma que você gostaria de aprender a configurar.

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>Pastas compartilhadas

O proprietário da caixa de correio [deve primeiro fornecer acesso a um delegado](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926). O representante deve seguir as instruções descritas na seção "Adicionar caixa de correio de outra pessoa ao seu perfil" do artigo Gerenciar itens de email e calendário de outra [pessoa](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5).

#### <a name="shared-mailboxes-preview"></a>Caixas de correio compartilhadas (versão prévia)

Os administradores do Exchange Server podem criar e gerenciar caixas de correio compartilhadas para conjuntos de usuários acessarem. [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) [ambientes do Exchange](/exchange/collaboration/shared-mailboxes/create-shared-mailboxes) locais e locais têm suporte.

Um Exchange Server conhecido como "automação" está ativado por padrão, o que significa que, subsequentemente, a [](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) caixa de correio compartilhada deve aparecer automaticamente no aplicativo Outlook de um usuário depois que o Outlook for fechado e reaberto. No entanto, se um administrador desativar o automação, o usuário deverá seguir as etapas manuais descritas na seção "Adicionar uma caixa de correio compartilhada ao Outlook" do artigo Abrir e usar uma caixa de correio compartilhada no [Outlook](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd).

> [!WARNING]
> NÃO **entre** na caixa de correio compartilhada com uma senha. As APIs de recurso não funcionarão nesse caso.

### <a name="web-browser---modern-outlook"></a>[Navegador da Web – Outlook moderno](#tab/modern)

#### <a name="shared-folders"></a>Pastas compartilhadas

O proprietário da caixa de [correio deve primeiro fornecer acesso a um delegado](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) atualizando as permissões de pasta da caixa de correio. O delegado deve seguir as instruções descritas na seção "Adicionar caixa de correio de outra pessoa à sua lista de pastas Outlook Web App" do artigo Acessar a caixa de correio [de outra pessoa](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081).

#### <a name="shared-mailboxes"></a>Caixas de correio compartilhadas

Atualmente, não há suporte para cenários de caixa de correio compartilhada em suplementos do Outlook no Outlook na Web.

### <a name="mac"></a>[Mac](#tab/unix)

#### <a name="shared-mailboxes-preview"></a>Caixas de correio compartilhadas (versão prévia)

Email e calendário são compartilhados com um representante ou um usuário de caixa de correio compartilhada. Os suplementos estão disponíveis para o representante ou usuário nos modos de leitura e redação de mensagens e compromissos.

#### <a name="shared-folders"></a>Pastas compartilhadas

Se a **pasta Caixa de** Entrada for compartilhada com um delegado, os suplementos estarão disponíveis para o delegado no modo de leitura de mensagem.

Se a **pasta Rascunhos** também for compartilhada com o delegado, os suplementos estarão disponíveis no modo de composição.

#### <a name="local-shared-calendar-new-model"></a>Calendário compartilhado local (novo modelo)

Se o proprietário do calendário compartilhou explicitamente seu calendário com um representante (a caixa de correio inteira pode não ser compartilhada), os suplementos estarão disponíveis para o representante nos modos de leitura e redação do compromisso.

#### <a name="remote-shared-calendar-previous-model"></a>Calendário compartilhado remoto (modelo anterior)

Se o proprietário do calendário concedeu acesso amplo ao calendário (por exemplo, tornou-o editável para uma DL específica ou toda a organização), os usuários poderão ter permissão indireta ou implícita e os suplementos estarão disponíveis para esses usuários nos modos de leitura e redação de compromissos.

---

Para saber mais sobre onde os suplementos fazem e não são ativados em geral, consulte a seção Itens de caixa de correio disponíveis para [suplementos](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) da página de visão geral de suplementos do Outlook.

## <a name="supported-permissions"></a>Permissões com suporte

A tabela a seguir descreve as permissões compatíveis com a API JavaScript do Office para representantes e usuários de caixa de correio compartilhada.

|Permissão|Valor|Descrição|
|---|---:|---|
|Ler|1 (000001)|Pode ler itens.|
|Gravar|2 (000010)|Pode criar itens.|
|DeleteOwn|4 (000100)|Pode excluir somente os itens que eles criaram.|
|DeleteAll|8 (001000)|Pode excluir todos os itens.|
|EditOwn|16 (010000)|Pode editar somente os itens que eles criaram.|
|EditAll|32 (100000)|Pode editar qualquer item.|

> [!NOTE]
> Atualmente, a API dá suporte à obtendo permissões existentes, mas não à definição de permissões.

O [objeto DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma máscara de bits para indicar as permissões. Cada posição na máscara de bits representa uma permissão específica e, `1` se estiver definida como, o usuário terá a respectiva permissão. Por exemplo, se o segundo bit da direita for `1`, o usuário terá permissão **de** gravação. Você pode ver um exemplo de como verificar se há uma permissão específica na seção [](#perform-an-operation-as-delegate-or-shared-mailbox-user) Executar uma operação como representante ou usuário de caixa de correio compartilhada posteriormente neste artigo.

## <a name="sync-across-shared-folder-clients"></a>Sincronizar entre clientes de pasta compartilhada

As atualizações de um representante para a caixa de correio do proprietário geralmente são sincronizadas entre caixas de correio imediatamente.

No entanto, se as operações REST ou EWS (Exchange Web Services) forem usadas para definir uma propriedade estendida em um item, essas alterações poderão levar algumas horas para serem sincronizadas. Em vez disso, recomendamos que você use o [objeto CustomProperties](/javascript/api/outlook/office.customproperties) e as APIs relacionadas para evitar esse atraso. Para saber mais, confira a [seção de propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "Obter e definir metadados em um suplemento do Outlook".

> [!IMPORTANT]
> Em um cenário de delegado, você não pode usar o EWS com os tokens atualmente fornecidos pela office.js API.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para habilitar pastas compartilhadas e cenários de caixa de correio compartilhada em seu suplemento, você deve habilitar as permissões necessárias no manifesto.

Primeiro, para dar suporte a chamadas REST de um delegado, o suplemento deve solicitar a permissão de caixa de correio **de leitura/** gravação. A marcação varia dependendo do tipo de manifesto.

- **Manifesto XML**: defina o **\<Permissions\>** elemento **como ReadWriteMailbox**.
- **Manifesto do Teams (** versão prévia):defina a propriedade "name" de um objeto na matriz "authorization.permissions.resourceSpecific" como "Mailbox.ReadWrite.User".

Em segundo lugar, habilite o suporte para pastas compartilhadas. A marcação varia dependendo do tipo de manifesto.

# <a name="xml-manifest"></a>[Manifesto XML](#tab/xmlmanifest)

Defina [o elemento SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) como `true` no manifesto sob o elemento pai `DesktopFormFactor`. No momento, não há suporte para outros fatores forma.

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

# <a name="teams-manifest-developer-preview"></a>[Manifesto do Teams (versão prévia do desenvolvedor)](#tab/jsonmanifest)

Adicione um objeto adicional à matriz "authorization.permissions.resourceSpecific" e defina sua propriedade "name" como "Mailbox.SharedFolder".

```json
"authorization": {
  "permissions": {
    "resourceSpecific": [
      ...
      {
        "name": "Mailbox.SharedFolder",
        "type": "Delegated"
      },
    ]
  }
},
```

---

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a>Executar uma operação como representante ou usuário de caixa de correio compartilhada

Você pode obter as propriedades compartilhadas de um item no modo Redigir ou Ler chamando o [método item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) . Isso retorna [um objeto SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do usuário, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.

O exemplo a seguir mostra como obter as propriedades compartilhadas de uma mensagem ou compromisso, verificar se o representante ou usuário da caixa  de correio compartilhada tem permissão de Gravação e fazer uma chamada REST.

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
> Como representante, você pode usar REST para obter o conteúdo de uma mensagem do [Outlook anexada a um item ou postagem de grupo do Outlook](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Manipular chamada REST em itens compartilhados e não compartilhados

Se você quiser chamar uma operação REST em um item, se o item for compartilhado ou não, `getSharedPropertiesAsync` poderá usar a API para determinar se o item é compartilhado. Depois disso, você pode construir a URL REST para a operação usando o objeto apropriado.

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a>Limitações

Dependendo dos cenários do suplemento, há algumas limitações a serem consideradas ao lidar com situações de pasta compartilhada ou caixa de correio compartilhada.

### <a name="message-compose-mode"></a>Modo de composição de mensagem

No modo Redigir Mensagem, não há suporte para [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1)) no Outlook na Web ou no Windows, a menos que as condições a seguir sejam atendidas.

a. **Delegar acesso/pastas compartilhadas**

1. O proprietário da caixa de correio inicia uma mensagem. Pode ser uma nova mensagem, uma resposta ou um encaminhamento.
1. Eles salvam a mensagem e a movem de sua própria pasta **Rascunhos** para uma pasta compartilhada com o representante.
1. O delegado abre o rascunho da pasta compartilhada e continua redigindo.

b. **Caixa de correio compartilhada (aplica-se somente ao Outlook no Windows)**

1. Um usuário de caixa de correio compartilhada inicia uma mensagem. Pode ser uma nova mensagem, uma resposta ou um encaminhamento.
1. Eles salvam a mensagem e a movem da própria pasta **Rascunhos** para uma pasta na caixa de correio compartilhada.
1. Outro usuário de caixa de correio compartilhada abre o rascunho da caixa de correio compartilhada e continua redigindo.

A mensagem agora está em um contexto compartilhado e os suplementos que dão suporte a esses cenários compartilhados podem obter as propriedades compartilhadas do item. Depois que a mensagem for enviada, ela geralmente será encontrada na pasta Itens **Enviados do** remetente.

### <a name="rest-and-ews"></a>REST e EWS

Seu suplemento pode usar REST. Para habilitar o acesso REST à caixa de correio do proprietário ou à caixa de correio compartilhada conforme aplicável, o suplemento deve solicitar a permissão de caixa de correio de leitura **/** gravação no manifesto. A marcação varia dependendo do tipo de manifesto.

- **Manifesto XML**: defina o **\<Permissions\>** elemento **como ReadWriteMailbox**.
- **Manifesto do Teams (** versão prévia):defina a propriedade "name" de um objeto na matriz "authorization.permissions.resourceSpecific" como "Mailbox.ReadWrite.User".

Não há suporte para EWS.

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>Usuário ou caixa de correio compartilhada oculta de uma lista de endereços

Se um administrador escondeu um usuário ou endereço de caixa de correio compartilhado de uma lista de endereços, como a GAL (lista de endereços global), os itens de email `Office.context.mailbox.item` afetados foram abertos no relatório da caixa de correio como nulo. Por exemplo, se o usuário abrir um item de email em uma caixa de correio compartilhada ocultada da GAL, `Office.context.mailbox.item` representar esse item de email será nulo.

## <a name="see-also"></a>Confira também

- [Permitir que outra pessoa gerencie seu e-mail e seu calendário](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [Compartilhamento de calendário no Microsoft 365](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Adicionar uma caixa de correio compartilhada ao Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [Como ordenar elementos de manifesto](../develop/manifest-element-ordering.md)
- [Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Operadores bit a bit do JavaScript](https://www.w3schools.com/js/js_bitwise.asp)
