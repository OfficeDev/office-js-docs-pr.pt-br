---
title: Habilitar pastas compartilhadas e cenários de caixa de correio compartilhadas em um Outlook de entrada
description: Discute como configurar o suporte ao complemento para pastas compartilhadas (a.k.a. acesso delegado) e caixas de correio compartilhadas.
ms.date: 10/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: e359f4b63aec979d68b0798866fb06bf559a0f67
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484653"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>Habilitar pastas compartilhadas e cenários de caixa de correio compartilhadas em um Outlook de entrada

Este artigo descreve como habilitar pastas compartilhadas (também conhecidas como acesso de [representante) e](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview#shared-mailboxes) cenários de caixa de correio compartilhada (agora em visualização) no seu Outlook add-in, incluindo quais permissões Office API JavaScript suporta.

## <a name="supported-clients-and-platforms"></a>Clientes e plataformas com suporte

A tabela a seguir mostra combinações de cliente-servidor com suporte para esse recurso, incluindo a Atualização Cumulativa mínima necessária, quando aplicável. Não há suporte para combinações excluídas.

| Client | Exchange Online | Exchange 2019 local<br>(Atualização Cumulativa 1 ou posterior) | Exchange 2016 local<br>(Atualização Cumulativa 6 ou posterior) | Exchange 2013 local |
|---|:---:|:---:|:---:|:---:|
|Windows:<br>versão 1910 (build 12130.20272) ou posterior|Sim|Não|Não|Não|
|Mac:<br>build 16.47 ou posterior|Sim|Sim|Sim|Sim|
|Navegador da Web:<br>interface do usuário Outlook moderna|Sim|Não aplicável|Não aplicável|Não aplicável|
|Navegador da Web:<br>interface do usuário Outlook clássica|Não aplicável|Não|Não|Não|

> [!IMPORTANT]
> O suporte para esse recurso foi introduzido no [conjunto de requisitos 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) (para obter detalhes, consulte [clientes e plataformas](/javascript/api/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)). No entanto, observe que a matriz de suporte do recurso é um superconjunto do conjunto de requisitos.

## <a name="supported-setups"></a>Configurações com suporte

As seções a seguir descrevem configurações com suporte para caixas de correio compartilhadas (agora em visualização) e pastas compartilhadas. As APIs de recurso podem não funcionar conforme o esperado em outras configurações. Selecione a plataforma que você gostaria de aprender a configurar.

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>Pastas compartilhadas

O proprietário da caixa de correio [deve primeiro fornecer acesso a um representante](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926). O representante deve seguir as instruções descritas na seção "Adicionar caixa de correio de outra pessoa ao seu perfil" do artigo Gerenciar itens de email e [calendário de outra pessoa](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5).

#### <a name="shared-mailboxes-preview"></a>Caixas de correio compartilhadas (visualização)

Exchange administradores de servidor podem criar e gerenciar caixas de correio compartilhadas para conjuntos de usuários acessarem. No momento, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) é a única versão de servidor com suporte para esse recurso.

Um recurso Exchange Server conhecido como "automapping" está ativado por padrão, o que significa que, posteriormente, [](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) a caixa de correio compartilhada deve aparecer automaticamente no aplicativo Outlook do usuário depois que o Outlook tiver sido fechado e reaberto. No entanto, se um administrador tiver desabilitado a automação, o usuário deverá seguir as etapas manuais descritas na seção "Adicionar uma caixa de correio compartilhada ao Outlook" do artigo [Abrir](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd) e usar uma caixa de correio compartilhada no Outlook.

> [!WARNING]
> Não **entre** na caixa de correio compartilhada com uma senha. As APIs de recurso não funcionarão nesse caso.

### <a name="web-browser---modern-outlook"></a>[Navegador da Web – Outlook moderno](#tab/modern)

#### <a name="shared-folders"></a>Pastas compartilhadas

O proprietário da caixa de correio [deve primeiro fornecer acesso a um representante](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) atualizando as permissões de pasta de caixa de correio. O representante deve seguir as instruções descritas na seção "Adicionar caixa de correio de outra pessoa à sua lista de pastas Outlook Web App" do artigo Acessar a caixa de correio [de outra pessoa](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081).

#### <a name="shared-mailboxes-preview"></a>Caixas de correio compartilhadas (visualização)

Exchange administradores de servidor podem criar e gerenciar caixas de correio compartilhadas para conjuntos de usuários acessarem. No momento, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) é a única versão de servidor com suporte para esse recurso.

Depois de receber acesso, um usuário de caixa de correio compartilhada deve seguir as etapas descritas na seção "Adicionar a caixa de correio compartilhada para que ela seja exibida sob sua caixa de correio principal" do artigo [Abrir](https://support.microsoft.com/office/98b5a90d-4e38-415d-a030-f09a4cd28207) e usar uma caixa de correio compartilhada no Outlook na Web.

> [!WARNING]
> NÃO **use** outras opções como "Abrir outra caixa de correio". As APIs de recurso podem não funcionar corretamente.

### <a name="mac"></a>[Mac](#tab/unix)

#### <a name="shared-mailboxes-preview"></a>Caixas de correio compartilhadas (visualização)

Email e calendário são compartilhados com um representante ou usuário de caixa de correio compartilhado. Os complementos estão disponíveis para o representante ou usuário nos modos de leitura e redação de mensagens e compromissos.

#### <a name="shared-folders"></a>Pastas compartilhadas

Se a **pasta Caixa de** Entrada for compartilhada com um representante, os complementos estarão disponíveis para o representante no modo de leitura de mensagens.

Se a **pasta Rascunhos** também for compartilhada com o representante, os complementos estarão disponíveis no modo de redação.

#### <a name="local-shared-calendar-new-model"></a>Calendário compartilhado local (novo modelo)

Se o proprietário do calendário compartilhou explicitamente seu calendário com um representante (a caixa de correio inteira pode não ser compartilhada), os complementos estarão disponíveis para o representante nos modos de leitura e redação do compromisso.

#### <a name="remote-shared-calendar-previous-model"></a>Calendário compartilhado remoto (modelo anterior)

Se o proprietário do calendário concedeu amplo acesso ao calendário (por exemplo, o tornou editável para um DL específico ou para toda a organização), os usuários poderão ter permissão indireta ou implícita e os complementos estarão disponíveis para esses usuários nos modos de leitura e redação de compromissos.

---

Para saber mais sobre onde os complementos fazem e não são ativados em geral, consulte a seção Itens de Caixa de Correio disponíveis para os [complementos](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) da página de visão geral de Outlook de complementos.

## <a name="supported-permissions"></a>Permissões com suporte

A tabela a seguir descreve as permissões que a API JavaScript Office suporta para representantes e usuários de caixa de correio compartilhados.

|Permissão|Valor|Descrição|
|---|---:|---|
|Read|1 (000001)|Pode ler itens.|
|Gravar|2 (000010)|Pode criar itens.|
|DeleteOwn|4 (000100)|Pode excluir apenas os itens criados.|
|DeleteAll|8 (001000)|Pode excluir qualquer item.|
|EditOwn|16 (010000)|Pode editar apenas os itens criados.|
|EditAll|32 (100000)|Pode editar todos os itens.|

> [!NOTE]
> Atualmente, a API oferece suporte para obter permissões existentes, mas não para definir permissões.

O [objeto DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma máscara de bits para indicar as permissões. Cada posição na máscara de bits representa uma permissão específica e, se estiver definida `1` como, o usuário terá a respectiva permissão. Por exemplo, se o segundo bit da direita for `1`, o usuário terá **permissão Gravar** . Você pode ver um exemplo de como verificar uma permissão específica na seção Executar uma operação como representante ou usuário de caixa de [correio compartilhada mais](#perform-an-operation-as-delegate-or-shared-mailbox-user) adiante neste artigo.

## <a name="sync-across-shared-folder-clients"></a>Sincronizar entre clientes de pasta compartilhada

As atualizações de um representante para a caixa de correio do proprietário geralmente são sincronizadas entre caixas de correio imediatamente.

No entanto, se as operações REST ou Exchange Web Services (EWS) foram usadas para definir uma propriedade estendida em um item, essas alterações podem levar algumas horas para sincronizar. Em vez disso, recomendamos que você use o [objeto CustomProperties](/javascript/api/outlook/office.customproperties) e APIs relacionadas para evitar esse atraso. Para saber mais, consulte a seção [propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "Obter e definir metadados em um Outlook de complemento".

> [!IMPORTANT]
> Em um cenário de representante, você não pode usar o EWS com os tokens atualmente fornecidos pela API office.js.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para habilitar pastas compartilhadas e cenários de caixa de correio compartilhadas no seu complemento, você deve definir o [elemento SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) como `true` no manifesto sob o elemento pai `DesktopFormFactor`. Atualmente, outros fatores de formulário não são suportados.

Para dar suporte a chamadas REST de um representante, de definir o nó [Permissões](/javascript/api/manifest/permissions) no manifesto como `ReadWriteMailbox`.

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

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a>Executar uma operação como representante ou usuário de caixa de correio compartilhada

Você pode obter as propriedades compartilhadas de um item no modo Redação ou Leitura chamando o [método item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) . Isso retorna um [objeto SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do usuário, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.

O exemplo a seguir mostra como obter as propriedades compartilhadas de uma mensagem ou compromisso, verificar se o representante ou usuário de caixa  de correio compartilhada tem permissão Gravar e fazer uma chamada REST.

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
> Como representante, você pode usar REST para obter o conteúdo de uma mensagem [Outlook anexada a um item Outlook ou postagem de grupo](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Manipular a chamada REST em itens compartilhados e não compartilhados

Se você quiser chamar uma operação REST em um item, se o item é compartilhado ou não, `getSharedPropertiesAsync` você pode usar a API para determinar se o item é compartilhado. Depois disso, você pode construir a URL REST para a operação usando o objeto apropriado.

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

No modo Redação de Mensagem, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1)) não é suportado no Outlook na Web ou no Windows a menos que as seguintes condições sejam atendidas.

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

Seu complemento pode usar REST `ReadWriteMailbox` e a permissão do complemento deve ser definida como para habilitar o acesso REST à caixa de correio do proprietário ou à caixa de correio compartilhada conforme aplicável. Não há suporte para EWS.

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>Usuário ou caixa de correio compartilhada oculta de uma lista de endereços

Se um administrador ocultou um usuário ou endereço de caixa de correio compartilhado de uma lista de endereços, como a GAL (lista de endereços global), os itens de email `Office.context.mailbox.item` afetados abriram no relatório de caixa de correio como nulos. Por exemplo, se o usuário abrir um item de email em uma caixa de correio compartilhada oculta da GAL, `Office.context.mailbox.item` representar esse item de email será nulo.

## <a name="see-also"></a>Confira também

- [Permitir que outra pessoa gerencie seu email e calendário](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [Compartilhamento de calendário em Microsoft 365](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Adicionar uma caixa de correio compartilhada ao Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [Como solicitar elementos de manifesto](../develop/manifest-element-ordering.md)
- [Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Operadores de bit do JavaScript](https://www.w3schools.com/js/js_bitwise.asp)