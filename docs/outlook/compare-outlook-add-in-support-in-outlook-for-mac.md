---
title: Comparar o suporte a suplementos do Outlook no Outlook no Mac
description: Saiba como o suporte a suplementos no Outlook no Mac se compara a outros clientes do Outlook.
ms.date: 09/21/2022
ms.localizationpriority: medium
ms.openlocfilehash: c3f991865921583561e4c2db2132fad3ceba3625
ms.sourcegitcommit: 09bb0b5edd6af03c9822e1742095c7df94735120
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/23/2022
ms.locfileid: "67990410"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Comparar o suporte a suplementos do Outlook no Outlook para Mac com outros clientes do Outlook

Você pode criar e executar um suplemento do Outlook da mesma maneira no Outlook para Mac como nos outros clientes, incluindo Outlook na Web, Windows, iOS e Android, sem personalizar o JavaScript para cada cliente. As mesmas chamadas do suplemento para a API JavaScript do Office geralmente funcionam da mesma maneira, exceto para as áreas descritas na tabela a seguir.

Para saber mais, confira [implantar e instalar suplementos do Outlook para teste](testing-and-tips.md).

Para obter informações sobre o novo suporte à interface do usuário, consulte o suporte a [suplementos no Outlook na nova interface do usuário do Mac](#add-in-support-in-outlook-on-new-mac-ui).

| Área | Outlook na Web, Windows e dispositivos móveis | Outlook no Mac |
|:-----|:-----|:-----|
| Versões compatíveis do office.js e do esquema do manifesto de suplementos do Office | Todas as APIs no Office.js e esquema versão 1.1. | Todas as APIs no Office.js e esquema versão 1.1.<br><br>**OBSERVAÇÃO**: no Outlook para Mac, somente o build 16.35.308 ou posterior dá suporte ao salvamento de uma reunião. Caso contrário, o `saveAsync` método falhará quando chamado de uma reunião no modo de composição. Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa. |
| Instâncias de uma série de compromissos recorrentes | <ul><li>Pode obter a ID do item e outras propriedades de um compromisso mestre ou a instância de compromisso de uma série recorrente.</li><li>Pode usar [mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) para exibir uma instância ou o mestre de uma série recorrente.</li></ul> | <ul><li>Pode obter a ID do item e outras propriedades do compromisso mestre, mas não de uma instância de uma série recorrente.</li><li>Can display the master appointment of a recurring series. Without the item ID, cannot display an instance of a recurring series.</li></ul> |
| Tipo de destinatário do participante de um compromisso | Pode usar [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-recipienttype-member) para identificar o tipo de destinatário de um participante. | `EmailAddressDetails.recipientType` retorna `undefined` para participantes do compromisso. |
| Cadeia de caracteres de versão do aplicativo cliente | O formato da cadeia de [caracteres de versão retornada por diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) depende do tipo real de cliente. Por exemplo:<ul><li>Outlook no Windows: `15.0.4454.1002`</li><li>Outlook na Web:`15.0.918.2`</li></ul> |Um exemplo da cadeia de caracteres de versão retornada `Diagnostics.hostVersion` pelo Outlook no Mac: `15.0 (140325)` |
| Propriedades personalizadas de um item | Se a rede falhar, um suplemento ainda poderá acessar as propriedades personalizadas armazenadas em cache. | Como o Outlook no Mac não armazena em cache propriedades personalizadas, se a rede ficar inoperante, os suplementos não poderão accessá-las. |
| Detalhes de anexo | O tipo de conteúdo e os nomes de anexo em [um objeto AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) dependem do tipo de cliente:<ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` does not contain any filename extension. As an example, if the attachment is a message that has the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` always includes a filename extension. Attachments that are mail items have a .eml extension, and appointments have a .ics extension. As an example, if an attachment is an email with the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity.eml"`.<p>**Observação**: se um arquivo for anexado programaticamente (por exemplo, por meio de um suplemento) sem uma extensão, `AttachmentDetails.name` não conterá essa extensão como parte do nome do arquivo.</p></li></ul> |
| Cadeia de caracteres que representa o fuso horário nas propriedades `dateTimeCreated` e `dateTimeModified` |Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Precisão do tempo de `dateTimeCreated` e `dateTimeModified` | Se um suplemento usa o código a seguir, a precisão é de até millisecond.<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| A precisão é apenas de até um segundo. |

## <a name="add-in-support-in-outlook-on-new-mac-ui"></a>Suporte a suplementos no Outlook na nova interface do usuário do Mac

Os suplementos do Outlook agora têm suporte na nova interface do usuário do Mac (disponível na versão 16.38.506 do Outlook). Para conjuntos de requisitos com suporte atualmente na nova interface do usuário do Mac, consulte o suporte ao cliente do conjunto de [requisitos da API do Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support).

Para saber mais sobre a nova interface do usuário do Mac, confira [o novo Outlook para Mac](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439).

Você pode determinar em qual versão da interface do usuário você está, da seguinte maneira:

**Interface do usuário clássica**

![Interface do usuário clássica no Mac.](../images/outlook-on-mac-classic.png)

**Nova interface do usuário**

![Nova interface do usuário no Mac.](../images/outlook-on-mac-new.png)
