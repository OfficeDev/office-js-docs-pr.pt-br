---
title: Comparar o suporte a suplementos do Outlook no Outlook no Mac
description: Saiba como o suporte a suplementos no Outlook no Mac compara com outros clientes do Outlook.
ms.date: 06/04/2020
localization_priority: Normal
ms.openlocfilehash: f6aa9914e1320de05a67b3ec227e373bac5c2402
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293916"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Comparar o suporte a suplementos do Outlook no Outlook no Mac com outros clientes do Outlook

Você pode criar e executar um suplemento do Outlook da mesma maneira no Outlook no Mac, como nos outros clientes, incluindo Outlook na Web, Windows, iOS e Android, sem personalizar o JavaScript para cada cliente. As mesmas chamadas do suplemento para a API JavaScript do Office geralmente funcionam da mesma maneira, exceto as áreas descritas na tabela a seguir.

Para saber mais, confira [implantar e instalar suplementos do Outlook para teste](testing-and-tips.md).

Para obter informações sobre o novo suporte de interface do usuário no Mac, consulte [novo Outlook no Mac](#new-outlook-on-mac-preview).

| Área | Outlook na Web, Windows e dispositivos móveis | Outlook no Mac |
|:-----|:-----|:-----|
| Versões compatíveis do office.js e do esquema do manifesto de suplementos do Office | Todas as APIs no Office.js e esquema versão 1.1. | Todas as APIs no Office.js e esquema versão 1.1.<br><br>**Observação**: no Outlook no Mac, somente compilar o 16.35.308 ou posterior oferece suporte para salvar uma reunião. Caso contrário, o `saveAsync` método falhará quando for chamado a partir de uma reunião no modo de composição. Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa. |
| Instâncias de uma série de compromissos recorrentes | <ul><li>Pode obter a ID do item e outras propriedades de um compromisso mestre ou a instância de compromisso de uma série recorrente.</li><li>Pode usar [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para exibir uma instância ou o mestre de uma série recorrente.</li></ul> | <ul><li>Pode obter a ID do item e outras propriedades do compromisso mestre, mas não de uma instância de uma série recorrente.</li><li>Pode exibir o compromisso mestre de uma série recorrente. Sem a ID do item, não pode exibir uma instância de uma série recorrente.</li></ul> |
| Tipo de destinatário do participante de um compromisso | Pode usar [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) para identificar o tipo de destinatário de um participante. | `EmailAddressDetails.recipientType` retorna `undefined` para participantes do compromisso. |
| Cadeia de caracteres de versão do aplicativo cliente | O formato da cadeia de caracteres de versão retornada por [Diagnostics. hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) depende do tipo real de cliente. Por exemplo:<ul><li>Outlook no Windows: `15.0.4454.1002`</li><li>Outlook na Web: `15.0.918.2`</li></ul> |Um exemplo da cadeia de caracteres de versão retornada por `Diagnostics.hostVersion` no Outlook no Mac: `15.0 (140325)` |
| Propriedades personalizadas de um item | Se a rede falhar, um suplemento ainda poderá acessar as propriedades personalizadas armazenadas em cache. | Como o Outlook no Mac não armazena propriedades personalizadas em cache, se a rede for desativada, os suplementos não poderão acessá-los. |
| Detalhes de anexo | O tipo de conteúdo e os nomes de anexos em um objeto [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) dependem do tipo de cliente:<ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` não contém nenhuma extensão de nome de arquivo. Por exemplo, se o anexo é uma mensagem que tem o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` sempre inclui uma extensão de nome de arquivo. Anexos que são itens de email têm uma extensão .eml, e compromissos têm uma extensão .ics. Por exemplo, se um anexo é um email com o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity.eml"`.<p>**Observação**: se um arquivo for anexado programaticamente (por exemplo, por meio de um suplemento) sem uma extensão, `AttachmentDetails.name` não conterá essa extensão como parte do nome do arquivo.</p></li></ul> |
| Cadeia de caracteres que representa o fuso horário nas propriedades `dateTimeCreated` e `dateTimeModified` |Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Precisão do tempo de `dateTimeCreated` e `dateTimeModified` | Se um suplemento usar o código a seguir, a precisão será de até um milissegundo:<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| A precisão é de até um segundo. |

## <a name="new-outlook-on-mac-preview"></a>Novo Outlook no Mac (versão prévia)

Os suplementos do Outlook agora têm suporte na nova interface do usuário do Mac, até o conjunto de requisitos 1,6. No entanto, os seguintes conjuntos de requisitos e recursos ainda **não** têm suporte.

1. Conjuntos de requisitos de API 1,7 e 1,8
1. Painel de tarefas fixável, `ItemChanged` evento
1. Suplementos contextuais
1. Ao enviar
1. Suporte a pastas compartilhadas
1. `saveAsync` ao redigir uma reunião
1. SSO (logon único)

Recomendamos que você visualize o novo Outlook no Mac, disponível na versão 16.38.506. Para saber mais sobre como experimentá-lo, confira [Outlook para Mac-Release Notes for Insider Builds Fast](https://support.microsoft.com/office/d6347358-5613-433e-a49e-a9a0e8e0462a).

Você pode determinar qual versão da interface do usuário você está, como a seguir.

**UI atual**

&nbsp;&nbsp;&nbsp;&nbsp;![UI atual no Mac](../images/outlook-on-mac-classic.png)

**Nova interface do usuário (versão prévia)**

&nbsp;&nbsp;&nbsp;&nbsp;![Nova interface do usuário na visualização no Mac](../images/outlook-on-mac-new.png)