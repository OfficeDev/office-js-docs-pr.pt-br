---
title: Comparar Outlook suporte a suplementos no Outlook no Mac
description: Saiba como o suporte a suplementos Outlook no Mac se compara a outros Outlook clientes.
ms.date: 06/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: 36a10f0454bebf3f069464277c7eb2a8a18f42b7
ms.sourcegitcommit: 2eeb0423a793b3a6db8a665d9ae6bcb10e867be3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/10/2022
ms.locfileid: "66019602"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Comparar Outlook suporte a suplementos Outlook no Mac com outros Outlook clientes

Você pode criar e executar um suplemento do Outlook da mesma maneira no Outlook no Mac como nos outros clientes, incluindo Outlook na Web, Windows, iOS e Android, sem personalizar o JavaScript para cada cliente. As mesmas chamadas do suplemento para a API javaScript do Office geralmente funcionam da mesma maneira, exceto para as áreas descritas na tabela a seguir.

Para saber mais, confira [implantar e instalar suplementos do Outlook para teste](testing-and-tips.md).

Para obter informações sobre o novo suporte à interface do usuário, consulte o suporte a [suplementos Outlook nova interface do usuário do Mac](#add-in-support-in-outlook-on-new-mac-ui).

| Área | Outlook na Web, Windows e dispositivos móveis | Outlook no Mac |
|:-----|:-----|:-----|
| Versões compatíveis do office.js e do esquema do manifesto de suplementos do Office | Todas as APIs no Office.js e esquema versão 1.1. | Todas as APIs no Office.js e esquema versão 1.1.<br><br>**OBSERVAÇÃO**: no Outlook no Mac, somente o build 16.35.308 ou posterior dá suporte ao salvamento de uma reunião. Caso contrário, o `saveAsync` método falhará quando chamado de uma reunião no modo de composição. Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa. |
| Instâncias de uma série de compromissos recorrentes | <ul><li>Pode obter a ID do item e outras propriedades de um compromisso mestre ou a instância de compromisso de uma série recorrente.</li><li>Pode usar [mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) para exibir uma instância ou o mestre de uma série recorrente.</li></ul> | <ul><li>Pode obter a ID do item e outras propriedades do compromisso mestre, mas não de uma instância de uma série recorrente.</li><li>Pode exibir o compromisso mestre de uma série recorrente. Sem a ID do item, não pode exibir uma instância de uma série recorrente.</li></ul> |
| Tipo de destinatário do participante de um compromisso | Pode usar [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-recipienttype-member) para identificar o tipo de destinatário de um participante. | `EmailAddressDetails.recipientType` retorna `undefined` para participantes do compromisso. |
| Cadeia de caracteres de versão do aplicativo cliente | O formato da cadeia de [caracteres de versão retornada por diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) depende do tipo real de cliente. Por exemplo:<ul><li>Outlook no Windows:`15.0.4454.1002`</li><li>Outlook na Web:`15.0.918.2`</li></ul> |Um exemplo da cadeia de caracteres de versão retornada `Diagnostics.hostVersion` por Outlook no Mac:`15.0 (140325)` |
| Propriedades personalizadas de um item | Se a rede falhar, um suplemento ainda poderá acessar as propriedades personalizadas armazenadas em cache. | Como Outlook no Mac não armazena em cache propriedades personalizadas, se a rede ficar inoperante, os suplementos não poderão accessá-las. |
| Detalhes de anexo | O tipo de conteúdo e os nomes de anexo em [um objeto AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) dependem do tipo de cliente:<ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` não contém nenhuma extensão de nome de arquivo. Por exemplo, se o anexo é uma mensagem que tem o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` sempre inclui uma extensão de nome de arquivo. Anexos que são itens de email têm uma extensão .eml, e compromissos têm uma extensão .ics. Por exemplo, se um anexo é um email com o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity.eml"`.<p>**Observação**: se um arquivo for anexado programaticamente (por exemplo, por meio de um suplemento) sem uma extensão, `AttachmentDetails.name` não conterá essa extensão como parte do nome do arquivo.</p></li></ul> |
| Cadeia de caracteres que representa o fuso horário nas propriedades `dateTimeCreated` e `dateTimeModified` |Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Precisão do tempo de `dateTimeCreated` e `dateTimeModified` | Se um suplemento usa o código a seguir, a precisão é de até millisecond.<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| A precisão é apenas de até um segundo. |

## <a name="add-in-support-in-outlook-on-new-mac-ui"></a>Suporte a suplementos no Outlook nova interface do usuário do Mac

Outlook suplementos agora têm suporte na nova interface do usuário do Mac (disponível no Outlook versão 16.38.506), até o conjunto de requisitos 1.10. No entanto, ainda não há suporte para os seguintes conjuntos **de requisitos** e recursos.

- Conjunto de requisitos de API 1.11

Para saber mais sobre a nova interface do usuário do Mac, confira [o novo Outlook para Mac](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439).

Você pode determinar em qual versão da interface do usuário você está, da seguinte maneira:

**Interface do usuário clássica**

![Interface do usuário clássica no Mac.](../images/outlook-on-mac-classic.png)

**Nova interface do usuário**

![Nova interface do usuário no Mac.](../images/outlook-on-mac-new.png)
