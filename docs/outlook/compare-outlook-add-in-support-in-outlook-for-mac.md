---
title: Comparar o suporte a suplementos do Outlook no Outlook no Mac
description: Saiba como o suporte a suplementos no Outlook no Mac compara com outros hosts do Outlook.
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: ec5e9934249ddc1a90240890d7f0201d23599446
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325458"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-hosts"></a>Comparar o suporte a suplementos do Outlook no Outlook no Mac com outros hosts do Outlook

Você pode criar e executar um suplemento do Outlook da mesma maneira no Outlook no Mac, como nos outros hosts, incluindo Outlook na Web, Windows, iOS e Android, sem personalizar o JavaScript para cada host. As mesmas chamadas do suplemento para a API JavaScript do Office geralmente funcionam da mesma maneira, exceto as áreas descritas na tabela a seguir.

Para obter mais informações, consulte [Implantar e instalar suplementos do Outlook para teste](testing-and-tips.md).

| Área | Outlook na Web, Windows e dispositivos móveis | Outlook no Mac |
|:-----|:-----|:-----|
| Versões compatíveis do office.js e do esquema do manifesto de suplementos do Office | Todas as APIs no Office.js e esquema versão 1.1. | Todas as APIs no Office.js e esquema versão 1.1.<br><br>**Observação**: o Outlook no Mac não dá suporte à gravação de uma reunião. O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição. Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa. |
| Instâncias de uma série de compromissos recorrentes | <ul><li>Pode obter a ID do item e outras propriedades de um compromisso mestre ou a instância de compromisso de uma série recorrente.</li><li>Pode usar [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para exibir uma instância ou o mestre de uma série recorrente.</li></ul> | <ul><li>Pode obter a ID do item e outras propriedades do compromisso mestre, mas não de uma instância de uma série recorrente.</li><li>Pode exibir o compromisso mestre de uma série recorrente. Sem a ID do item, não pode exibir uma instância de uma série recorrente.</li></ul> |
| Tipo de destinatário do participante de um compromisso | Pode usar [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) para identificar o tipo de destinatário de um participante. | `EmailAddressDetails.recipientType` retorna `undefined` para participantes do compromisso. |
| Cadeia de caracteres da versão do host | O formato da cadeia de caracteres de versão retornada por [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) depende do tipo real do host. Por exemplo:<ul><li>Outlook no Windows:`15.0.4454.1002`</li><li>Outlook na Web:`15.0.918.2`</li></ul> |Um exemplo da cadeia de caracteres de versão `Diagnostics.hostVersion` retornada por no Outlook no Mac:`15.0 (140325)` |
| Propriedades personalizadas de um item | Se a rede falhar, um suplemento ainda poderá acessar as propriedades personalizadas armazenadas em cache. | Como o Outlook no Mac não armazena propriedades personalizadas em cache, se a rede for desativada, os suplementos não poderão acessá-los. |
| Detalhes de anexo | Os nomes de anexos e o tipo de conteúdo em um objeto [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) dependem do tipo de host:<ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` não contém nenhuma extensão de nome de arquivo. Por exemplo, se o anexo é uma mensagem que tem o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` sempre inclui uma extensão de nome de arquivo. Anexos que são itens de email têm uma extensão .eml, e compromissos têm uma extensão .ics. Por exemplo, se um anexo é um email com o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity.eml"`.<p>**Observação**: se um arquivo for anexado programaticamente (por exemplo, por meio de um suplemento) sem uma extensão, `AttachmentDetails.name` não conterá essa extensão como parte do nome do arquivo.</p></li></ul> |
| Cadeia de caracteres que representa o fuso horário nas propriedades `dateTimeCreated` e `dateTimeModified` |Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Precisão do tempo de `dateTimeCreated` e `dateTimeModified` | Se um suplemento usar o código a seguir, a precisão será de até um milissegundo:<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| A precisão é de até um segundo. |

