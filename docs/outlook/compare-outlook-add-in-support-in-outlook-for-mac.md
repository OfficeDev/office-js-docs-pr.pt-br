---
title: Comparar Outlook suporte a um Outlook no Mac
description: Saiba como o suporte ao Outlook no Mac se compara a outros Outlook clientes.
ms.date: 07/01/2021
localization_priority: Normal
ms.openlocfilehash: 23bdd6938cb7c795a2b652f4649714dcbb46d446175d87a889b5c8ecffcf718a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096994"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Comparar Outlook suporte a um Outlook no Mac com outros Outlook clientes

Você pode criar e executar um Outlook do mesmo modo no Outlook no Mac, como nos outros clientes, incluindo Outlook na Web, Windows, iOS e Android, sem personalizar o JavaScript para cada cliente. As mesmas chamadas do add-in para a API JavaScript Office geralmente funcionam da mesma maneira, exceto para as áreas descritas na tabela a seguir.

Para saber mais, confira [implantar e instalar suplementos do Outlook para teste](testing-and-tips.md).

Para obter informações sobre o novo suporte à interface do usuário, consulte Suporte ao [Outlook nova interface do usuário do Mac.](#add-in-support-in-outlook-on-new-mac-ui-preview)

| Área | Outlook na Web, Windows e dispositivos móveis | Outlook no Mac |
|:-----|:-----|:-----|
| Versões compatíveis do office.js e do esquema do manifesto de suplementos do Office | Todas as APIs no Office.js e esquema versão 1.1. | Todas as APIs no Office.js e esquema versão 1.1.<br><br>**OBSERVAÇÃO**: Outlook no Mac, apenas a com build 16.35.308 ou posterior oferece suporte para salvar uma reunião. Caso contrário, `saveAsync` o método falhará quando chamado de uma reunião no modo de redação. Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa. |
| Instâncias de uma série de compromissos recorrentes | <ul><li>Pode obter a ID do item e outras propriedades de um compromisso mestre ou a instância de compromisso de uma série recorrente.</li><li>Pode usar [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para exibir uma instância ou o mestre de uma série recorrente.</li></ul> | <ul><li>Pode obter a ID do item e outras propriedades do compromisso mestre, mas não de uma instância de uma série recorrente.</li><li>Pode exibir o compromisso mestre de uma série recorrente. Sem a ID do item, não pode exibir uma instância de uma série recorrente.</li></ul> |
| Tipo de destinatário do participante de um compromisso | Pode usar [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipientType) para identificar o tipo de destinatário de um participante. | `EmailAddressDetails.recipientType` retorna `undefined` para participantes do compromisso. |
| Cadeia de caracteres de versão do aplicativo cliente | O formato da cadeia de caracteres de versão retornada [por diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostVersion) depende do tipo real de cliente. Por exemplo:<ul><li>Outlook no Windows:`15.0.4454.1002`</li><li>Outlook na Web:`15.0.918.2`</li></ul> |Um exemplo da cadeia de caracteres de versão `Diagnostics.hostVersion` retornada por Outlook no Mac:`15.0 (140325)` |
| Propriedades personalizadas de um item | Se a rede falhar, um suplemento ainda poderá acessar as propriedades personalizadas armazenadas em cache. | Como Outlook no Mac não armazena propriedades personalizadas em cache, se a rede ficar para baixo, os complementos não poderão acessá-las. |
| Detalhes de anexo | O tipo de conteúdo e os nomes de anexos em [um objeto AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) dependem do tipo de cliente:<ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` não contém nenhuma extensão de nome de arquivo. Por exemplo, se o anexo é uma mensagem que tem o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` sempre inclui uma extensão de nome de arquivo. Anexos que são itens de email têm uma extensão .eml, e compromissos têm uma extensão .ics. Por exemplo, se um anexo é um email com o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity.eml"`.<p>**Observação**: se um arquivo for anexado programaticamente (por exemplo, por meio de um suplemento) sem uma extensão, `AttachmentDetails.name` não conterá essa extensão como parte do nome do arquivo.</p></li></ul> |
| Cadeia de caracteres que representa o fuso horário nas propriedades `dateTimeCreated` e `dateTimeModified` |Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Precisão do tempo de `dateTimeCreated` e `dateTimeModified` | Se um suplemento usa o código a seguir, a precisão é de até millisecond.<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| A precisão é apenas de até um segundo. |

## <a name="add-in-support-in-outlook-on-new-mac-ui-preview"></a>Suporte ao complemento no Outlook nova interface do usuário do Mac (visualização)

Outlook os complementos agora são suportados na nova interface do usuário do Mac (visualização), até o conjunto de requisitos 1.8. No entanto, os seguintes conjuntos de requisitos e recursos **ainda não** são suportados.

- Os requisitos de API definem 1,9 e 1,10

Recomendamos que você visualize Outlook na nova interface do usuário do Mac, disponível na versão 16.38.506. Para saber mais sobre como experimentar, consulte Outlook para Mac - Notas de versão para [builds do Insider Fast.](https://support.microsoft.com/office/d6347358-5613-433e-a49e-a9a0e8e0462a)

Você pode determinar qual versão da interface do usuário você está, da seguinte forma:

**Interface do usuário atual**

![Interface do usuário atual no Mac.](../images/outlook-on-mac-classic.png)

**Nova interface do usuário (visualização)**

![Nova interface do usuário em visualização no Mac.](../images/outlook-on-mac-new.png)
