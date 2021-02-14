---
title: Comparar o suporte a um complemento do Outlook no Outlook no Mac
description: Saiba como o suporte a um complemento no Outlook para Mac se compara a outros clientes do Outlook.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: 83cebf20cc4ead4bb50fd1a49653ac15f8501792
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234265"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Comparar o suporte a um complemento do Outlook no Outlook para Mac com outros clientes do Outlook

Você pode criar e executar um complemento do Outlook da mesma maneira no Outlook no Mac como em outros clientes, incluindo o Outlook na Web, Windows, iOS e Android, sem personalizar o JavaScript para cada cliente. As mesmas chamadas do complemento para a API JavaScript do Office geralmente funcionam da mesma maneira, exceto para as áreas descritas na tabela a seguir.

Para saber mais, confira [implantar e instalar suplementos do Outlook para teste](testing-and-tips.md).

For information about new UI support, see [Add-in support in Outlook on new Mac UI](#add-in-support-in-outlook-on-new-mac-ui-preview).

| Área | Outlook na Web, Windows e dispositivos móveis | Outlook no Mac |
|:-----|:-----|:-----|
| Versões compatíveis do office.js e do esquema do manifesto de suplementos do Office | Todas as APIs no Office.js e esquema versão 1.1. | Todas as APIs no Office.js e esquema versão 1.1.<br><br>**OBSERVAÇÃO:** no Outlook para Mac, apenas o build 16.35.308 ou posterior oferece suporte ao salvar uma reunião. Caso contrário, `saveAsync` o método falhará quando chamado de uma reunião no modo de redação. Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa. |
| Instâncias de uma série de compromissos recorrentes | <ul><li>Pode obter a ID do item e outras propriedades de um compromisso mestre ou a instância de compromisso de uma série recorrente.</li><li>Pode usar [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para exibir uma instância ou o mestre de uma série recorrente.</li></ul> | <ul><li>Pode obter a ID do item e outras propriedades do compromisso mestre, mas não de uma instância de uma série recorrente.</li><li>Pode exibir o compromisso mestre de uma série recorrente. Sem a ID do item, não pode exibir uma instância de uma série recorrente.</li></ul> |
| Tipo de destinatário do participante de um compromisso | Pode usar [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) para identificar o tipo de destinatário de um participante. | `EmailAddressDetails.recipientType` retorna `undefined` para participantes do compromisso. |
| Cadeia de caracteres de versão do aplicativo cliente | O formato da cadeia de caracteres de versão retornada [por diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) depende do tipo real de cliente. Por exemplo:<ul><li>Outlook no Windows: `15.0.4454.1002`</li><li>Outlook na Web: `15.0.918.2`</li></ul> |Um exemplo da cadeia de caracteres de versão `Diagnostics.hostVersion` retornada pelo Outlook no Mac: `15.0 (140325)` |
| Propriedades personalizadas de um item | Se a rede falhar, um suplemento ainda poderá acessar as propriedades personalizadas armazenadas em cache. | Como o Outlook no Mac não armazena propriedades personalizadas em cache, se a rede ficar inoca, os complementos não poderão acessá-las. |
| Detalhes de anexo | O tipo de conteúdo e os nomes de anexo em [um objeto AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) dependem do tipo de cliente:<ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` não contém nenhuma extensão de nome de arquivo. Por exemplo, se o anexo é uma mensagem que tem o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Um exemplo JSON de `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` sempre inclui uma extensão de nome de arquivo. Anexos que são itens de email têm uma extensão .eml, e compromissos têm uma extensão .ics. Por exemplo, se um anexo é um email com o assunto "RES: Atividade de verão", o objeto JSON que representa o nome do anexo é `"name": "RE: Summer activity.eml"`.<p>**Observação**: se um arquivo for anexado programaticamente (por exemplo, por meio de um suplemento) sem uma extensão, `AttachmentDetails.name` não conterá essa extensão como parte do nome do arquivo.</p></li></ul> |
| Cadeia de caracteres que representa o fuso horário nas propriedades `dateTimeCreated` e `dateTimeModified` |Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Como exemplo: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Precisão do tempo de `dateTimeCreated` e `dateTimeModified` | Se um suplemento usar o código a seguir, a precisão será de até um milissegundo:<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| A precisão é de até um segundo. |

## <a name="add-in-support-in-outlook-on-new-mac-ui-preview"></a>Suporte a um complemento no Outlook na nova interface do usuário do Mac (visualização)

Os complementos do Outlook agora têm suporte na nova interface do usuário do Mac (visualização), até o conjunto de requisitos 1.7. No entanto, os seguintes conjuntos de requisitos e recursos **AINDA NÃO** têm suporte.

1. Conjuntos de requisitos de API 1.8 e 1.9
1. Suplementos contextuais
1. Ao enviar
1. Pop-out de janela de redação
1. Suporte a pastas compartilhadas
1. `saveAsync` ao compor uma reunião

Recomendamos que você visualize o Outlook na nova interface do usuário do Mac, disponível na versão 16.38.506. Para saber mais sobre como experimentar, confira Outlook para Mac - Notas de versão para [builds do Insider Fast.](https://support.microsoft.com/office/d6347358-5613-433e-a49e-a9a0e8e0462a)

Você pode determinar em qual versão da interface do usuário você está, da seguinte forma.

**Interface do usuário atual**

&nbsp;&nbsp;&nbsp;&nbsp;![Interface do usuário atual no Mac](../images/outlook-on-mac-classic.png)

**Nova interface do usuário (visualização)**

&nbsp;&nbsp;&nbsp;&nbsp;![Nova interface do usuário em visualização no Mac](../images/outlook-on-mac-new.png)
