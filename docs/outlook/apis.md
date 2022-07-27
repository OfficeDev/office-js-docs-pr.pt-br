---
title: APIs de suplemento do Outlook
description: Saiba como fazer referência a APIs de suplemento do Outlook e declarar permissões em seu suplemento do Outlook.
ms.date: 07/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: bd0bcdd3dfac6def9443b09d9797bfd0667c3b3d
ms.sourcegitcommit: 15714ef1118083032e640413ede69a162c43daed
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/27/2022
ms.locfileid: "67037812"
---
# <a name="outlook-add-in-apis"></a>APIs de suplemento do Outlook

Para usar APIs no seu suplemento do Outlook, você deve especificar o local da biblioteca Office.js, o conjunto de requisitos, o esquema e as permissões. Você usará principalmente as APIs JavaScript do Office expostas por meio do [objeto Mailbox](#mailbox-object) .

## <a name="officejs-library"></a>Biblioteca Office.js

Para interagir com a [API de suplemento do Outlook](/javascript/api/outlook), você precisa usar as APIs JavaScript em Office.js. A CDN (rede de distribuição de conteúdo) da biblioteca é `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. Suplementos enviados ao AppSource devem fazer referência ao Office.js por essa CDN. Eles não podem usar uma referência local.

Referência CDN em um `<script>` marca na `<head>` marca da página da web (arquivo. HTML,. aspx ou. PHP) implementa interface do usuário do seu suplemento.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

À medida que adicionamos novas APIs, a URL para Office.js permanecerá a mesma. Somente mudaremos a versão na URL se mudarmos um comportamento de API existente.

> [!IMPORTANT]
> Ao desenvolver um suplemento para qualquer aplicativo cliente do Office, faça referência à API `<head>` JavaScript do Office de dentro da seção da página. Isso garante que a API seja totalmente inicializada antes de qualquer elemento de corpo.

## <a name="requirement-sets"></a>Conjuntos de requisitos

Todas as APIs do Outlook pertencem ao conjunto [de requisitos de Caixa de Correio](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets). O conjunto de requisitos `Mailbox` tem versões, e cada novo conjunto de APIs lançado pertence a uma versão superior. Nem todos os clientes do Outlook terão suporte ao conjunto mais recente de APIs quando for lançado, mas se um cliente do Outlook declarar suporte a um conjunto de requisitos, ele dará suporte a todas as APIs nesse conjunto.

Especifique uma versão mínima de conjunto de requisitos no manifesto para controlar em quais clientes do Outlook o suplemento aparecerá. Por exemplo, se você especificar a versão 1.3 do conjunto de requisitos, o suplemento não aparecerá nos clientes do Outlook incompatíveis com a versão mínima 1.3.

A especificação de um conjunto de requisitos não limita seu suplemento às APIs nessa versão. Se o suplemento especificar a versão 1.1 do conjunto de requisitos, mas estiver sendo executado em um cliente do Outlook que dá suporte à versão 1.3, ele poderá usar as APIs v1.3. O conjunto de requisitos controla somente quais clientes do Outlook exibirão o suplemento.

Para verificar a disponibilidade das APIs de um conjunto de requisitos superior ao especificado no manifesto, use JavaScript padrão:

```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> [!NOTE]
> essas verificações não são necessárias para APIs que estão na versão do conjunto de requisitos especificada no manifesto.

Especifique o conjunto de requisitos mínimo que proporciona suporte ao conjunto essencial de APIs para seu cenário, sem o qual os recursos do suplemento não funcionam. Especifique o conjunto de requisitos no manifesto. A marcação varia dependendo do manifesto que você está usando. 

- **Manifesto XML**: use o **\<Requirements\>** elemento. Observe que não **\<Methods\>** há suporte **\<Requirements\>** para o elemento filho em suplementos do Outlook, portanto, você não pode declarar suporte para métodos específicos.
- **Manifesto do Teams (versão prévia):** use a propriedade "extensions.capabilities". 

Para obter mais informações, consulte [manifestos de suplemento do Outlook](manifests.md) e [Noções básicas sobre conjuntos de requisitos da API do Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

## <a name="permissions"></a>Permissões

Seu suplemento requer as permissões apropriadas para usar as APIs de que precisa. Em geral, você deve especificar a permissão mínima necessária para o seu suplemento. As permissões são declaradas no manifesto. A marcação varia dependendo do tipo de manifesto.

- **Manifesto XML**: use o **\<Permissions\>** elemento.
- **Manifesto do Teams (versão prévia):** use a propriedade "authorization.permissions.resourceSpecific". 

Há quatro níveis de permissões. Para obter mais detalhes, confira [Noções básicas sobre suplementos do Outlook](understanding-outlook-add-in-permissions.md).

<br/>

|Nível da permissão</br>Nome do manifesto XML|Nível da permissão</br>Nome do manifesto do Teams|Descrição|
|:-----|:-----|:-----|
| **Restrito** | **MailboxItem.Restricted.User** | Permite o uso de entidades, mas não de expressões regulares. |
| **ReadItem** | **MailboxItem.Read.User** | Além do que é permitido em **Restrito**, ele permite:<ul><li>expressões regulares</li><li>acesso de leitura para a API do suplemento do Outlook</li><li>obter as propriedades do item e o token de retorno de chamada</li></ul> |
| **ReadWriteItem** | **MailboxItem.ReadWrite.User** | Além do que é permitido no **ReadItem**, ele permite:<ul><li>acesso completo à API do Suplemento do Outlook, exceto `makeEwsRequestAsync`</li><li>definição das propriedades do item</li></ul> |
| **ReadWriteMailbox** | **Mailbox.ReadWrite.User** | Além do que é permitido em **ReadWriteItem**, ele permite:<ul><li>criar, ler, gravar itens e pastas</li><li>enviar itens</li><li>chamar [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)</li></ul> |

> [!NOTE]
> Há uma permissão complementar necessária para suplementos que usam o recurso de acréscimo ao enviar. Com o manifesto XML, você especifica a permissão no [elemento ExtendedPermissions](/javascript/api/manifest/extendedpermissions) . Para obter detalhes [, consulte Implementar append-on-send em seu suplemento do Outlook](append-on-send.md). Com o manifesto do Teams (versão prévia), você especifica essa permissão com o nome **Mailbox.AppendOnSend.User** em um objeto adicional na matriz "authorization.permissions.resourceSpecific".

Para saber mais, confira [Manifestos de suplementos do Outlook](manifests.md). Para obter informações sobre problemas de segurança, consulte [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md).

## <a name="mailbox-object"></a>Objeto Mailbox

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>Confira também

- [Manifestos de suplementos do Outlook](manifests.md)
- [Noções básicas sobre conjuntos de requisitos da API do Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
