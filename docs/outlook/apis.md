---
title: APIs de suplemento do Outlook
description: Saiba como fazer referência a APIs de suplemento do Outlook e declarar permissões em seu suplemento do Outlook.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 69043646add5e32502efb0d2a5b1259667e564d9
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467073"
---
# <a name="outlook-add-in-apis"></a>APIs de suplemento do Outlook

Para usar APIs no seu suplemento do Outlook, você deve especificar o local da biblioteca Office.js, o conjunto de requisitos, o esquema e as permissões. Você usará principalmente as APIs JavaScript do Office expostas por meio do [objeto Mailbox](#mailbox-object) .

## <a name="officejs-library"></a>Biblioteca Office.js

Para interagir com a [API de suplemento do Outlook](/javascript/api/outlook), você precisa usar as APIs JavaScript em Office.js. A CDN (rede de distribuição de conteúdo) da biblioteca é `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. Suplementos enviados ao AppSource devem fazer referência ao Office.js por essa CDN. Eles não podem usar uma referência local.

Referência CDN em um `<script>` marca na `<head>` marca da página da web (arquivo. HTML,. aspx ou. PHP) implementa interface do usuário do seu suplemento.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

As we add new APIs, the URL to Office.js will stay the same. We will change the version in the URL only if we break an existing API behavior.

> [!IMPORTANT]
> Ao desenvolver um suplemento para qualquer aplicativo cliente do Office, faça referência à API `<head>` JavaScript do Office de dentro da seção da página. Isso garante que a API seja totalmente inicializada antes de qualquer elemento de corpo.

## <a name="requirement-sets"></a>Conjuntos de requisitos

Todas as APIs do Outlook pertencem ao conjunto [de requisitos de Caixa de Correio](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets). O conjunto de requisitos `Mailbox` tem versões, e cada novo conjunto de APIs lançado pertence a uma versão superior. Nem todos os clientes do Outlook terão suporte ao conjunto mais recente de APIs quando for lançado, mas se um cliente do Outlook declarar suporte a um conjunto de requisitos, ele dará suporte a todas as APIs nesse conjunto.

To control which Outlook clients the add-in appears in, specify a minimum requirement set version in the manifest. For example, if you specify requirement set version 1.3, the add-in will not show up in any Outlook client that doesn't support a minimum version of 1.3.

Specifying a requirement set doesn't limit your add-in to the APIs in that version. If the add-in specifies requirement set v1.1 but is running in an Outlook client that supports v1.3, the add-in can still use v1.3 APIs. The requirement set only controls which Outlook clients the add-in appears in.

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

Seu suplemento requer as permissões apropriadas para usar as APIs de que precisa. Em geral, você deve especificar a permissão mínima necessária para o seu suplemento.

Há quatro níveis de permissões; **restrito**, item **de leitura**, **item de leitura/gravação** e caixa **de correio de leitura/gravação**. Para obter mais detalhes. Para obter mais detalhes, confira [Noções básicas sobre suplementos do Outlook](understanding-outlook-add-in-permissions.md).

## <a name="mailbox-object"></a>Objeto Mailbox

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>Confira também

- [Manifestos de suplementos do Outlook](manifests.md)
- [Noções básicas sobre conjuntos de requisitos da API do Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md).
- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
