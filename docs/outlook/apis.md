---
title: APIs de suplemento do Outlook
description: Saiba como fazer referência a APIs de suplemento do Outlook e declarar permissões em seu suplemento do Outlook.
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5a44d389bb480ec17b73fe445c885c45aff768f7
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074291"
---
# <a name="outlook-add-in-apis"></a>APIs de suplemento do Outlook

Para usar APIs no seu suplemento do Outlook, você deve especificar o local da biblioteca Office.js, o conjunto de requisitos, o esquema e as permissões. Você usará principalmente as APIs javaScript Office expostas por meio do [objeto Mailbox.](#mailbox-object)

## <a name="officejs-library"></a>Biblioteca Office.js

Para interagir com a API do suplemento do Outlook, você precisará usar as APIs JavaScript no Office.js. A rede de distribuição de conteúdo (CDN) para a biblioteca é `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js` . Suplementos enviados ao AppSource devem fazer referência ao Office.js por essa CDN. Eles não podem usar uma referência local.

Referência CDN em um `<script>` marca na `<head>` marca da página da web (arquivo. HTML,. aspx ou. PHP) implementa interface do usuário do seu suplemento.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

À medida que adicionamos novas APIs, a URL para Office.js permanecerá a mesma. Somente mudaremos a versão na URL se mudarmos um comportamento de API existente.

> [!IMPORTANT]
> Ao desenvolver um add-in para qualquer aplicativo cliente Office, consulte a API JavaScript Office de dentro `<head>` da seção da página. Isso garante que a API seja totalmente inicializada antes de qualquer elemento de corpo.

## <a name="requirement-sets"></a>Conjuntos de requisitos

Todas as APIs do Outlook pertencem ao conjunto de requisitos `Mailbox`. O conjunto de requisitos `Mailbox` tem versões, e cada novo conjunto de APIs lançado pertence a uma versão superior. Nem todos os clientes do Outlook terão suporte ao conjunto mais recente de APIs quando for lançado, mas se um cliente do Outlook declarar suporte a um conjunto de requisitos, ele dará suporte a todas as APIs nesse conjunto.

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

Especifique o conjunto de requisitos mínimo que proporciona suporte ao conjunto essencial de APIs para seu cenário, sem o qual os recursos do suplemento não funcionam. Especifique o conjunto de requisitos no manifesto nos elementos `<Requirements>`. Para saber mais, confira os [Manifestos de Suplementos do Outlook](manifests.md) e [Noções básicas sobre os conjuntos de requisitos de APIs do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md).

O elemento `<Methods>` não se aplica a suplementos do Outlook e, portanto, você não pode declarar suporte a métodos específicos.

## <a name="permissions"></a>Permissões

Seu suplemento requer as permissões apropriadas para usar as APIs de que precisa. Há quatro níveis de permissões. Para obter mais detalhes, confira [Noções básicas sobre suplementos do Outlook](understanding-outlook-add-in-permissions.md).

<br/>

|Nível da permissão|Descrição|
|:-----|:-----|
| **Restrito** | Permite o uso de entidades, mas não de expressões regulares. |
| **Leitura de item** | Além do que é permitido em **Restrito**, ele permite:<ul><li>expressões regulares</li><li>acesso de leitura para a API do suplemento do Outlook</li><li>obter as propriedades do item e o token de retorno de chamada</li></ul> |
| **Leitura/gravação** | Além do que é permitido em **Leitura do item**, ele permite:<ul><li>acesso completo à API do Suplemento do Outlook, exceto `makeEwsRequestAsync`</li><li>definição das propriedades do item</li></ul> |
| **Leitura/gravação de caixa de correio** | Além do que é permitido em **Leitura/gravação**, ele permite:<ul><li>criar, ler, gravar itens e pastas</li><li>enviar itens</li><li>chamar [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)</li></ul> |

Em geral, você deve especificar a permissão mínima necessária para o seu suplemento. As permissões são declaradas no elemento `<Permissions>` no manifesto. Para saber mais, confira [Manifestos de suplementos do Outlook](manifests.md). Para obter informações sobre problemas de segurança, consulte Privacidade e [segurança para Office Desadições](../concepts/privacy-and-security.md).

## <a name="mailbox-object"></a>Objeto Mailbox

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>Confira também

- [Manifestos de suplementos do Outlook](manifests.md)
- [Noções básicas sobre conjuntos de requisitos da API do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
