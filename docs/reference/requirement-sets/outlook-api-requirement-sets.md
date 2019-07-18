---
title: Conjuntos de requisitos de API JavaScript do Outlook
description: ''
ms.date: 06/20/2019
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 1d4c703a9a62e316b2bed9d3fcdd631eb4aec81e
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771888"
---
# <a name="outlook-javascript-api-requirement-sets"></a>Conjuntos de requisitos de API JavaScript do Outlook

Os Suplementos do Outlook declaram quais versões de API exigem usando o elemento Requisitos em seu manifesto. Os suplementos do Outlook sempre incluem um elemento Conjunto com um atributo  definido como  e um atributo  definido como o conjunto de requisitos mínimo de API compatível com os cenários do suplemento.

Por exemplo, o trecho de código de manifesto a seguir indica um conjunto mínimo de requisitos de 1.1.

```xml
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Todas as APIs do Outlook pertencem ao [conjunto de requisitos](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements) `Mailbox`. O conjunto de requisitos `Mailbox` tem versões, e cada novo conjunto de APIs que lançamos pertence a uma versão superior. Nem todos os clientes do Outlook serão compatíveis com o conjunto mais recente de APIs, mas se um cliente do Outlook declarar suporte a um conjunto de requisitos, será compatível com todas as APIs nesse conjunto.

A especificação de uma versão mínima de conjunto de requisitos controla em quais clientes do Outlook o suplemento aparecerá. Se um cliente não oferece suporte para o conjunto de requisitos mínimos, ele não carrega o suplemento. Por exemplo, se for especificada a versão 1.3 do conjunto de requisitos, significa que o suplemento não aparecerá nos clientes do Outlook incompatíveis com a versão 1.3.

## <a name="using-apis-from-later-requirement-sets"></a>Usar APIs de conjuntos de requisitos posteriores

Definir um conjunto de requisitos não limita as APIs disponíveis que o suplemento pode usar. Por exemplo, se o suplemento especificar o conjunto de requisitos 1.1, mas estiver sendo executado em um cliente Outlook que suporte 1.3, o suplemento poderá usar APIs do conjunto de requisitos 1.3.

Para usar uma API mais recente, os desenvolvedores podem verificar se um determinado host oferece suporte ao conjunto de requisitos, fazendo o seguinte.

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.3') === true) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

Como alternativa, os desenvolvedores podem verificar a existência de uma API mais recente usando a técnica JavaScript padrão.

```js
if (item.somePropertyOrFunction !== undefined) {
  // Use item.somePropertyOrFunction.
  item.somePropertyOrFunction;
}
```

Nenhuma verificação desse tipo é necessária para qualquer API que esteja presente na versão do conjunto de requisitos especificada no manifesto.

## <a name="choosing-a-minimum-requirement-set"></a>Escolher um conjunto de requisitos mínimos

Os desenvolvedores devem usar o conjunto de requisitos mínimos que contém o conjunto essencial de APIs para seu cenário, sem o qual o suplemento não funcionará.

## <a name="clients"></a>Clientes

Os clientes a seguir oferecem suporte para suplementos do Outlook.

| Cliente | Conjuntos de requisitos de API com suporte |
| --- | --- |
| Outlook no Windows (conectado à assinatura do Office 365) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 para Windows (compra única) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 para Windows (compra única) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook 2013 para Windows (compra única) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook no Mac (conectado à assinatura do Office 365) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 no Mac (compra única) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2016 no Mac (compra única) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook no iOS | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook no Android | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook na Web (novo) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook na Web (clássico) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Um cliente do Outlook conectado ao Exchange 2019 no local | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Um cliente do Outlook conectado ao Exchange 2016 no local | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |
| Um cliente do Outlook conectado ao Exchange 2013 no local | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |

> [!NOTE]
> O suporte para a versão 1.3 no Outlook 2013 foi adicionado como parte da [atualização para Outlook 2013 de 8 de dezembro de 2015 (KB3114349)](https://support.microsoft.com/kb/3114349). O suporte para a versão 1.4 no Outlook 2013 foi adicionado como parte da [atualização para Outlook 2013 de 13 de setembro de 2016 (KB3118280)](https://support.microsoft.com/help/3118280). O suporte para a versão 1.4 no Outlook 2016 (MSI) foi adicionado como parte da [atualização para Office 2016 de 3 de julho de 2018 (KB4022223)](https://support.microsoft.com/help/4022223).

## <a name="using-preview-apis"></a>Usando APIs de visualização

As novas APIs do JavaScript para Outlook são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários. Para fornecer feedback sobre uma API de visualização, use o mecanismo de feedback no final da página da Web em que a API está documentada.

> [!NOTE]
> As APIs de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção.

Para saber mais detalhes sobre as APIs de visualização, confira o artigo sobre o [conjunto de requisitos da API de visualização do Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md).
