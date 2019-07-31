---
title: Conjuntos de requisitos da API JavaScript do OneNote
description: ''
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: e1012b337b3713f57a5d3df7f7c7ccbcf509b5aa
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940841"
---
# <a name="onenote-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do OneNote

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

A tabela a seguir lista os conjuntos de requisitos do OneNote, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou datas de disponibilidade.

|  Conjunto de requisitos  |  Office na Web |
|:-----|:-----|
| OneNoteApi 1.1  | Setembro de 2016 |

## <a name="onenote-javascript-api-11"></a>OneNote JavaScript API 1.1

A OneNote JavaScript API 1.1 é a primeira versão da API. Para obter mais detalhes sobre a API, confira o artigo [Visão geral da programação da API JavaScript do OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).

## <a name="runtime-requirement-support-check"></a>Verificação do suporte a requisitos de tempo de execução

No tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, fazendo o seguinte.

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1') === true) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Verificação de suporte a requisitos com base em manifesto

Use o `Requirements` elemento no manifesto do suplemento para especificar conjuntos de requisitos críticos ou membros da API que seu suplemento deve usar. Se o host ou a plataforma do Office não oferecer suporte aos conjuntos de requisitos ou membros `Requirements` de API especificados no elemento, o suplemento não será executado nesse host ou plataforma e não será exibido em meus suplementos.

O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos host do Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do OneNote](/javascript/api/onenote)
- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
