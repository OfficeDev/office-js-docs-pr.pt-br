---
title: Conjuntos de requisitos da API JavaScript do OneNote
description: ''
ms.date: 03/19/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 287e405955477a98854b1df4a81fe90ec16e5bbc
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871595"
---
# <a name="onenote-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do OneNote

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

A tabela a seguir lista os conjuntos de requisitos do OneNote, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou datas de disponibilidade.

|  Conjunto de requisitos  |  Office Online | 
|:-----|:-----|
| OneNoteApi 1.1  | Setembro de 2016 |  

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="onenote-javascript-api-11"></a>OneNote JavaScript API 1.1 

A OneNote JavaScript API 1.1 é a primeira versão da API. Para obter mais detalhes sobre a API, confira o artigo [Visão geral da programação da API JavaScript do OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).

## <a name="runtime-requirement-support-check"></a>Verificação do suporte a requisitos de tempo de execução

Durante o tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, realizando a seguinte verificação: 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a>Verificação de suporte a requisitos com base em manifesto

Use o elemento Requirements no manifesto do suplemento para especificar conjuntos de requisitos ou membros de API cruciais que o seu suplemento precisa usar. Se o host do Office ou a plataforma não der suporte ao conjunto de requisitos ou membros da API especificados no elemento Requirements, o suplemento não será executado no host ou na plataforma e não será exibido em Meus Suplementos.

O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos host do Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
