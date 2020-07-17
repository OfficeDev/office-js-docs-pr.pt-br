---
title: Conjuntos de requisitos da API JavaScript do PowerPoint
description: Saiba mais sobre os conjuntos de requisitos da API JavaScript do PowerPoint
ms.date: 03/11/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 8349ff4c96e73ce93a77c86b709081f9355b1f06
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094299"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do PowerPoint

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

A tabela a seguir lista os conjuntos de requisitos do PowerPoint, os aplicativos de host do Office que oferecem suporte a esses conjuntos de requisitos e os versões de compilação ou datas de disponibilidade.

|  Conjunto de requisitos  |  Office no Windows<br>(Conectado à assinatura do Microsoft 365)  |  Office no iPad<br>(Conectado à assinatura do Microsoft 365)  |  Office no Mac<br>(Conectado à assinatura do Microsoft 365)  | Office na Web |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1.1 | Versão 1810 (Build 11001.20074) ou posterior | 2.17 ou posterior | 16.19 ou posterior | Outubro de 2018 |

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre as versões do Office e os números de build, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a>API JavaScript do PowerPoint 1.1

A API JavaScript do PowerPoint 1.1 contém uma única API para criar uma nova apresentação. Para obter detalhes sobre a API, confira [API JavaScript para o PowerPoint](../../powerpoint/powerpoint-add-ins.md).

## <a name="runtime-requirement-support-check"></a>Verificação do suporte a requisitos de tempo de execução

Durante o tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, realizando a seguinte verificação.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Verificação de suporte a requisitos com base em manifesto

Use o `Requirements` elemento no manifesto do suplemento para especificar conjuntos de requisitos ou membros de API cruciais que o seu suplemento precisa usar. Se o host do Office ou a plataforma não der suporte ao conjunto de requisitos ou membros da API especificados no. `Requirements` elemento, o suplemento não será executado no host ou na plataforma e não será exibido em Meus Suplementos.

O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos host do Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

A maior parte da funcionalidade do suplemento do PowerPoint vem do conjunto de APIs comuns. Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do PowerPoint](/javascript/api/powerpoint)
- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
