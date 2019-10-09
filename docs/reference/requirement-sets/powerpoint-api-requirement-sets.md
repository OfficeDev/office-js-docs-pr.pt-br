---
title: Conjuntos de requisitos da API JavaScript do PowerPoint
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 4f64654a4130cc0d4bf96d9c59e364e77c808748
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35941141"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do PowerPoint

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

A tabela a seguir lista os conjuntos de requisitos do PowerPoint, os aplicativos host do Office que dão suporte a esses conjuntos de requisitos e às versões de compilação ou data de disponibilidade.

|  Conjunto de requisitos  |  Office no Windows<br>(conectado à assinatura do Office 365)  |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1,1 | Versão 1810 (Build 11001,20074) ou posterior | 2.17 ou posterior | 16,19 ou posterior | Outubro de 2018 |

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de compilação

Para obter mais informações sobre versões e números de compilação do Office, consulte:

- 
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="powerpoint-javascript-api-11"></a>API JavaScript do PowerPoint 1,1

A API JavaScript do PowerPoint 1,1 contém uma única API para criar uma nova apresentação. Para obter detalhes sobre a API, consulte [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).

## <a name="runtime-requirement-support-check"></a>Verificação do suporte a requisitos de tempo de execução

No tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, fazendo o seguinte.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
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
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API Comum do Office

A maior parte da funcionalidade do suplemento do PowerPoint vem do conjunto de APIs comuns. Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do PowerPoint](/javascript/api/powerpoint)
- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
