---
title: Conjuntos de requisitos de Coerção de Imagens
description: Suporte para conjuntos de requisitos de coerção de imagens com suplementos do Office no Excel, PowerPoint e Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 83817bfc7cf8a193138a805b0e90b4357d605801
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596967"
---
# <a name="image-coercion-requirement-sets"></a>Conjuntos de requisitos de Coerção de Imagens

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1,1 permite a conversão para uma imagem`Office.CoercionType.Image`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) o método. Há suporte para os seguintes hosts:

- Excel 2013 e posterior no Windows
- Excel 2016 e posterior no Mac
- Excel no iPad
- OneNote Online
- PowerPoint 2013 e posterior no Windows
- PowerPoint 2016 e posterior no Mac
- PowerPoint Online
- PowerPoint no iPad
- Word 2013 e posterior no Windows
- Word 2016 e posterior no Mac
- Word Online
- Word no iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1,2 permite conversão para o formato SVG`Office.CoercionType.XmlSvg`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) o método. Há suporte para os seguintes hosts:

- Excel no Windows (conectado a uma assinatura do Office 365)
- Excel no Mac (conectado a uma assinatura do Office 365)
- PowerPoint no Windows (conectado a uma assinatura do Office 365)
- PowerPoint no Mac (conectado a uma assinatura do Office 365)
- PowerPoint Online
- Word no Windows (conectado a uma assinatura do Office 365)
- Word no Mac (conectado a uma assinatura do Office 365)
- Word Online

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
