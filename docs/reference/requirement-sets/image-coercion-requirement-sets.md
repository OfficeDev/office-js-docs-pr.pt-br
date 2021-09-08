---
title: Conjuntos de requisitos de Coerção de Imagens
description: Suporte para conjuntos de requisitos de Coerção de Imagem com Office de Excel, PowerPoint e Word.
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 35fed16003fe217e6f1f53d8c790cf78547308cf
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937513"
---
# <a name="image-coercion-requirement-sets"></a>Conjuntos de requisitos de Coerção de Imagens

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 permite a conversão em uma imagem ( ) ao escrever `Office.CoercionType.Image` dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) método. Os aplicativos a seguir são suportados.

- Excel 2013 e posterior em Windows
- Excel 2016 e posterior no Mac
- Excel no iPad
- OneNote Online
- PowerPoint 2013 e posterior em Windows
- PowerPoint 2016 e posterior no Mac
- PowerPoint Online
- PowerPoint no iPad
- Word 2013 e posterior no Windows
- Word 2016 e posterior no Mac
- Word Online
- Word no iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 permite a conversão para o formato SVG ( ) ao escrever `Office.CoercionType.XmlSvg` dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) método. Os aplicativos a seguir são suportados.

- Excel no Windows (conectado a uma assinatura de Microsoft 365)
- Excel no Mac (conectado a uma assinatura de Microsoft 365)
- PowerPoint no Windows (conectado a uma assinatura Microsoft 365 assinatura)
- PowerPoint no Mac (conectado a uma assinatura de Microsoft 365)
- PowerPoint Online
- Word no Windows (conectado a uma assinatura Microsoft 365 de assinatura)
- Word no Mac (conectado a Microsoft 365 assinatura)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
