---
title: Conjuntos de requisitos de coerção de imagem
description: Suporte para conjuntos de requisitos de coerção de imagens com suplementos do Office no Excel, PowerPoint e Word.
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: bffe6c074d9e0734299d0087f2488524875931ed
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940840"
---
# <a name="image-coercion-requirement-sets"></a>Conjuntos de requisitos de coerção de imagem

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos de coerção de imagem, os aplicativos host do Office que dão suporte a esse conjunto de requisitos e os números de compilação ou versão para o aplicativo do Office.

## <a name="imagecoercion-11"></a>ImageCoercion 1,1

ImageCoercion 1,1 permite a conversão para uma imagem`Office.CoercionType.Image`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) o método. Há suporte para os seguintes hosts:

- Excel 2013 e posterior no Windows
- Excel 2016 e posterior no Mac
- Excel na Web
- Excel no iPad
- OneNote na Web
- PowerPoint 2013 e posterior no Windows
- PowerPoint 2016 e posterior no Mac
- PowerPoint na Web
- PowerPoint no iPad
- Word 2013 e posterior no Windows
- Word 2016 e posterior no Mac
- Word na Web
- Word no iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1,2

ImageCoercion 1,2 permite conversão para o formato SVG`Office.CoercionType.XmlSvg`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) o método. Há suporte para os seguintes hosts:

- Excel no Windows (conectado a uma assinatura do Office 365)
- Excel no Mac (conectado a uma assinatura do Office 365)
- Excel na Web
- PowerPoint no Windows (conectado a uma assinatura do Office 365)
- PowerPoint no Mac (conectado a uma assinatura do Office 365)
- PowerPoint na Web
- Word no Windows (conectado a uma assinatura do Office 365)
- Word no Mac (conectado a uma assinatura do Office 365)
- Word na Web

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
