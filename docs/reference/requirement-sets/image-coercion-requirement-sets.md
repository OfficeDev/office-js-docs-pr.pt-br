---
title: Conjuntos de requisitos de Coerção de Imagens
description: Suporte para conjuntos de requisitos de coerção de imagem com os complementos do Office no Excel, no PowerPoint e no Word.
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 52ce46a46580500f5a292bf898674d4798378319
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505525"
---
# <a name="image-coercion-requirement-sets"></a>Conjuntos de requisitos de Coerção de Imagens

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 permite a conversão em uma imagem ( ) ao escrever `Office.CoercionType.Image` dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) método. Os seguintes aplicativos são suportados:

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

ImageCoercion 1.2 permite a conversão para o formato SVG ( ) ao escrever `Office.CoercionType.XmlSvg` dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) método. Os seguintes aplicativos são suportados:

- Excel no Windows (conectado a uma assinatura do Microsoft 365)
- Excel no Mac (conectado a uma assinatura do Microsoft 365)
- PowerPoint no Windows (conectado a uma assinatura do Microsoft 365)
- PowerPoint no Mac (conectado a uma assinatura do Microsoft 365)
- PowerPoint Online
- Word no Windows (conectado a uma assinatura do Microsoft 365)
- Word no Mac (conectado a uma assinatura do Microsoft 365)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
