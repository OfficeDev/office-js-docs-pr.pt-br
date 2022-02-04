---
title: Conjuntos de requisitos de Coerção de Imagens
description: 'Suporte para conjuntos de requisitos de Coerção de Imagem com Office de Excel, PowerPoint e Word.'
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# <a name="image-coercion-requirement-sets"></a>Conjuntos de requisitos de Coerção de Imagens

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 permite a conversão para uma imagem (`Office.CoercionType.Image`) ao escrever dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) método. Os aplicativos a seguir são suportados.

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

ImageCoercion 1.2 permite a conversão para o formato SVG (`Office.CoercionType.XmlSvg`) ao escrever dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) método. Os aplicativos a seguir são suportados.

- Excel 2021 e posterior em Windows
- Excel 2021 e posterior no Mac
- PowerPoint 2021 e posterior em Windows
- PowerPoint 2021 e posterior no Mac
- PowerPoint Online
- Word 2021 e posterior no Windows
- Word 2021 e posterior no Mac

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
