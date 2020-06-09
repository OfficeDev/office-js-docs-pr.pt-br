---
title: Elemento AllowSnapshot no arquivo de manifesto
description: Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c46dcd882592c0b015dae4b9774533b96fe75cfe
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608786"
---
# <a name="allowsnapshot-element"></a>Elemento AllowSnapshot

Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.

**Tipo de suplemento:** Conteúdo

## <a name="syntax"></a>Sintaxe

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Comentários

 > [!IMPORTANT]
 > **AllowSnapshot** é `true` por padrão. Isso cria uma imagem do suplemento visível para os usuários que abrirem o documento em uma versão do aplicativo host que não oferece suporte a Suplementos do Office,ou fornece uma imagem estática do suplemento se o aplicativo host não se conectar ao servidor que hospeda o suplemento. However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.

