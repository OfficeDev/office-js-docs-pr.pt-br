---
title: Elemento AllowSnapshot no arquivo de manifesto
description: Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8bb143d13a17b3e184af64f1bf18f2a32a55b60c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720957"
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

