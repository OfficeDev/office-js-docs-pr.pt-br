---
title: Elemento AllowSnapshot no arquivo de manifesto
description: Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 723817557020f4ec3dbe5b3135877fe49bf67acb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151678"
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
 > **AllowSnapshot** é `true` por padrão. Isso torna uma imagem do complemento visível para os usuários que abrem o documento em uma versão do aplicativo Office que não oferece suporte a complementos Office ou fornece uma imagem estática do add-in se o aplicativo não puder se conectar ao servidor que hospeda o complemento. However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.
