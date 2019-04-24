---
title: Elemento HighResolutionIconUrl no arquivo de manifesto
description: ''
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 5264fc969bda30a9b2212996800b984533a3188c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452085"
---
# <a name="highresolutioniconurl-element"></a>Elemento HighResolutionIconUrl

Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store em telas de DPI alto.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Pode conter

[Override](override.md)

## <a name="attributes"></a>Atributos

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|DefaultValue|cadeia de caracteres (URL)|obrigatório|Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Comentários

Para um suplemento de email, o ícone é exibido na interface de usuário **Arquivo**  >  **Gerenciar suplementos**. Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir**  >  **Suplementos**.

A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF. Para aplicativos do painel de tarefas e de conteúdo, a resolução de imagem recomendada é 64 x 64 pixels. Para aplicativos de email, a imagem deve ter 128 x 128 pixels. Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
