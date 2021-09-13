---
title: Elemento HighResolutionIconUrl no arquivo de manifesto
description: Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store em telas de DPI alto.
ms.date: 03/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 83a86b4a27f2cdfea54b657f70400601207464d3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152108"
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

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|DefaultValue|cadeia de caracteres (URL)|obrigatório|Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Comentários

Para um complemento de email, o ícone é exibido na interface do usuário Gerenciar arquivos de  >  **complementos.** Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir** > **Suplementos**.

A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF. Para aplicativos de conteúdo e painel de tarefas, a resolução de imagem deve ser de 64 x 64 pixels. Para aplicativos de email, a imagem deve ter 128 x 128 pixels. Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
