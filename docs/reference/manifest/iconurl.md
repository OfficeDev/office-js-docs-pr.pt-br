---
title: Elemento IconUrl no arquivo de manifesto
description: O elemento IconUrl especifica a URL da imagem que representa seu Complemento do Office no UX de inserção e no Office Store.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 68a449b40f6084d26140d59fec61967e163196df
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604635"
---
# <a name="iconurl-element"></a>Elemento IconUrl

Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Pode conter

[Override](override.md)

## <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|DefaultValue|cadeia de caracteres|obrigatório|Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Comentários

Para um complemento de email, o ícone é exibido na interface do usuário Gerenciar arquivos (Outlook) ou Configurações Gerenciar interface do usuário de  >     >  **complementos** (Outlook na Web). Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir** > **Suplementos**. Para todos os tipos de add-in, o ícone também é usado no [AppSource](https://appsource.microsoft.com), se você publicar seu complemento no AppSource.

A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF. Para aplicativos de conteúdo e de painel de tarefas, a imagem especificada deve ter 32 x 32 pixels. Para aplicativos de email, a resolução de imagem deve ser de 64 x 64 pixels. Você também deve especificar um ícone para uso com aplicativos cliente do Office em execução em telas DPI altas usando o [elemento HighResolutionIconUrl.](highresolutioniconurl.md) Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).

Não há suporte para alterar o valor do elemento no tempo de `IconUrl` execução.
