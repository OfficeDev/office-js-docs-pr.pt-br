---
title: Elemento IconUrl no arquivo de manifesto
description: O elemento IconUrl especifica a URL da imagem que representa o suplemento do Office no UX de inserção e na Office Store.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 27001f4109b2dcf93ac71d0a931bb6b4a2b38f2f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292276"
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

Para um suplemento de email, o ícone é exibido na interface do **usuário**  >  **gerenciar suplementos** (Outlook) ou **configurações**  >  **gerenciar suplemento** (Outlook na Web). Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir** > **Suplementos**. Para todos os tipos de suplemento, o ícone também é usado no [AppSource](https://appsource.microsoft.com), se você publicar o suplemento no AppSource.

A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF. Para aplicativos de conteúdo e de painel de tarefas, a imagem especificada deve ter 32 x 32 pixels. Para aplicativos de email, a resolução de imagem recomendada é 64 x 64 pixels. Você também deve especificar um ícone para uso com aplicativos cliente do Office executados em telas de DPI alto usando o elemento [HighResolutionIconUrl](highresolutioniconurl.md) . Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).

`IconUrl`Não há suporte atualmente para a alteração do valor do elemento no tempo de execução.