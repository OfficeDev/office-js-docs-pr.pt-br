---
title: Elemento IconUrl no arquivo de manifesto
description: ''
ms.date: 12/04/2018
ms.openlocfilehash: 471a168b5aa0091292132a1e078fa2b3f5efb448
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433121"
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

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|DefaultValue|cadeia de caracteres|obrigatório|Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Comentários

Para um suplemento de email, o ícone é exibido na interface de usuário **Arquivo**  >  **Gerenciar suplementos** (Outlook) ou na interface de usuário **Configurações**  >  **Gerenciar suplementos** (Outlook Web App). Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir**  >  **Suplementos**. Se você publicar o seu suplemento na Office Store, o ícone também será usado no site da Office Store para todos os tipos de suplementos.

A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF. Para aplicativos de conteúdo e de painel de tarefas, a imagem especificada deve ter 32 x 32 pixels. Para aplicativos de email, a resolução de imagem recomendada é 64 x 64 pixels. Você também deve especificar um ícone para ser usado com aplicativos host do Office executados em telas de DPI alto que utilizam o elemento [HighResolutionIconUrl](highresolutioniconurl.md). Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
