---
title: Posicione uma guia personalizada sobre a faixa de opções
description: Saiba como controlar onde uma guia personalizada aparece na faixa Office faixa de opções e se ela tem foco por padrão.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: bced5bf5506d0366b29d8e2ad6803b0ddfaad80b
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222091"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>Posicione uma guia personalizada sobre a faixa de opções

Você pode especificar onde deseja que a guia personalizada do seu add-in apareça na faixa Office do aplicativo usando a marcação no manifesto do complemento.

> [!NOTE]
> Este artigo supõe que você está familiarizado com o artigo [Conceitos básicos para comandos de complemento.](add-in-commands.md) Revise-o se não tiver feito isso recentemente.

> [!IMPORTANT]
>
> - O recurso de complemento e a marcação descritos neste artigo só estão disponíveis *em PowerPoint na Web*.
> - A marcação descrita neste artigo só funciona em plataformas que suportam o conjunto de **requisitos AddinCommands 1.3**. Consulte [Comportamento em plataformas sem suporte](#behavior-on-unsupported-platforms) abaixo.

Especifique onde você deseja que uma guia personalizada apareça identificando qual guia de Office integrado você deseja que ela seja ao lado e especificando se ela deve estar no lado esquerdo ou direito da guia integrado. Faça essas especificações incluindo um [insertBefore](../reference/manifest/customtab.md#insertbefore) (à esquerda) ou um elemento [InsertAfter](../reference/manifest/customtab.md#insertafter) (à direita) no elemento [CustomTab](../reference/manifest/customtab.md) do manifesto do seu complemento. (Você não pode ter ambos os elementos.)

No exemplo a seguir, a guia personalizada é configurada para aparecer *logo após a* **guia** Revisão. Observe que o valor do **elemento InsertAfter** é a ID da guia Office. 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

Lembre-se dos seguintes pontos.

- Os **elementos InsertBefore** e **InsertAfter** são opcionais. Se você não usar nenhum dos dois, sua guia personalizada aparecerá como a guia mais à direita na faixa de opções.
- Os **elementos InsertBefore** e **InsertAfter** são mutuamente exclusivos. Não é possível usar ambos.
- Se o usuário instalar mais de um add-in cuja guia personalizada está  configurada para o mesmo local, digamos após a guia Revisão, a guia para o complemento instalado mais recentemente estará localizada nesse local. As guias dos complementos instalados anteriormente serão movidas sobre um local. Por exemplo, o usuário instala os complementos A, B e C nessa ordem e  todos são configurados para inserir uma guia após a guia Revisão, em seguida, as guias serão exibidas nesta ordem: **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.
- Os usuários podem personalizar a faixa de opções no Office aplicativo. Por exemplo, um usuário pode mover ou ocultar a guia do seu complemento. Não é possível impedir ou detectar que isso aconteceu.
- Se um usuário mover uma das guias internas, Office interpretar os elementos **InsertBefore** e **InsertAfter** em termos do local padrão da guia *interna*. Por exemplo, se o  usuário mover a guia Revisão para a extremidade direita da faixa de opções, Office interpretará a marcação no exemplo anterior como significando "colocar a guia personalizada à direita de onde a guia Revisão estaria por padrão *."*

## <a name="specify-which-tab-has-focus-when-the-document-opens"></a>Especificar qual guia tem foco quando o documento é aberto

Office sempre dá foco padrão à guia que está imediatamente à direita da **guia Arquivo.** Por padrão, esta é a **guia** Início. Se você configurar sua guia personalizada antes da guia **Página** Inicial, com , sua guia personalizada terá `<InsertBefore>TabHome</InsertBefore>` foco quando o documento for aberto.

> [!IMPORTANT]
> Dar destaque excessivo as inconveniências do seu suplemento e incomodar os usuários e os administradores. Não posicione uma guia personalizada antes da guia **Página** Inicial, a menos que seu complemento seja a principal maneira como os usuários interagirão com o documento.

## <a name="behavior-on-unsupported-platforms"></a>Comportamento em plataformas sem suporte

Se o seu add-in estiver instalado em uma plataforma que não oferece suporte ao conjunto de [requisitos AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), a marcação descrita neste artigo será ignorada e sua guia personalizada aparecerá como a guia mais à direita na faixa de opções. Para impedir que o seu complemento seja instalado em plataformas que não suportam a marcação, adicione uma referência ao conjunto de requisitos na seção Requisitos do manifesto.  Para obter instruções, [consulte Especificar quais Office e plataformas podem hospedar seu complemento](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Como alternativa, projete seu complemento para ter uma experiência alternativa quando **AddinCommands 1.3** não tiver suporte, conforme descrito em Design para experiências [alternativas.](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences) Por exemplo, se o seu add-in contiver instruções que pressuem que a guia personalizada é onde você deseja, você pode ter uma versão alternativa que presume que a guia seja a mais à direita.
