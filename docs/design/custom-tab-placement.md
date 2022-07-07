---
title: Posicione uma guia personalizada sobre a faixa de opções
description: Saiba como controlar onde uma guia personalizada aparece na faixa de opções do Office e se ela tem foco por padrão.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 42445898623e082c3c85e756625307dc5a237c28
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659812"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>Posicione uma guia personalizada sobre a faixa de opções

Você pode especificar onde deseja que a guia personalizada do suplemento apareça na faixa de opções do aplicativo do Office usando a marcação no manifesto do suplemento.

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com o artigo [Conceitos básicos para comandos de suplemento](add-in-commands.md). Examine-o se você não tiver feito isso recentemente.

> [!IMPORTANT]
>
> - O recurso de suplemento e a marcação descritos neste artigo só *estão disponíveis PowerPoint na Web*.
> - A marcação descrita neste artigo só funciona em plataformas que dão suporte ao conjunto de requisitos **AddinCommands 1.3**. Veja [o comportamento em plataformas sem suporte](#behavior-on-unsupported-platforms) abaixo.

Especifique onde você deseja que uma guia personalizada apareça identificando qual guia interna do Office você deseja que ela esteja ao lado e especificando se ela deve estar no lado esquerdo ou direito da guia interna. Faça essas especificações incluindo um [elemento InsertBefore](/javascript/api/manifest/customtab#insertbefore) (esquerda) ou [InsertAfter](/javascript/api/manifest/customtab#insertafter) (direita) no elemento [CustomTab](/javascript/api/manifest/customtab) do manifesto do suplemento. (Você não pode ter ambos os elementos.)

No exemplo a seguir, a guia personalizada é configurada para aparecer *logo após a* **guia** Revisão. Observe que o valor do elemento **\<InsertAfter\>** é a ID da guia interna do Office. 

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

Lembre-se dos pontos a seguir.

- O **\<InsertBefore\>** e os **\<InsertAfter\>** elementos são opcionais. Se você não usar nenhum, sua guia personalizada aparecerá como a guia mais à direita na faixa de opções.
- O **\<InsertBefore\>** e **\<InsertAfter\>** os elementos são mutuamente exclusivos. Você não pode usar ambos.
- Se o usuário instalar mais de um suplemento cuja guia personalizada está configurada para o mesmo local, digamos após a guia Revisão,  a guia do suplemento instalado mais recentemente estará localizada nesse local. As guias dos suplementos instalados anteriormente serão movidas para um só lugar. Por exemplo, o usuário instala os suplementos A, B e C nessa ordem e todos estão configurados para inserir uma guia após a guia Revisão, em  seguida, as guias aparecerão nesta ordem: **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.
- Os usuários podem personalizar a faixa de opções no aplicativo do Office. Por exemplo, um usuário pode mover ou ocultar a guia do suplemento. Você não pode impedir isso ou detectar que isso aconteceu.
- Se um usuário mover uma das guias internas, o Office **\<InsertBefore\>** **\<InsertAfter\>** interpretará os elementos e o local padrão da *guia interna*. Por exemplo, se o usuário mover a  guia Revisão para a extremidade direita da faixa de opções, o Office interpretará a marcação no exemplo anterior como "colocar a guia personalizada à direita de onde a guia Revisão estaria por *padrão".*

## <a name="specify-which-tab-has-focus-when-the-document-opens"></a>Especificar qual guia tem foco quando o documento é aberto

O Office sempre dá foco padrão à guia que está imediatamente à direita da **guia** Arquivo. Por padrão, essa é a **guia** Página Inicial. Se você configurar sua guia personalizada para estar antes da  guia Página Inicial, `<InsertBefore>TabHome</InsertBefore>`com , sua guia personalizada terá foco quando o documento for aberto.

> [!IMPORTANT]
> Dar destaque excessivo as inconveniências do seu suplemento e incomodar os usuários e os administradores. Não posicione uma guia personalizada antes da **guia** Página Inicial, a menos que seu suplemento seja a principal maneira como os usuários interagirão com o documento.

## <a name="behavior-on-unsupported-platforms"></a>Comportamento em plataformas sem suporte

Se o suplemento estiver instalado em uma plataforma que não dá suporte ao conjunto de [requisitos AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets), a marcação descrita neste artigo será ignorada e sua guia personalizada aparecerá como a guia mais à direita na faixa de opções. Para impedir que o suplemento seja instalado em plataformas que não dão suporte à marcação, **\<Requirements\>** adicione uma referência ao conjunto de requisitos na seção do manifesto. Para obter instruções, [consulte Especificar quais versões e plataformas do Office podem hospedar seu suplemento](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Como alternativa, projete seu suplemento para ter uma experiência alternativa quando **o AddinCommands 1.3** não tiver suporte, conforme descrito em Design para experiências [alternativas](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences). Por exemplo, se o suplemento contiver instruções que pressupõem que a guia personalizada é onde você deseja, você poderá ter uma versão alternativa que pressuponha que a guia seja a mais à direita.
