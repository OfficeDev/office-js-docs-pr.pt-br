---
title: Posicionar uma guia personalizada na faixa de opções
description: Saiba como controlar onde uma guia personalizada aparece na faixa de opções do Office e se tem o foco por padrão.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 2c1e2ae66805212e78868cf7c07a0e5c14cd4025
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088163"
---
# <a name="position-a-custom-tab-on-the-ribbon-preview"></a>Posicionar uma guia personalizada na faixa de opções (visualização)

Você pode especificar onde deseja que a guia personalizada do seu suplemento apareça na faixa de opções do aplicativo do Office usando a marcação no manifesto do suplemento.

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com os [conceitos básicos do artigo para comandos de suplemento](add-in-commands.md). Verifique se você não fez isso recentemente.

> [!IMPORTANT]
>
> - O recurso de suplemento e a marcação descritos neste artigo estão em visualização e só estão *disponíveis no PowerPoint na Web*. Recomendamos que você experimente a marcação apenas em ambientes de teste e desenvolvimento. Não use a marcação de visualização em um ambiente de produção ou em documentos de negócios críticos.
> - A marcação descrita neste artigo funciona somente em plataformas que dão suporte ao conjunto de requisitos **AddinCommands 1,3**. Veja o [comportamento de plataformas não suportadas](#behavior-on-unsupported-platforms) abaixo.

Especifique onde você deseja que uma guia personalizada seja exibida identificando a guia do Office interna que você deseja que ele fique próximo e especificando se deve estar no lado esquerdo ou direito da guia interna. Faça essas especificações incluindo um elemento [InsertBefore](../reference/manifest/customtab.md#insertbefore) (à esquerda) ou [InsertAfter](../reference/manifest/customtab.md#insertafter) (à direita) no elemento [CustomTab](../reference/manifest/customtab.md) do manifesto do seu suplemento. (Não é possível ter ambos os elementos.)

No exemplo a seguir, a guia personalizada é configurada para aparecer *imediatamente após* a guia **revisão** . Observe que o valor do `<InsertAfter>` elemento é a ID da guia interna do Office. 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

Tenha em mente os seguintes pontos.

- Os  `<InsertBefore>`  `<InsertAfter>` elementos e são opcionais. Se você não usar nenhuma, a guia personalizada será exibida como a guia mais à direita na faixa de opções.
- Os  `<InsertBefore>`  `<InsertAfter>` elementos e são mutuamente exclusivos. Você não pode usar ambos.
- Se o usuário instalar mais de um suplemento cuja guia personalizada é configurada para o mesmo local, diga após a guia **revisão** e, em seguida, a guia para o suplemento instalado mais recentemente estará localizada nesse local. As guias dos suplementos instalados anteriormente serão movidas em um só lugar. Por exemplo, o usuário instala os suplementos A, B e C nessa ordem, e todos estão configurados para inserir uma Tabulação após a guia **revisão** e, em seguida, as guias serão exibidas nesta ordem: **revisar**, **AddinCTab**, **AddinBTab**, **AddinATab**.
- Os usuários podem personalizar a faixa de opções no aplicativo do Office. Por exemplo, um usuário pode mover ou ocultar a guia do seu suplemento. Você não pode impedir isso nem detectar que aconteceu.
- Se um usuário mover uma das guias internas, o Office interpretará os `<InsertBefore>`  `<InsertAfter>` elementos e em termos do *local padrão da guia interna*. Por exemplo, se o usuário mover a guia **revisão** para a extremidade direita da faixa de opções, o Office interpretará a marcação no exemplo acima como significado "Coloque a guia personalizada à direita de *onde a guia **revisão** seria por padrão*".

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a>Especificar qual guia tem o foco quando o documento é aberto

O Office sempre dá o foco padrão à guia imediatamente à direita da guia **arquivo** . Por padrão, essa é a guia **página inicial** . Se você configurar sua guia personalizada para que seja antes da guia **página inicial** , com `<InsertBefore>TabHome</InsertBefore>` , sua guia personalizada terá o foco quando o documento for aberto.

> [!IMPORTANT]
> Fornecer excesso de importância para o seu suplemento inconvenientes e incomodar usuários e administradores. Não posicione uma Tabulação personalizada antes da guia **página inicial** , a menos que o suplemento seja o principal modo como os usuários irão interagir com o documento.

## <a name="behavior-on-unsupported-platforms"></a>Comportamento de plataformas sem suporte

Se seu suplemento estiver instalado em uma plataforma que não ofereça suporte ao [conjunto de requisitos AddinCommands 1,3](../reference/requirement-sets/add-in-commands-requirement-sets.md), a marcação descrita neste artigo será ignorada e sua guia personalizada aparecerá como a guia mais à direita na faixa de opções. Para impedir que o suplemento seja instalado em plataformas que não suportam a marcação, adicione uma referência ao conjunto de requisitos na `<Requirements>` seção do manifesto. Para obter instruções, consulte [definir o elemento requirements no manifesto](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest). Como alternativa, você pode criar seu suplemento para ter uma experiência alternativa quando o **AddinCommands 1,3** não é suportado, conforme descrito em [usar verificações de tempo de execução no código JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). Por exemplo, se o suplemento contiver instruções que presumim que a guia personalizada esteja onde você deseja, você poderia ter uma versão alternativa que assume que a guia é a direita.
