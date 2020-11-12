---
title: Office UI Fabric em suplementos do Office
description: Obtenha uma visão geral de como usar os componentes do Office UI Fabric em suplementos do Office.
ms.date: 10/29/2020
localization_priority: Normal
ms.openlocfilehash: c4a13c615fe63183f595e24895b9fe6054fdc05d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996372"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric em suplementos do Office

O Office UI Fabric é uma estrutura de front-end JavaScript destinada à criação de experiências de usuário para Office e Office 365. O Fabric fornece componentes com foco em efeitos visuais que você pode estender, reformular e usar no suplemento do Office. Como o Fabric usa a linguagem de design da Microsoft, os componentes da experiência de usuário do Fabric são semelhantes a uma extensão natural do Office.

Se estiver criando um suplemento, recomendamos usar o Office UI Fabric para criar a experiência de usuário. O uso do Office UI Fabric é opcional.

As seções a seguir explicam como começar a usar o Fabric para atender às suas necessidades.

## <a name="use-fabric-core-icons-fonts-colors"></a>Uso do Fabric Core: ícones, fontes, cores

O Fabric Core contém os elementos principais da linguagem de design, como ícones, cores, tipo e grade.  O Fabric Core é independente de estrutura. O Fabric Core é usado pelo Fabric React e incluído nele.

Para começar a usar o Fabric Core:

1. Adicione a referência da CDN ao HTML da sua página.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. Use ícones e fontes do Fabric.

    Para usar um ícone do Fabric, inclua o elemento "i" na sua página e, em seguida, faça referência às classes apropriadas. Para controlar o tamanho do ícone, você pode alterar o tamanho da fonte. Por exemplo, o código a seguir mostra como criar um ícone de tabela muito grande que usa a cor themePrimary (#0078d7).

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Para localizar mais ícones disponíveis no Office UI Fabric, use o recurso de pesquisa na página [Ícones](https://developer.microsoft.com/fabric#/styles/icons). Quando encontrar um ícone para usar no suplemento, não deixe de adicionar um prefixo ao nome do ícone com `ms-Icon--`.

    Para saber mais sobre os tamanhos de fonte e as cores disponíveis no Office UI Fabric, confira [Tipografia](https://developer.microsoft.com/fabric#/styles/typography) e [Cores](https://developer.microsoft.com/fabric#/styles/colors).

## <a name="use-fabric-components"></a>Uso dos componentes do Fabric

O Fabric fornece uma variedade de componentes do UX que você pode usar para criar seu suplemento. Não esperamos que todos os componentes do Fabric sejam usados por um único suplemento. Determinar os melhores componentes para a experiência do seu cenário e do usuário (por exemplo, pode ser difícil exibir um [breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) no painel de tarefas corretamente).

Veja a seguir uma lista de [componentes de UX de reagem de malha](https://developer.microsoft.com/fluentui#/controls/web) comum que recomendamos para uso em um suplemento:

- [Botão](https://developer.microsoft.com/fabric#/components/button)
- [Caixa de seleção](https://developer.microsoft.com/fabric#/components/checkbox)
- [ChoiceGroup](https://developer.microsoft.com/fabric#/components/choicegroup)
- [Lista suspensa](https://developer.microsoft.com/fabric#/components/dropdown)
- [Rótulo](https://developer.microsoft.com/fabric#/components/label)
- [Lista](https://developer.microsoft.com/fabric#/components/list)
- [Tabela dinâmica](https://developer.microsoft.com/fabric#/components/pivot)
- [Campo de texto](https://developer.microsoft.com/fabric#/components/textfield)
- [Alternância](https://developer.microsoft.com/fabric#/components/toggle)

Você pode usar diferentes estruturas do JavaScript, como Angular ou React, para criar o suplemento. Para começar a usar componentes do Fabric com sua estrutura, confira os recursos a seguir.

|**Framework**|**Exemplo**|
|:------------|:----------|
|**React**|[Uso do Office UI Fabric React em suplementos do Office](using-office-ui-fabric-react.md )|
|**Angular**| [Considere quebrar componentes do Fabric com componentes angulares 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
