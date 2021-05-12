---
title: Office UI Fabric em suplementos do Office
description: Obter uma visão geral de como usar os componentes Office UI Fabric em Office de complementos.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 20f926913335197a65ac24e4ec30ed0106b81bae
ms.sourcegitcommit: 132f5082f5bf9500dad0a2eaf89d924c823e575d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330529"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric em suplementos do Office

Office UI Fabric é uma estrutura front-end JavaScript para criar experiências de usuário para Office. O Fabric fornece componentes com foco em efeitos visuais que você pode estender, reformular e usar no suplemento do Office. Como o Fabric usa a linguagem de design da Microsoft, os componentes da experiência de usuário do Fabric são semelhantes a uma extensão natural do Office.

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

O Fabric fornece uma variedade de componentes deux que você pode usar para criar o seu complemento. Não esperamos que todos os componentes de malha sejam usados por um único complemento. Determine os melhores componentes para seu cenário e experiência do usuário (por exemplo, pode ser difícil exibir corretamente um [Breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) no painel de tarefas).

Veja a seguir uma lista de componentes React [UX](https://developer.microsoft.com/fluentui#/controls/web) comuns do Fabric que recomendamos para uso em um complemento:

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
