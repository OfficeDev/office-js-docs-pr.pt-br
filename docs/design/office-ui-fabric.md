---
title: Office UI Fabric em suplementos do Office 
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7f66a16743da4e7f1e03aeb2c7317d918e941746
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270590"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric em suplementos do Office 

O Office UI Fabric é uma estrutura de front-end JavaScript destinada à criação de experiências de usuário para Office e Office 365. O Fabric fornece componentes com foco em efeitos visuais que você pode estender, reformular e usar no suplemento do Office. Como o Fabric usa a linguagem de design da Microsoft, os componentes da experiência de usuário do Fabric são semelhantes a uma extensão natural do Office. 

Se estiver criando um suplemento, recomendamos usar o Office UI Fabric para criar a experiência de usuário. O uso do Office UI Fabric é opcional.

As seções a seguir explicam como começar a usar o Fabric para atender às suas necessidades. 

## <a name="use-fabric-core-icons-fonts-colors"></a>Uso do Fabric Core: ícones, fontes, cores
O Fabric Core contém os elementos principais da linguagem de design, como ícones, cores, tipo e grade. O Fabric Core é independente de estrutura. O Fabric Core é usado pelo Fabric React e incluído nele.

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
O Fabric oferece uma variedade de componentes da experiência do usuário que você pode usar para criar o suplemento. Alguns desses componentes incluem:

- Componentes de entrada – por exemplo, botão, caixa de seleção e alternância
- Componentes de navegação – por exemplo, dinâmico e trilha
- Componentes de notificação – por exemplo, MessageBar e balão  

Nem todos os componentes do Fabric são recomendados para usar em suplementos. Aqui está uma lista de componentes de experiência de usuário do Fabric React que recomendamos para uso em um suplemento:

- [Navegação estrutural](https://developer.microsoft.com/fabric#/components/breadcrumb)
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
|**Angular**| Confira [ngOfficeUIFabric](http://ngofficeuifabric.com/), que é um projeto comunitário com diretivas do Angular 1.5, e [Considere a possibilidade de dispor componentes do Fabric com componentes do Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
