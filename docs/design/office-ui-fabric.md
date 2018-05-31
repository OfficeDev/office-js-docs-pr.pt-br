---
title: Office UI Fabric em suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 8fafe8a68c477868c12bff61c7f9ff23fc7314e0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437364"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric em suplementos do Office 

O Office UI Fabric é uma estrutura de front-end JavaScript destinada à criação de experiências de usuário para Office e Office 365. O Fabric fornece componentes com foco em efeitos visuais que você pode estender, reformular e usar no suplemento do Office. Como o Fabric usa a linguagem de design da Microsoft, os componentes da experiência de usuário do Fabric são semelhantes a uma extensão natural do Office. 

Se estiver criando um suplemento, recomendamos usar o Office UI Fabric para criar a experiência de usuário. O uso do Office UI Fabric é opcional.

As seções a seguir explicam como começar a usar o Fabric para atender às suas necessidades. 

## <a name="use-fabric-core-icons-fonts-colors"></a>Uso do Fabric Core: ícones, fontes, cores
O Fabric Core contém os elementos principais da linguagem de design, como ícones, cores, tipo e grade. O Fabric Core é independente de estrutura. Tanto o Fabric JS como o Fabric React usam o Fabric Core.

Para começar a usar o Fabric Core:

1. Adicione a referência da CDN ao HTML da sua página.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. Use ícones e fontes do Fabric. 

    Para usar um ícone do Fabric, inclua o elemento "i" na sua página e, em seguida, faça referência às classes apropriadas. Para controlar o tamanho do ícone, você pode alterar o tamanho da fonte. Por exemplo, o código a seguir mostra como criar um ícone de tabela muito grande que usa a cor themePrimary (#0078d7). 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Para localizar mais ícones disponíveis no Office UI Fabric, use o recurso de pesquisa na página [Ícones](https://dev.office.com/fabric#/styles/icons). Quando encontrar um ícone para usar no suplemento, não deixe de adicionar um prefixo ao nome do ícone com `ms-Icon--`. 

    Para saber mais sobre os tamanhos de fonte e as cores disponíveis no Office UI Fabric, confira [Tipografia](https://dev.office.com/fabric#/styles/typography) e [Cores](https://dev.office.com/fabric#/styles/colors).
 
## <a name="use-fabric-components"></a>Uso dos componentes do Fabric 
O Fabric oferece uma variedade de componentes da experiência do usuário que você pode usar para criar o suplemento. Alguns desses componentes incluem:

- Componentes de entrada – por exemplo, botão, caixa de seleção e alternância
- Componentes de navegação – por exemplo, dinâmico e trilha
- Componentes de notificação – por exemplo, MessageBar e balão  

Nem todos os componentes do Fabric são recomendados para uso em suplementos. Fornecemos diretrizes sobre como usar os componentes recomendados nesta seção. Por exemplo, para ver orientações de como usar um botão do Fabric no suplemento, confira [Botão](button.md). 

Você pode usar diferentes estruturas do JavaScript, como Angular ou React, para criar o suplemento. Para começar a usar componentes do Fabric com sua estrutura, confira os recursos a seguir.

|**Estrutura**|**Exemplo**|
|:------------|:----------|
|**Reação**|[Uso do Office UI Fabric React em suplementos do Office](using-office-ui-fabric-react.md )|
|**Angular**| Confira [ngOfficeUIFabric](http://ngofficeuifabric.com/), que é um projeto comunitário com diretivas do Angular 1.5, e [Considere a possibilidade de dispor componentes do Fabric com componentes do Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
