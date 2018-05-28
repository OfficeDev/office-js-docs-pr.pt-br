---
title: Office UI Fabric em suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 8fafe8a68c477868c12bff61c7f9ff23fc7314e0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric em suplementos do Office 

O Office UI Fabric ? uma estrutura de front-end JavaScript destinada ? cria??o de experi?ncias de usu?rio para Office e Office 365. O Fabric fornece componentes com foco em efeitos visuais que voc? pode estender, reformular e usar no suplemento do Office. Como o Fabric usa a linguagem de design da Microsoft, os componentes da experi?ncia de usu?rio do Fabric s?o semelhantes a uma extens?o natural do Office. 

Se estiver criando um suplemento, recomendamos usar o Office UI Fabric para criar a experi?ncia de usu?rio. O uso do Office UI Fabric ? opcional.

As se??es a seguir explicam como come?ar a usar o Fabric para atender ?s suas necessidades. 

## <a name="use-fabric-core-icons-fonts-colors"></a>Uso do Fabric Core: ?cones, fontes, cores
O Fabric Core cont?m os elementos principais da linguagem de design, como ?cones, cores, tipo e grade. O Fabric Core ? independente de estrutura. Tanto o Fabric JS como o Fabric React usam o Fabric Core.

Para come?ar a usar o Fabric Core:

1. Adicione a refer?ncia da CDN ao HTML da sua p?gina.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. Use ?cones e fontes do Fabric. 

    Para usar um ?cone do Fabric, inclua o elemento "i" na sua p?gina e, em seguida, fa?a refer?ncia ?s classes apropriadas. Para controlar o tamanho do ?cone, voc? pode alterar o tamanho da fonte. Por exemplo, o c?digo a seguir mostra como criar um ?cone de tabela muito grande que usa a cor themePrimary (#0078d7). 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Para localizar mais ?cones dispon?veis no Office UI Fabric, use o recurso de pesquisa na p?gina [?cones](https://dev.office.com/fabric#/styles/icons). Quando encontrar um ?cone para usar no suplemento, n?o deixe de adicionar um prefixo ao nome do ?cone com `ms-Icon--`. 

    Para saber mais sobre os tamanhos de fonte e as cores dispon?veis no Office UI Fabric, confira [Tipografia](https://dev.office.com/fabric#/styles/typography) e [Cores](https://dev.office.com/fabric#/styles/colors).
 
## <a name="use-fabric-components"></a>Uso dos componentes do Fabric 
O Fabric oferece uma variedade de componentes da experi?ncia do usu?rio que voc? pode usar para criar o suplemento. Alguns desses componentes incluem:

- Componentes de entrada ? por exemplo, bot?o, caixa de sele??o e altern?ncia
- Componentes de navega??o ? por exemplo, din?mico e trilha
- Componentes de notifica??o ? por exemplo, MessageBar e bal?o  

Nem todos os componentes do Fabric s?o recomendados para uso em suplementos. Fornecemos diretrizes sobre como usar os componentes recomendados nesta se??o. Por exemplo, para ver orienta??es de como usar um bot?o do Fabric no suplemento, confira [Bot?o](button.md). 

Voc? pode usar diferentes estruturas do JavaScript, como Angular ou React, para criar o suplemento. Para come?ar a usar componentes do Fabric com sua estrutura, confira os recursos a seguir.

|**Estrutura**|**Exemplo**|
|:------------|:----------|
|**Rea??o**|[Uso do Office UI Fabric React em suplementos do Office](using-office-ui-fabric-react.md )|
|**Angular**| Confira [ngOfficeUIFabric](http://ngofficeuifabric.com/), que ? um projeto comunit?rio com diretivas do Angular 1.5, e [Considere a possibilidade de dispor componentes do Fabric com componentes do Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
